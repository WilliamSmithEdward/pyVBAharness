"""VBA procedure-signature parsing.

The harness must call a target procedure directly rather than through
``Application.Run``: an error raised inside an ``Application.Run`` callee
crosses a COM boundary, so the calling VBA procedure's ``On Error`` handler
never sees it and Excel raises its runtime-error dialog instead (observed
live, 2026-07-25). Direct calls unwind normally and are trapped.

A direct call has to match the callee's shape exactly (Sub versus Function,
and argument count), so the harness parses the declaration before generating
the dispatcher. Getting this wrong produces a VBA compile error at run time,
which is why arity is validated in Python first.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

KIND_SUB = "sub"
KIND_FUNCTION = "function"
KIND_PROPERTY_GET = "property-get"

_DECL_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*"
    r"(?:Declare[ \t]+(?:PtrSafe[ \t]+)?)?"
    r"(?P<kind>Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+"
    r"(?P<name>[A-Za-z][A-Za-z0-9_]*)[ \t]*(?P<rest>.*)$",
    re.IGNORECASE,
)


@dataclass
class ProcedureSignature:
    name: str
    kind: str
    required: int
    optional: int
    has_param_array: bool

    @property
    def returns_value(self) -> bool:
        return self.kind in (KIND_FUNCTION, KIND_PROPERTY_GET)

    def accepts(self, arg_count: int) -> bool:
        if arg_count < self.required:
            return False
        if self.has_param_array:
            return True
        return arg_count <= self.required + self.optional

    def arity_text(self) -> str:
        if self.has_param_array:
            return f"{self.required} or more"
        if self.optional:
            return f"{self.required} to {self.required + self.optional}"
        return str(self.required)


def strip_comment(line: str) -> str:
    """Drop a trailing apostrophe comment, honoring string literals."""
    in_string = False
    for index, char in enumerate(line):
        if char == '"':
            in_string = not in_string
        elif char == "'" and not in_string:
            return line[:index]
    return line


def logical_lines(source: str) -> list[str]:
    """Join VBA line continuations into single logical lines."""
    joined: list[str] = []
    buffer = ""
    for raw in source.splitlines():
        line = strip_comment(raw).rstrip()
        if line.endswith(" _"):
            buffer += line[:-1]
            continue
        buffer += line
        joined.append(buffer)
        buffer = ""
    if buffer:
        joined.append(buffer)
    return joined


def split_top_level(text: str) -> list[str]:
    """Split a parameter list on commas outside parentheses and strings."""
    parts: list[str] = []
    depth = 0
    in_string = False
    current = ""
    for char in text:
        if char == '"':
            in_string = not in_string
        if not in_string:
            if char in "([":
                depth += 1
            elif char in ")]":
                depth -= 1
            elif char == "," and depth == 0:
                parts.append(current)
                current = ""
                continue
        current += char
    if current.strip():
        parts.append(current)
    return [p.strip() for p in parts if p.strip()]


def _param_list(rest: str) -> str | None:
    """Extract the text inside the declaration's parentheses, or None."""
    rest = rest.lstrip()
    if not rest.startswith("("):
        return None
    depth = 0
    in_string = False
    for index, char in enumerate(rest):
        if char == '"':
            in_string = not in_string
        if in_string:
            continue
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return rest[1:index]
    return None


def parse_declaration(line: str) -> ProcedureSignature | None:
    match = _DECL_RE.match(line)
    if not match:
        return None
    raw_kind = " ".join(match.group("kind").split()).lower()
    if raw_kind == "sub":
        kind = KIND_SUB
    elif raw_kind == "function":
        kind = KIND_FUNCTION
    elif raw_kind == "property get":
        kind = KIND_PROPERTY_GET
    else:
        # Property Let / Set cannot be invoked as a call expression.
        return None

    params_text = _param_list(match.group("rest"))
    required = optional = 0
    has_param_array = False
    for param in split_top_level(params_text or ""):
        head = param.lower()
        if head.startswith("paramarray"):
            has_param_array = True
        elif head.startswith("optional"):
            optional += 1
        else:
            required += 1
    return ProcedureSignature(name=match.group("name"), kind=kind,
                              required=required, optional=optional,
                              has_param_array=has_param_array)


def find_procedure(source: str, name: str) -> ProcedureSignature | None:
    """First callable declaration of ``name`` in a module's source."""
    wanted = name.lower()
    for line in logical_lines(source):
        signature = parse_declaration(line)
        if signature is not None and signature.name.lower() == wanted:
            return signature
    return None


def list_procedures(source: str) -> list[ProcedureSignature]:
    found: list[ProcedureSignature] = []
    for line in logical_lines(source):
        signature = parse_declaration(line)
        if signature is not None:
            found.append(signature)
    return found


def discover_tests(source: str, prefix: str = "Test") -> list[str]:
    """Names of zero-required-argument Subs/Functions matching the prefix.

    Pure-Python test discovery, shared by ExcelSession.run_tests and
    SessionPool sharding so both see the same test list without a COM round
    trip.
    """
    wanted = prefix.lower()
    return [
        signature.name
        for signature in list_procedures(source)
        if signature.kind in (KIND_SUB, KIND_FUNCTION)
        and signature.required == 0
        and signature.name.lower().startswith(wanted)
    ]
