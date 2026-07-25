"""VBA source instrumentation for error-line capture.

VBA reports the line of a runtime error only through ``Erl``, and ``Erl``
only works when the executing code carries classic BASIC line numbers.
Numbering alone is not enough, though: ``Erl`` is per-procedure error
context. When an error propagates out of the erroring procedure into the
harness dispatcher's handler, ``Erl`` there reports the dispatcher's own
(unnumbered) lines, i.e. 0 - verified live on 2026-07-25. This is exactly
why per-procedure handler injection is the established practice in VBA
tooling.

So the instrumentation does two things per procedure:

1. numbers each executable statement with its 1-based physical line index in
   the ORIGINAL source (``add_line_numbers``), and
2. installs a minimal handler (``instrument_error_lines``) that notes
   ``Erl`` and re-raises:

       Public Sub Main()
           On Error GoTo PyVbaErl__
       3   x = 1
           ...
           Exit Sub
       PyVbaErl__:
           PyVbaFail "Mod.Main", Erl, Err.Number, Err.Source, Err.Description
       End Sub

``PyVbaFail`` (support module) appends a stack frame per handler that fires
(deepest first, so the chain of a propagating error becomes a VBA stack
trace), keeps the FIRST noted line as the error's origin line, and
re-raises, so the error still propagates to the dispatcher with its number,
source, and description intact. The inserted
lines carry no numbers themselves (a number on the handler's call line would
overwrite ``Erl`` before it is read) and declare no locals (no name
collisions with user code). A user procedure's own ``On Error`` statements
simply override the injected one from that point on, preserving the user's
error semantics.

The transformer is deliberately conservative: a line it is not sure about is
left unnumbered, which only means an error there reports the previous
numbered line. Lines never numbered:

- anything outside a Sub/Function/Property body
- declaration and End lines, and single-line procedures
- continuation lines (a number is a label; it may only start a statement)
- blank, comment, Rem, and preprocessor (#If/#Const) lines
- existing labels and existing line numbers
- structure keywords that cannot carry a label (Case, Else, ElseIf, End,
  Next, Loop, Wend) and declaration statements (Dim, Const, Static)

If the body already contains any numeric line label, or the source already
mentions the handler label, the source is returned unchanged: mixing labels
risks duplicates, which is a VBA compile error.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .vbasig import strip_comment

HANDLER_LABEL = "PyVbaErl__"

_DECL_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*"
    r"(?P<kind>Sub|Function|Property)"
    r"(?:[ \t]+(?:Get|Let|Set))?[ \t]+(?P<name>[A-Za-z]\w*)",
    re.IGNORECASE,
)
_END_PROC_RE = re.compile(r"^[ \t]*End[ \t]+(?:Sub|Function|Property)\b",
                          re.IGNORECASE)
_INLINE_END_PROC_RE = re.compile(r"\bEnd[ \t]+(?:Sub|Function|Property)\b",
                                 re.IGNORECASE)
_LABEL_RE = re.compile(r"^[ \t]*[A-Za-z_]\w*[ \t]*:")
_NUMBERED_RE = re.compile(r"^[ \t]*\d")

# First words that must not (or should not) carry a line-number label.
_SKIP_FIRST_WORDS = {
    "case", "else", "elseif", "end", "next", "loop", "wend",
    "option", "dim", "const", "static", "rem", "attribute",
    "implements", "declare", "type", "enum",
    "sub", "function", "property", "public", "private", "friend",
}

_FIRST_WORD_RE = re.compile(r"^[ \t]*([A-Za-z_#]\w*)")

# Statements that open a multi-line block must be the FIRST statement on
# their line: "PyVbaCovHit 1, 4: If flag Then" turns a block If into a
# single-line If, and the matching End If then fails to compile (observed
# live as a modal compile-error dialog). Such lines are still numbered, so
# Erl mapping is unaffected; they simply carry no coverage hit and are not
# counted as coverable, because their execution cannot be observed.
_BLOCK_OPENERS = {"if", "for", "do", "while", "select", "with"}


def _is_continued(line: str) -> bool:
    stripped = strip_comment(line).rstrip()
    return stripped.endswith("_") and (
        len(stripped) == 1 or stripped[-2] in " \t")


def _numberable(line: str) -> bool:
    stripped = line.strip()
    if not stripped or stripped.startswith("'") or stripped.startswith("#"):
        return False
    if _NUMBERED_RE.match(line) or _LABEL_RE.match(line):
        return False
    match = _FIRST_WORD_RE.match(line)
    if match is None:
        return False
    return match.group(1).lower() not in _SKIP_FIRST_WORDS


def add_line_numbers(source: str) -> str:
    """Return the source with executable body lines numbered, or unchanged
    when numbering would be unsafe (existing numeric labels)."""
    lines = source.splitlines()
    if len(lines) > 60000:
        # Classic line-number labels are limited; a source this large is not
        # a harness snippet anyway.
        return source

    in_body = False
    continued = False
    for line in lines:
        was_continued = continued
        continued = _is_continued(line)
        if was_continued:
            continue
        if not in_body and _DECL_RE.match(line):
            in_body = not _INLINE_END_PROC_RE.search(
                line[_DECL_RE.match(line).end():])
            continue
        if in_body and _END_PROC_RE.match(line):
            in_body = False
            continue
        if in_body and _NUMBERED_RE.match(line):
            return source  # existing numbers: do not mix

    out: list[str] = []
    in_body = False
    continued = False
    for index, line in enumerate(lines, start=1):
        was_continued = continued
        continued = _is_continued(line)
        if was_continued:
            out.append(line)
            continue
        if not in_body:
            if _DECL_RE.match(line):
                in_body = not _INLINE_END_PROC_RE.search(
                    line[_DECL_RE.match(line).end():])
            out.append(line)
            continue
        if _END_PROC_RE.match(line):
            in_body = False
            out.append(line)
            continue
        if _numberable(line):
            out.append(f"{index} {line.lstrip()}")
        else:
            out.append(line)
    return "\r\n".join(out) + ("\r\n" if source.endswith(("\n", "\r")) else "")


@dataclass
class InstrumentedModule:
    """Instrumented source plus the lines the coverage hooks can observe."""

    source: str
    coverable_lines: list[int]

    @property
    def changed(self) -> bool:
        return bool(self.coverable_lines)


def instrument_module(source: str, module_name: str = "",
                      coverage_id: int | None = None) -> InstrumentedModule:
    """Number executable lines, install per-procedure Erl capture, and
    (optionally) insert coverage hits.

    One pass produces all three because they interlock: numbers make ``Erl``
    meaningful, the per-procedure handler reports frames as
    ``module_name.Proc``, and a coverage hit rides the numbered line as a
    colon-compound statement (``42 PyVbaCovHit 3, 42: x = 1``) so the line
    count, and therefore the Erl-to-source mapping, never shifts.

    Returns the source unchanged (empty coverable_lines) when
    instrumentation would be unsafe: existing numeric labels or the handler
    label already present.
    """
    if HANDLER_LABEL in source or _has_existing_numbers(source):
        return InstrumentedModule(source, [])
    lines = source.splitlines()
    if len(lines) > 60000:
        return InstrumentedModule(source, [])

    hit_prefix = (f"PyVbaCovHit {coverage_id}, "
                  if coverage_id is not None else None)
    out: list[str] = []
    coverable: list[int] = []
    in_body = False
    continued = False
    decl_pending = False
    proc_kind = ""
    proc_name = ""
    for index, line in enumerate(lines, start=1):
        was_continued = continued
        continued = _is_continued(line)
        if was_continued:
            out.append(line)
            if decl_pending and not continued:
                out.append(f"    On Error GoTo {HANDLER_LABEL}")
                decl_pending = False
            continue
        if not in_body:
            match = _DECL_RE.match(line)
            if match and not _INLINE_END_PROC_RE.search(line[match.end():]):
                in_body = True
                proc_kind = match.group("kind").capitalize()
                proc_name = match.group("name")
                out.append(line)
                if continued:
                    decl_pending = True  # inject after the decl finishes
                else:
                    out.append(f"    On Error GoTo {HANDLER_LABEL}")
                continue
            out.append(line)
            continue
        if _END_PROC_RE.match(line):
            in_body = False
            frame = (f"{module_name}.{proc_name}" if module_name
                     else proc_name)
            out.append(f"    Exit {proc_kind}")
            out.append(f"{HANDLER_LABEL}:")
            out.append(f'    PyVbaFail "{frame}", Erl, Err.Number, '
                       "Err.Source, Err.Description")
            out.append(line)
            continue
        if _numberable(line):
            body = line.lstrip()
            first_word = _FIRST_WORD_RE.match(line)
            opens_block = (first_word is not None
                           and first_word.group(1).lower() in _BLOCK_OPENERS)
            if hit_prefix is not None and not opens_block:
                coverable.append(index)
                out.append(f"{index} {hit_prefix}{index}: {body}")
            else:
                if hit_prefix is None:
                    coverable.append(index)
                out.append(f"{index} {body}")
        else:
            out.append(line)
    text = "\r\n".join(out) + ("\r\n" if source.endswith(("\n", "\r"))
                               else "")
    return InstrumentedModule(text, coverable)


def instrument_error_lines(source: str, module_name: str = "") -> str:
    """Back-compatible wrapper: numbering plus handlers, no coverage."""
    return instrument_module(source, module_name=module_name).source


def _has_existing_numbers(source: str) -> bool:
    in_body = False
    continued = False
    for line in source.splitlines():
        was_continued = continued
        continued = _is_continued(line)
        if was_continued:
            continue
        if not in_body:
            match = _DECL_RE.match(line)
            if match and not _INLINE_END_PROC_RE.search(line[match.end():]):
                in_body = True
            continue
        if _END_PROC_RE.match(line):
            in_body = False
            continue
        if _NUMBERED_RE.match(line):
            return True
    return False
