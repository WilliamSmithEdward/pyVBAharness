"""Property-based testing of VBA functions via Hypothesis.

Strategies are derived from the parsed VBA signature (vbasig types), so

    check_vba_function(session, source, "Add",
                       check=lambda args, value: value == sum(args))

fuzzes the function with type-appropriate inputs and shrinks any failure to
a minimal counterexample. Hypothesis is an optional dependency:
``pip install pyvbaharness[fuzz]``.

Session recovery composes automatically: if a generated input hangs the run,
the session recycles, the injection cache reinjects the module on the next
example, and Hypothesis shrinks toward the hanging input like any other
failure.
"""
from __future__ import annotations

from typing import Any, Callable

from . import codegen, vbasig
from .results import PASSED, HarnessError
from .session import ExcelSession

_FUZZ_MODULE = "PyVbaFuzz"

# VBA integer limits; Double avoids NaN/Inf, which cannot cross the JSON
# pipe and are unrepresentable in most VBA arithmetic anyway.
_INT_LIMITS = {
    "byte": (0, 255),
    "integer": (-32768, 32767),
    "long": (-2147483648, 2147483647),
    "longlong": (-2**63, 2**63 - 1),
}


def _require_hypothesis():
    try:
        from hypothesis import strategies
        return strategies
    except ImportError as err:  # pragma: no cover - environment specific
        raise HarnessError(
            "Property-based testing needs Hypothesis: "
            "pip install pyvbaharness[fuzz]") from err


def strategy_for(parameter: vbasig.Parameter):
    """Hypothesis strategy for one declared VBA parameter."""
    strategies = _require_hypothesis()
    type_name = parameter.type_name.lower()
    if parameter.is_array or parameter.param_array:
        raise HarnessError(
            f"Parameter {parameter.name!r} is an array; array fuzzing is "
            "not supported.")
    if type_name in _INT_LIMITS:
        low, high = _INT_LIMITS[type_name]
        return strategies.integers(min_value=low, max_value=high)
    if type_name in ("double", "single"):
        width = 32 if type_name == "single" else 64
        return strategies.floats(allow_nan=False, allow_infinity=False,
                                 width=width)
    if type_name == "currency":
        return strategies.floats(min_value=-1e12, max_value=1e12,
                                 allow_nan=False, allow_infinity=False)
    if type_name == "boolean":
        return strategies.booleans()
    if type_name == "string":
        return strategies.text(
            alphabet=strategies.characters(min_codepoint=32,
                                           max_codepoint=0xFFFD),
            max_size=100)
    if type_name == "variant":
        return strategies.one_of(
            strategies.integers(min_value=-2147483648,
                                max_value=2147483647),
            strategies.floats(allow_nan=False, allow_infinity=False),
            strategies.text(max_size=50),
            strategies.booleans(),
        )
    raise HarnessError(
        f"Parameter {parameter.name!r} has unsupported type "
        f"{parameter.type_name!r} for fuzzing (supported: integer types, "
        "Double, Single, Currency, Boolean, String, Variant).")


def check_vba_function(session: ExcelSession, source: str, proc: str,
                       check: Callable[[tuple, Any], bool] | None = None,
                       max_examples: int = 100,
                       run_timeout: float = 30.0,
                       module_name: str = _FUZZ_MODULE) -> None:
    """Fuzz a VBA function; raises with a shrunk counterexample on failure.

    A run is a failure when it does not pass (VBA error, timeout, blocked
    modal) or when ``check(args, value)`` returns falsy or raises. On
    success, returns None.
    """
    strategies = _require_hypothesis()
    from hypothesis import given, settings

    codegen.validate_module_name(module_name)
    signature = vbasig.find_procedure(source, proc)
    if signature is None:
        raise HarnessError(
            f"{proc!r} is not declared in the provided source.")
    fuzzed = [p for p in signature.params if not p.optional]
    argument_strategies = tuple(strategy_for(p) for p in fuzzed)

    def run_one(args: tuple) -> None:
        # Reinjection is a cache no-op normally, and exactly what is needed
        # right after a hang recycled the session.
        if not session.has_workbook:
            session.new_workbook()
        session.add_module(module_name, source, line_numbers=True)
        result = session.run_macro(f"{module_name}.{proc}", *args,
                                   timeout=run_timeout)
        if result.outcome != PASSED:
            detail = str(result.error) if result.error else result.message
            raise AssertionError(
                f"{proc}{args!r} -> {result.outcome}: {detail}")
        if check is not None:
            verdict = check(args, result.value)
            if verdict is False:
                raise AssertionError(
                    f"check rejected {proc}{args!r} -> {result.value!r}")

    wrapped = given(strategies.tuples(*argument_strategies))(run_one)
    configured = settings(max_examples=max_examples, deadline=None,
                          database=None)(wrapped)
    configured()
