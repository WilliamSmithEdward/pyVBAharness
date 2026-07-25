"""Result types and exceptions.

Outcomes are literal (RIDM: status must be literal). A timeout or a blocked
modal is infrastructure state, never evidence about the VBA code itself; only
``passed`` and ``vba-error`` describe the code under test.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

PASSED = "passed"
VBA_ERROR = "vba-error"
TIMEOUT = "timeout"
MODAL_BLOCKED = "modal-blocked"
RUNNER_ERROR = "runner-error"

RUN_OUTCOMES = (PASSED, VBA_ERROR, TIMEOUT, MODAL_BLOCKED, RUNNER_ERROR)

COMPILE_ACCEPTED = "accepted"
COMPILE_REJECTED = "rejected"
COMPILE_INFRA_FAILURE = "infrastructure-failure"


@dataclass
class VbaError:
    """A VBA runtime error captured inside VBA by the injected runner.

    ``line`` is the 1-based line in the ORIGINAL injected source when line
    numbering was applied (run_vba does this by default); None when the
    source was not numbered or no numbered line had executed yet. It reports
    the last executed numbered line, so a line the numberer skipped
    attributes to the nearest numbered line above it.

    ``stack`` lists (procedure, line) frames deepest-first: each
    per-procedure handler that fired while the error propagated contributed
    one frame, so a nested failure reads as a real VBA stack trace.
    """

    number: int
    source: str
    description: str
    line: int | None = None
    stack: list[tuple[str, int]] = field(default_factory=list)

    def __str__(self) -> str:
        prefix = f"VBA error {self.number}"
        if self.source:
            prefix += f" from {self.source}"
        if self.line is not None:
            prefix += f" at line {self.line}"
        text = f"{prefix}: {self.description}"
        if self.stack:
            frames = " <- ".join(f"{proc}:{line}"
                                 for proc, line in self.stack)
            text += f" [stack: {frames}]"
        return text


@dataclass
class DialogRecord:
    """One captured Excel/VBE dialog, as observed by the watcher."""

    title: str
    message: str
    texts: list[str] = field(default_factory=list)
    buttons: list[str] = field(default_factory=list)
    button_ids: list[int] = field(default_factory=list)
    classification: str = "excel-modal"
    action: str = "none"  # none | click:<caption> | blocked:<reason>
    screenshot: str = ""  # PNG path captured before dismissal/kill, if any


@dataclass
class RunResult:
    """Outcome of one managed VBA run."""

    outcome: str
    duration_s: float
    value: Any = None
    output: list[str] = field(default_factory=list)
    error: VbaError | None = None
    message: str = ""
    dialogs: list[DialogRecord] = field(default_factory=list)
    screenshot: str = ""  # PNG of the Excel window on timeout, if captured

    @property
    def ok(self) -> bool:
        return self.outcome == PASSED


@dataclass
class ModuleCoverage:
    """Line coverage for one instrumented module."""

    module: str
    coverable: list[int]
    hit: list[int]

    @property
    def missed(self) -> list[int]:
        hit_set = set(self.hit)
        return [line for line in self.coverable if line not in hit_set]

    @property
    def percent(self) -> float:
        if not self.coverable:
            return 100.0
        return 100.0 * len(self.hit) / len(self.coverable)


@dataclass
class CoverageReport:
    """Aggregated line coverage across instrumented modules."""

    modules: dict[str, ModuleCoverage]

    @property
    def percent(self) -> float:
        coverable = sum(len(m.coverable) for m in self.modules.values())
        if not coverable:
            return 100.0
        hit = sum(len(m.hit) for m in self.modules.values())
        return 100.0 * hit / coverable


@dataclass
class TestCaseResult:
    """One discovered VBA test procedure and its run result."""

    name: str
    result: RunResult

    @property
    def passed(self) -> bool:
        return self.result.outcome == PASSED

    @property
    def is_assertion_failure(self) -> bool:
        return (self.result.error is not None
                and self.result.error.source == "PyVbaHarness.Assert")


@dataclass
class CompileResult:
    """Outcome of a VBE compile check.

    ``accepted`` means the compile command completed and no compile-error
    dialog appeared during the watch window. ``rejected`` carries the captured
    dialog text. ``infrastructure-failure`` means the check itself could not
    be trusted (timeout, COM failure); it is never a verdict on the code.
    """

    outcome: str
    duration_s: float
    dialog: DialogRecord | None = None
    message: str = ""

    @property
    def ok(self) -> bool:
        return self.outcome == COMPILE_ACCEPTED


class HarnessError(Exception):
    """Base for harness infrastructure failures."""


class SessionDead(HarnessError):
    """The session was killed or recycled; no further commands may run on it."""


class WorkerProtocolError(HarnessError):
    """The worker produced output the supervisor could not interpret."""


class SessionLockHeld(HarnessError):
    """Another harness session holds the machine-wide Excel automation lock."""
