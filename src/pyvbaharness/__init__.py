"""pyVBAharness: a hang-resistant harness for running VBA in desktop Office.

Supports Excel, Word, PowerPoint and Access. Requires Windows, the desktop
application, and its "Trust access to the VBA project object model" option
(File > Options > Trust Center > Trust Center Settings > Macro Settings).
Access has no such option and needs no setting. See README.md.
"""
from .numbering import (
    add_line_numbers,
    instrument_error_lines,
    instrument_module,
)
from .results import (
    COMPILE_ACCEPTED,
    COMPILE_INFRA_FAILURE,
    COMPILE_REJECTED,
    MODAL_BLOCKED,
    PASSED,
    RUNNER_ERROR,
    TIMEOUT,
    VBA_ERROR,
    CompileResult,
    CoverageReport,
    DialogRecord,
    HarnessError,
    ModuleCoverage,
    RunResult,
    SessionDead,
    SessionLockHeld,
    TestCaseResult,
    VbaError,
    WorkerProtocolError,
)
from .pool import SessionPool
from .session import (
    AccessSession,
    ExcelSession,
    HarnessConfig,
    OfficeSession,
    PowerPointSession,
    WordSession,
    run_vba,
    session_for,
)

__version__ = "1.1.1"

__all__ = [
    "ExcelSession",
    "WordSession",
    "PowerPointSession",
    "AccessSession",
    "OfficeSession",
    "session_for",
    "SessionPool",
    "HarnessConfig",
    "run_vba",
    "add_line_numbers",
    "instrument_error_lines",
    "instrument_module",
    "RunResult",
    "CompileResult",
    "CoverageReport",
    "ModuleCoverage",
    "DialogRecord",
    "TestCaseResult",
    "VbaError",
    "HarnessError",
    "SessionDead",
    "SessionLockHeld",
    "WorkerProtocolError",
    "PASSED",
    "VBA_ERROR",
    "TIMEOUT",
    "MODAL_BLOCKED",
    "RUNNER_ERROR",
    "COMPILE_ACCEPTED",
    "COMPILE_REJECTED",
    "COMPILE_INFRA_FAILURE",
    "__version__",
]
