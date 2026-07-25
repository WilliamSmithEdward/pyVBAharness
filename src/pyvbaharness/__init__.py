"""pyVBAharness: a hang-resistant harness for running VBA in desktop Excel.

Requires Windows, desktop Excel, and the Excel option "Trust access to the
VBA project object model" (File > Options > Trust Center > Trust Center
Settings > Macro Settings). See README.md.
"""
from .numbering import add_line_numbers, instrument_error_lines
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
    DialogRecord,
    HarnessError,
    RunResult,
    SessionDead,
    SessionLockHeld,
    TestCaseResult,
    VbaError,
    WorkerProtocolError,
)
from .pool import SessionPool
from .session import ExcelSession, HarnessConfig, run_vba

__version__ = "0.3.0"

__all__ = [
    "ExcelSession",
    "SessionPool",
    "HarnessConfig",
    "run_vba",
    "add_line_numbers",
    "instrument_error_lines",
    "RunResult",
    "CompileResult",
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
