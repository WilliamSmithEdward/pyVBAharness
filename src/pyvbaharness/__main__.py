"""Command-line interface: python -m pyvbaharness <command>.

Commands:
  doctor            diagnose the environment (add --live for a real smoke run)
                    Reports every installed host: Excel, Word, PowerPoint,
                    Access. A host you do not have is a warning, not a
                    failure.
  run FILE          run a procedure from a .bas/.vba source file, or from an
                    Excel workbook opened read-only
  check FILE...     compile-check source files in a fresh workbook
  check WORKBOOK    compile-check an existing workbook's VBA project
  check --workbook  the same, named explicitly

A positional path routes on its extension: an Excel extension is opened as a
workbook, anything else is read as VBA source. A Word, PowerPoint, or Access
document is refused, because the command line drives Excel only.

Exit codes: 0 success/accepted, 1 VBA failure/rejected, 2 harness or
environment failure. A timeout is infrastructure (2), never a verdict on the
code.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from . import apps


def _print(line: str = "") -> None:
    print(line, flush=True)


# ----- doctor ---------------------------------------------------------------

def _read_registry(root, path: str, name: str | None):
    import winreg

    try:
        with winreg.OpenKey(root, path) as key:
            value, _kind = winreg.QueryValueEx(key, name)
            return value
    except OSError:
        return None


def _doctor_checks() -> list[tuple[str, bool | None, str]]:
    """(label, ok_or_None_for_warn, detail) rows."""
    import winreg

    rows: list[tuple[str, bool | None, str]] = []

    ok = sys.platform == "win32"
    rows.append(("Windows platform", ok, sys.platform))

    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        rows.append(("pywin32 importable", True, "ok"))
    except ImportError as err:
        rows.append(("pywin32 importable", False,
                     f"{err}; pip install pywin32"))

    office_version = ""
    for detail in apps.APPS.values():
        progid = _read_registry(
            winreg.HKEY_LOCAL_MACHINE,
            rf"SOFTWARE\Classes\{detail.progid}\CurVer", None)
        installed = progid is not None
        note = str(progid or "not installed")
        if installed and not detail.multi_instance:
            note += " (single-instance: one session at a time, no pool)"
        if installed and not detail.can_hide:
            note += " (cannot be hidden: runs on screen)"
        # A missing app is a warning, not a failure: nobody needs all four.
        rows.append((f"{detail.label} installed", True if installed else None,
                     note))
        if (installed and not office_version and isinstance(progid, str)
                and progid.rsplit(".", 1)[-1].isdigit()):
            office_version = progid.rsplit(".", 1)[-1] + ".0"

    for detail in apps.APPS.values():
        if not detail.needs_vbom:
            continue
        progid = _read_registry(
            winreg.HKEY_LOCAL_MACHINE,
            rf"SOFTWARE\Classes\{detail.progid}\CurVer", None)
        label = f"{detail.label}: trust access to the VBA project"
        if progid is None:
            continue
        if not office_version:
            rows.append((label, None, "could not determine the Office version"))
            continue
        access = _read_registry(
            winreg.HKEY_CURRENT_USER,
            rf"Software\Microsoft\Office\{office_version}"
            rf"\{detail.security_key}\Security",
            "AccessVBOM")
        rows.append((
            label,
            access == 1,
            "enabled" if access == 1 else
            f"DISABLED. In {detail.label}: File > Options > Trust Center > "
            "Trust Center Settings > Macro Settings > tick 'Trust access to "
            "the VBA project object model'."))

    # VBE error-trapping mode. 1 = Break on All Errors, which stops in the
    # debugger even for handled errors: every managed run would report
    # modal-blocked instead of vba-error.
    break_mode = None
    for vba_version in ("7.1", "7.0", "6.0"):
        break_mode = _read_registry(
            winreg.HKEY_CURRENT_USER,
            rf"Software\Microsoft\VBA\{vba_version}\Common",
            "BreakOnAllErrors")
        if break_mode is not None:
            break
    if break_mode == 1:
        rows.append(("VBE error trapping", False,
                     "'Break on All Errors' is set. In the VBE: Tools > "
                     "Options > General > Error Trapping > 'Break on "
                     "Unhandled Errors'. With break-on-all, handled VBA "
                     "errors stop in the debugger and runs report "
                     "modal-blocked."))
    else:
        rows.append(("VBE error trapping", True,
                     "break on unhandled errors"))

    return rows


def cmd_doctor(args: argparse.Namespace) -> int:
    rows = _doctor_checks()
    hard_fail = False
    for label, ok, detail in rows:
        if ok is True:
            mark = "pass"
        elif ok is None:
            mark = "warn"
        else:
            mark = "FAIL"
            hard_fail = True
        _print(f"[{mark}] {label}: {detail}")

    if args.live and not hard_fail:
        from . import ExcelSession

        _print()
        _print("Live smoke: starting an owned Excel...")
        started = time.perf_counter()
        try:
            with ExcelSession() as session:
                startup = time.perf_counter() - started
                result = session.run_vba(
                    "Public Function PyVbaDoctor() As Long\n"
                    "    PyVbaDoctor = 42\n"
                    "End Function\n", proc="PyVbaDoctor")
                if result.outcome == "passed" and result.value == 42:
                    _print(f"[pass] live run: startup {startup:.1f}s, "
                           f"run {result.duration_s * 1000:.0f}ms")
                else:
                    _print(f"[FAIL] live run: {result.outcome} "
                           f"{result.message}")
                    hard_fail = True
        except Exception as err:  # noqa: BLE001 - report, do not crash
            _print(f"[FAIL] live run: {type(err).__name__}: {err}")
            hard_fail = True

    return 2 if hard_fail else 0


# ----- input routing --------------------------------------------------------

class _CliError(Exception):
    """A bad invocation. Reported as one line, never as a traceback."""


def _read_source(path: Path) -> str:
    """Read a VBA source file as text.

    VBIDE Export writes .bas and .cls in the ANSI code page rather than
    UTF-8, so a module holding an accented identifier or string fails a
    strict UTF-8 read. Editors write UTF-8, so that is tried first.
    """
    for encoding in ("utf-8-sig", "mbcs"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
        except OSError as err:
            raise _CliError(f"{path}: {err.strerror or err}") from err
    raise _CliError(
        f"{path.name} is not readable as VBA source: it decodes as neither "
        "UTF-8 nor the ANSI code page. If it is a document rather than a "
        "module, pass it on its own so it can be opened as one.")


def _resolve_inputs(names: list[str]) -> tuple[list[Path], Path | None]:
    """Split command-line paths into VBA sources and one Excel document.

    Routing is by extension, because a workbook read as source fails deep in
    a decoder with an unhelpful message about byte 0xf6. Anything that is not
    a known Office extension is treated as source.
    """
    sources: list[Path] = []
    documents: list[Path] = []
    for name in names:
        path = Path(name)
        if not path.exists():
            raise _CliError(f"{name}: no such file")
        owner = apps.app_for_document(path.suffix)
        if owner is None:
            sources.append(path)
        elif owner == "excel":
            documents.append(path)
        else:
            label = apps.info(owner).label
            raise _CliError(
                f"{path.name} is a {label} document, and the command line "
                f"drives Excel only. Use {label}Session from Python to run "
                "VBA in that host.")
    if len(documents) > 1:
        listed = ", ".join(p.name for p in documents)
        raise _CliError(f"one workbook at a time, got: {listed}")
    if documents and sources:
        raise _CliError(
            "compile a workbook as it stands, or loose source files in a "
            "fresh workbook, but not both in one run.")
    return sources, documents[0] if documents else None


# ----- run ------------------------------------------------------------------

def _parse_cli_arg(text: str):
    for cast in (int, float):
        try:
            return cast(text)
        except ValueError:
            continue
    if text.lower() in ("true", "false"):
        return text.lower() == "true"
    return text


def cmd_run(args: argparse.Namespace) -> int:
    from . import ExcelSession

    sources, workbook = _resolve_inputs([args.file])
    # Everything that can fail on the arguments alone fails here, before
    # paying for an Excel start.
    source = _read_source(sources[0]) if sources else ""
    call_args = tuple(_parse_cli_arg(a) for a in args.arg)

    with ExcelSession() as session:
        if workbook is not None:
            # Read-only: running a macro should not rewrite the caller's
            # workbook as a side effect of asking for its result.
            session.open_workbook(workbook, read_only=True)
            result = session.run_macro(args.proc, *call_args,
                                       timeout=args.timeout)
        else:
            result = session.run_vba(source, proc=args.proc, args=call_args,
                                     timeout=args.timeout)
    for line in result.output:
        _print(line)
    if result.outcome == "passed":
        if result.value is not None:
            _print(f"value: {result.value!r}")
        return 0
    if result.error is not None:
        _print(f"error: {result.error}")
        return 1
    _print(f"{result.outcome}: {result.message}")
    return 2


# ----- check ----------------------------------------------------------------

def cmd_check(args: argparse.Namespace) -> int:
    from . import ExcelSession
    from .codegen import is_vba_identifier

    sources, workbook = _resolve_inputs(args.files)
    if args.workbook:
        if workbook is not None:
            raise _CliError(
                f"one workbook at a time: {workbook.name} was given "
                f"positionally and {Path(args.workbook).name} as --workbook.")
        if sources:
            raise _CliError(
                "compile a workbook as it stands, or loose source files in a "
                "fresh workbook, but not both in one run.")
        workbook = Path(args.workbook)
        if not workbook.exists():
            raise _CliError(f"{args.workbook}: no such file")
    if workbook is None and not sources:
        raise _CliError("check needs VBA source files or a workbook")

    # Read and validate every source before Excel starts, so a typo costs
    # nothing and never leaves a half-populated workbook behind.
    modules: list[tuple[str, str, str]] = []
    for path in sources:
        if not is_vba_identifier(path.stem):
            raise _CliError(
                f"{path.name}: the file stem is not a valid VBA module name")
        kind = "class" if path.suffix.lower() == ".cls" else "standard"
        modules.append((path.stem, kind, _read_source(path)))

    with ExcelSession() as session:
        if workbook is not None:
            # Compile the workbook exactly as-is: no injected modules.
            session.open_workbook(workbook, read_only=True)
            result = session.compile_project(watch_seconds=args.watch)
        else:
            session.new_workbook()
            for module, kind, text in modules:
                session.add_module(module, text, kind=kind)
            # Loose files are destined for the harness: compile them with
            # the support module present, the way they will actually run.
            result = session.compile_project(
                watch_seconds=args.watch, include_harness_support=True)
    if result.outcome == "accepted":
        _print(f"accepted in {result.duration_s:.1f}s")
        return 0
    if result.outcome == "rejected":
        _print("rejected: compile error")
        if result.dialog is not None:
            for text in (result.dialog.texts or [result.dialog.message]):
                _print(f"  {text}")
        return 1
    _print(f"infrastructure failure: {result.message}")
    return 2


# ----- entry ----------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m pyvbaharness",
        description="Hang-resistant VBA harness for desktop Excel.")
    commands = parser.add_subparsers(dest="command", required=True)

    doctor = commands.add_parser(
        "doctor", help="diagnose the environment for harness use")
    doctor.add_argument("--live", action="store_true",
                        help="also start Excel and run a smoke test")
    doctor.set_defaults(func=cmd_doctor)

    run = commands.add_parser(
        "run", help="run a procedure from a VBA source file or a workbook")
    run.add_argument(
        "file", metavar="FILE",
        help="a VBA source file (.bas, .cls, or any non-Office extension), "
             "injected into a fresh workbook; or an Excel workbook, opened "
             "read-only and run as it stands")
    run.add_argument("--proc", default="Main",
                     help="procedure to call, Proc or Module.Proc "
                          "(default Main)")
    run.add_argument("--timeout", type=float, default=None,
                     help="seconds before the run is killed")
    run.add_argument("--arg", action="append", default=[],
                     help="argument for the procedure (repeatable; int, "
                          "float, and true/false are auto-converted)")
    run.set_defaults(func=cmd_run)

    check = commands.add_parser(
        "check", help="compile-check VBA source files or a workbook")
    check.add_argument(
        "files", nargs="*", metavar="FILE",
        help="VBA source files to compile in a fresh workbook, or a single "
             "Excel workbook to compile as it stands")
    check.add_argument("--workbook", default=None, metavar="WORKBOOK",
                       help="an Excel workbook to compile as it stands; the "
                            "same as passing it positionally")
    check.add_argument("--watch", type=float, default=15.0,
                       help="dialog watch window in seconds (default 15)")
    check.set_defaults(func=cmd_check)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except _CliError as err:
        _print(str(err))
        return 2


if __name__ == "__main__":
    sys.exit(main())
