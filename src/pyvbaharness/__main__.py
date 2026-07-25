"""Command-line interface: python -m pyvbaharness <command>.

Commands:
  doctor            diagnose the environment (add --live for a real smoke run)
  run FILE          run a procedure from a .bas/.vba source file
  check FILE...     compile-check source files in a fresh workbook
  check --workbook  compile-check an existing workbook's VBA project

Exit codes: 0 success/accepted, 1 VBA failure/rejected, 2 harness or
environment failure. A timeout is infrastructure (2), never a verdict on the
code.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path


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

    progid = _read_registry(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Classes\Excel.Application\CurVer", None)
    rows.append(("Excel installed", progid is not None,
                 str(progid or "Excel.Application ProgID not registered")))

    office_version = ""
    if isinstance(progid, str) and progid.rsplit(".", 1)[-1].isdigit():
        office_version = progid.rsplit(".", 1)[-1] + ".0"
    if office_version:
        access = _read_registry(
            winreg.HKEY_CURRENT_USER,
            rf"Software\Microsoft\Office\{office_version}\Excel\Security",
            "AccessVBOM")
        rows.append((
            "Trust access to the VBA project object model",
            access == 1,
            "enabled" if access == 1 else
            "DISABLED. File > Options > Trust Center > Trust Center "
            "Settings > Macro Settings > tick 'Trust access to the VBA "
            "project object model'."))
    else:
        rows.append(("Trust access to the VBA project object model", None,
                     "could not determine the Office version"))

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

    source = Path(args.file).read_text(encoding="utf-8-sig")
    call_args = tuple(_parse_cli_arg(a) for a in args.arg)
    with ExcelSession() as session:
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

    if not args.files and not args.workbook:
        _print("check needs source files or --workbook")
        return 2
    with ExcelSession() as session:
        if args.workbook:
            # Compile the workbook exactly as-is: no injected modules.
            session.open_workbook(args.workbook, read_only=True)
            result = session.compile_project(watch_seconds=args.watch)
        else:
            session.new_workbook()
            for file_name in args.files:
                path = Path(file_name)
                module = path.stem
                if not is_vba_identifier(module):
                    _print(f"skipping {path.name}: file stem is not a valid "
                           "VBA module name")
                    return 2
                kind = "class" if path.suffix.lower() == ".cls" else "standard"
                session.add_module(module,
                                   path.read_text(encoding="utf-8-sig"),
                                   kind=kind)
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

    run = commands.add_parser("run", help="run a procedure from a VBA file")
    run.add_argument("file")
    run.add_argument("--proc", default="Main")
    run.add_argument("--timeout", type=float, default=None)
    run.add_argument("--arg", action="append", default=[],
                     help="argument for the procedure (repeatable; int, "
                          "float, and true/false are auto-converted)")
    run.set_defaults(func=cmd_run)

    check = commands.add_parser(
        "check", help="compile-check VBA files or a workbook")
    check.add_argument("files", nargs="*")
    check.add_argument("--workbook", default=None)
    check.add_argument("--watch", type=float, default=15.0,
                       help="dialog watch window in seconds (default 15)")
    check.set_defaults(func=cmd_check)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
