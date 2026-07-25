"""All Excel COM access for the worker process.

Rules enforced here (see docs/architecture.md):

- always a NEW Excel instance via CoCreateInstance with CLSCTX_LOCAL_SERVER;
  never GetActiveObject (a user's open Excel is out of bounds)
- purely late-bound dynamic dispatch: the pywin32 gencache is never touched
  (a corrupt gencache was observed on the reference machine; dynamic
  dispatch sidesteps that entire failure class)
- alerts, events, screen updating, and link prompts are disabled before any
  workbook exists
- the Excel PID is discovered from the Application window handle immediately
  after creation, so the supervisor can always kill exactly this instance
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import gc
import time
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client.dynamic
from pywintypes import com_error

from .. import codegen, vbasig
from ..process_control import KillOnCloseJob, process_ids_by_image
from ..ranges import plan_write_chunks, validate_block

# Office/Excel constants (hardcoded: no gencache, no typelib import).
MSO_AUTOMATION_SECURITY_LOW = 1
XL_FORMAT_BY_EXTENSION = {
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12
}
VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2

# Largest block written in one Value2 assignment. Measured on Excel 365 x64
# (2026-07-25): after any macro has run in the workbook, a single Value2
# assignment covering roughly 6000+ cells wedges Excel indefinitely, with the
# COM call never returning and every Excel window reporting IsHungAppWindow.
# The boundary is not a clean cell count (50x100 completed in 0.26 s while
# 100x50 took 5.5 s and 100x64 hung), so the cap is set well below the
# smallest observed failure. Chunking is also faster: 10000 cells written as
# 2500-cell blocks took 0.149 s against 0.05 s for a single pre-macro write,
# and reads are unaffected (10000 cells in 0.008 s).
MAX_WRITE_CELLS_PER_CHUNK = 2000
MSO_CONTROL_POPUP = 10
VBE_COMPILE_CONTROL_ID = 578  # "Compile <project>" on the VBE Debug menu

_ACCESS_VBOM_HINT = (
    "Excel refused programmatic access to the VBA project. Enable: File > "
    "Options > Trust Center > Trust Center Settings > Macro Settings > "
    "'Trust access to the VBA project object model'."
)


class HostError(Exception):
    """A COM operation failed; carries a diagnosable message."""

    def __init__(self, message: str, hresult: int | None = None) -> None:
        super().__init__(message)
        self.hresult = hresult


def describe_com_error(err: com_error) -> tuple[str, int | None]:
    hresult = getattr(err, "hresult", None)
    parts: list[str] = []
    if hresult is not None:
        parts.append(f"HRESULT 0x{hresult & 0xFFFFFFFF:08X}")
    excepinfo = getattr(err, "excepinfo", None)
    if excepinfo:
        source = excepinfo[1]
        description = excepinfo[2]
        scode = excepinfo[5]
        if description:
            parts.append(str(description).strip())
        if source:
            parts.append(f"source={source}")
        if scode not in (None, 0) and scode != hresult:
            parts.append(f"scode=0x{scode & 0xFFFFFFFF:08X}")
    elif getattr(err, "strerror", None):
        parts.append(str(err.strerror))
    return ("; ".join(parts) or str(err)), hresult


def _wrap_com(stage: str):
    """Decorator: convert com_error into HostError with the stage named."""

    def decorate(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except com_error as err:
                message, hresult = describe_com_error(err)
                if "not trusted" in message.lower():
                    message = f"{message} {_ACCESS_VBOM_HINT}"
                raise HostError(f"{stage}: {message}", hresult) from err
        return wrapper

    return decorate


class ExcelHost:
    def __init__(self) -> None:
        self.app: Any = None
        self.workbook: Any = None
        self.pid: int = 0
        self.job_active = False
        self.progress_path: str = ""
        self._progress_set = False
        self._cov_dims: tuple[int, int] | None = None
        self._cov_set = False
        self._job: KillOnCloseJob | None = None
        self._support_installed = False
        self._call_signature: tuple[str, str, int] | None = None
        self._batch_signature: frozenset | None = None
        self._resolved: dict[str, tuple[str, str, vbasig.ProcedureSignature]] = {}
        self._compile_control: Any = None

    # ----- lifecycle -------------------------------------------------------

    @_wrap_com("create Excel application")
    def create(self) -> dict[str, Any]:
        """Create a brand-new Excel instance and prove it is new.

        ``win32com.client.Dispatch("Excel.Application")`` must not be used
        here: given a ProgID string it first calls ``pythoncom.connect``,
        which is GetActiveObject, so it silently ATTACHES to whatever Excel
        is already running (observed live, 2026-07-25: two Dispatch calls
        returned the same PID). Attaching would put a user's own workbooks
        inside the harness's blast radius, since a timeout kills the
        instance. CoCreateInstance always launches a fresh server, and the
        PID snapshot below turns any regression into a refusal instead of a
        silent hijack.
        """
        pythoncom.CoInitialize()
        before = process_ids_by_image("EXCEL.EXE")
        clsid = pywintypes.IID("Excel.Application")
        dispatch = pythoncom.CoCreateInstance(
            clsid, None, pythoncom.CLSCTX_LOCAL_SERVER,
            pythoncom.IID_IDispatch)
        self.app = win32com.client.dynamic.Dispatch(dispatch)
        app = self.app
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.EnableEvents = False
        except com_error:
            pass
        try:
            app.ScreenUpdating = False
        except com_error:
            pass
        try:
            app.AskToUpdateLinks = False
        except com_error:
            pass
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        self.pid = self._pid_from_hwnd(int(app.Hwnd))
        if self.pid in before:
            # Do not touch it further: this Excel belongs to someone else.
            self.app = None
            raise HostError(
                f"Refusing to run: the new Excel Application resolved to "
                f"already-running process {self.pid}. The harness must own "
                "its Excel instance because it kills that process on a hang.")
        # Tie Excel's lifetime to this worker process: if the worker dies for
        # any reason, the kernel kills Excel (kill-on-close job). Assignment
        # can fail under restrictive job policies; the manifest sweep remains
        # as the fallback for that case.
        self._job = KillOnCloseJob()
        self.job_active = self._job.assign(self.pid)
        excel_version = ""
        excel_build = ""
        try:
            excel_version = str(app.Version)
            excel_build = str(app.Build)
        except com_error:
            pass
        return {
            "pid": self.pid,
            "attached": False,
            "visible": False,
            "display_alerts": False,
            "enable_events": False,
            "job_kill_on_close": self.job_active,
            "excel_version": excel_version,
            "excel_build": excel_build,
        }

    @staticmethod
    def _pid_from_hwnd(hwnd: int) -> int:
        pid = wt.DWORD(0)
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value

    @_wrap_com("close workbook")
    def close_workbook(self) -> None:
        if self.workbook is not None:
            self.workbook.Close(False)
            self.workbook = None
            self._reset_injection_state()

    @_wrap_com("quit Excel")
    def quit(self) -> None:
        if self.app is not None:
            self.app.Quit()

    def release(self) -> None:
        """Drop COM references; hangs here are covered by the supervisor's
        cleanup watchdog, not by local timeouts."""
        self.workbook = None
        self._compile_control = None
        self.app = None
        gc.collect()
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # ----- workbooks -------------------------------------------------------

    @_wrap_com("create workbook")
    def new_workbook(self) -> dict[str, Any]:
        self.close_workbook()
        self.workbook = self.app.Workbooks.Add()
        self._reset_injection_state()
        return {"name": str(self.workbook.Name), "path": "", "unsaved": True}

    @_wrap_com("open workbook")
    def open_workbook(self, path: str, read_only: bool) -> dict[str, Any]:
        self.close_workbook()
        workbooks = self.app.Workbooks
        missing = getattr(pythoncom, "Missing", None)
        opened = None
        if missing is not None:
            try:
                # Positional: Filename, UpdateLinks, ReadOnly, Format,
                # Password, WriteResPassword, IgnoreReadOnlyRecommended.
                opened = workbooks.Open(path, 0, read_only, missing, missing,
                                        missing, True)
            except (TypeError, ValueError):
                opened = None
        if opened is None:
            opened = workbooks.Open(path, 0, read_only)
        self.workbook = opened
        self._reset_injection_state()
        return {
            "name": str(opened.Name),
            "path": str(opened.FullName),
            "read_only": bool(opened.ReadOnly),
            "update_links": 0,
            "display_alerts": False,
        }

    @_wrap_com("save workbook")
    def save_as(self, path: str) -> dict[str, Any]:
        self._require_workbook()
        suffix = ("." + path.rsplit(".", 1)[-1].lower()) if "." in path else ""
        file_format = XL_FORMAT_BY_EXTENSION.get(suffix)
        if file_format is None:
            raise HostError(
                f"save_as supports .xlsm and .xlsb, not {suffix or path!r}: "
                "other formats silently drop VBA under suppressed alerts.")
        self.workbook.SaveAs(path, file_format)
        return {"path": str(self.workbook.FullName)}

    def _require_workbook(self) -> None:
        if self.workbook is None:
            raise HostError("No workbook is open in this session.")

    # ----- VBA project -----------------------------------------------------

    def _reset_injection_state(self) -> None:
        self._support_installed = False
        self._call_signature = None
        self._batch_signature = None
        self._progress_set = False
        self._resolved = {}

    @_wrap_com("add module")
    def add_module(self, name: str, source: str, kind: str) -> dict[str, Any]:
        codegen.validate_module_name(name)
        return self._write_module(name, source, kind)

    def _write_module(self, name: str, source: str,
                      kind: str) -> dict[str, Any]:
        """Create or replace a module. Skips the reserved-name check so the
        harness can inject its own modules."""
        self._require_workbook()
        self._resolved = {}  # module change invalidates cached signatures
        # Changing the VBProject resets VBA module-level state, so anything
        # pushed into the support module (progress path, coverage arrays)
        # is gone and must be re-pushed before the next run. Observed live
        # twice: progress events stopped and coverage hits vanished because
        # the dispatcher module was written after those were set.
        self._progress_set = False
        self._cov_set = False
        component_kind = (VBEXT_CT_CLASS_MODULE if kind == "class"
                          else VBEXT_CT_STD_MODULE)
        body = codegen.strip_module_header(source)
        components = self.workbook.VBProject.VBComponents
        existing = self._find_component(components, name)
        if existing is not None:
            components.Remove(existing)
        component = components.Add(component_kind)
        component.Name = name
        if body.strip():
            component.CodeModule.AddFromString(body)
        return {"name": name, "kind": kind, "lines": body.count("\n") + 1}

    @_wrap_com("remove module")
    def remove_module(self, name: str) -> dict[str, Any]:
        self._require_workbook()
        components = self.workbook.VBProject.VBComponents
        component = self._find_component(components, name)
        if component is None:
            return {"name": name, "removed": False}
        components.Remove(component)
        if name.lower() == codegen.SUPPORT_MODULE_NAME.lower():
            self._support_installed = False
        if name.lower() == codegen.CALL_MODULE_NAME.lower():
            self._call_signature = None
        return {"name": name, "removed": True}

    @staticmethod
    def _find_component(components: Any, name: str) -> Any:
        wanted = name.lower()
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            if str(component.Name).lower() == wanted:
                return component
        return None

    def _module_source(self, component: Any) -> str:
        code_module = component.CodeModule
        count = int(code_module.CountOfLines)
        if count <= 0:
            return ""
        return str(code_module.Lines(1, count))

    def _resolve_target(self, target: str
                        ) -> tuple[str, str, vbasig.ProcedureSignature]:
        """Find the target's module and parse its declaration (cached).

        A bare ``Proc`` target is searched across every component; harness
        modules are skipped so the search cannot resolve into the injected
        dispatcher. The cache is invalidated whenever a module changes.
        """
        cached = self._resolved.get(target.lower())
        if cached is not None:
            return cached
        resolved = self._resolve_target_uncached(target)
        self._resolved[target.lower()] = resolved
        return resolved

    def _resolve_target_uncached(self, target: str
                                 ) -> tuple[str, str,
                                            vbasig.ProcedureSignature]:
        parts = target.split(".")
        components = self.workbook.VBProject.VBComponents
        reserved = {n.lower() for n in codegen.HARNESS_MODULE_NAMES}
        if len(parts) == 2:
            module_name, proc_name = parts
            component = self._find_component(components, module_name)
            if component is None:
                raise HostError(
                    f"Module {module_name!r} does not exist in this "
                    "workbook's VBA project.")
            signature = vbasig.find_procedure(self._module_source(component),
                                              proc_name)
            if signature is None:
                raise HostError(
                    f"{module_name}.{proc_name} is not a callable Sub, "
                    "Function, or Property Get in that module.")
            return str(component.Name), proc_name, signature

        proc_name = parts[0]
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            name = str(component.Name)
            if name.lower() in reserved:
                continue
            signature = vbasig.find_procedure(self._module_source(component),
                                              proc_name)
            if signature is not None:
                return name, proc_name, signature
        raise HostError(
            f"No module in this workbook declares a callable {proc_name!r}.")

    def _ensure_support(self) -> None:
        if self._support_installed:
            return
        self._write_module(codegen.SUPPORT_MODULE_NAME,
                           codegen.support_module_source(), "standard")
        self._support_installed = True

    def _ensure_progress_path(self) -> None:
        """Push the progress-file path into VBA, after all module writes.

        Excel is not a child of the worker, so environment variables cannot
        carry the path. It must be (re)pushed whenever the VBProject
        changed, because that resets VBA module-level state.
        """
        if not self.progress_path or self._progress_set:
            return
        self.run_support("PyVbaSetProgressPath", [self.progress_path])
        self._progress_set = True

    def _ensure_coverage(self) -> None:
        """Re-arm the coverage arrays after a VBProject change.

        Same landmine as the progress path: writing a module resets VBA
        module-level state, which clears mCovReady and the hit arrays, so
        every hit after a dispatcher rewrite would be dropped silently.
        Re-initializing also zeroes accumulated hits, which is honest: VBA
        already discarded them.
        """
        if self._cov_dims is None or self._cov_set:
            return
        modules, max_line = self._cov_dims
        self.run_support("PyVbaCovInit", [modules, max_line])
        self._cov_set = True

    @_wrap_com("call support procedure")
    def run_support(self, proc: str, args: list[Any]) -> Any:
        """Internal Run into the support module (coverage, progress path)."""
        self._require_workbook()
        ref = codegen.qualified_run_ref(
            str(self.workbook.Name), codegen.SUPPORT_MODULE_NAME, proc)
        return self.app.Run(ref, *args)

    @_wrap_com("install support module")
    def ensure_support_module(self) -> dict[str, Any]:
        """Install the harness support module without running anything.

        Compile checks of code written for the harness need PyVbaLog and the
        assert helpers to resolve; running code gets them injected
        automatically, a bare compile does not.
        """
        self._require_workbook()
        self._ensure_support()
        return {"name": codegen.SUPPORT_MODULE_NAME}

    def _ensure_dispatcher(self, module: str, proc: str,
                           signature: vbasig.ProcedureSignature,
                           arg_count: int) -> None:
        key = (module.lower(), proc.lower(), arg_count)
        if self._call_signature == key:
            return
        self._write_module(
            codegen.CALL_MODULE_NAME,
            codegen.call_module_source(module, proc, signature, arg_count),
            "standard")
        self._call_signature = key

    # ----- running ---------------------------------------------------------

    @_wrap_com("run VBA")
    def run(self, target: str, args: list[Any]) -> str:
        """Managed run through the generated dispatcher.

        Returns the dispatcher's JSON string. The dispatcher calls the target
        directly (not through Application.Run) so a VBA error unwinds into
        its On Error handler instead of raising Excel's runtime dialog.
        """
        self._require_workbook()
        codegen.validate_run_target(target)
        if len(args) > codegen.MAX_RUN_ARGS:
            raise HostError(
                f"run supports at most {codegen.MAX_RUN_ARGS} arguments, "
                f"got {len(args)}.")
        module, proc, signature = self._resolve_target(target)
        if not signature.accepts(len(args)):
            raise HostError(
                f"{module}.{proc} takes {signature.arity_text()} "
                f"argument(s); {len(args)} were supplied.")
        self._ensure_support()
        self._ensure_dispatcher(module, proc, signature, len(args))
        self._ensure_progress_path()
        self._ensure_coverage()
        ref = codegen.qualified_run_ref(
            str(self.workbook.Name), codegen.CALL_MODULE_NAME,
            codegen.RUNNER_ENTRY)
        return str(self.app.Run(ref, *args))

    @_wrap_com("run VBA (raw)")
    def run_raw(self, target: str, args: list[Any]) -> Any:
        """Direct Application.Run without the runner (no in-VBA trapping)."""
        self._require_workbook()
        codegen.validate_run_target(target)
        parts = target.split(".")
        if len(parts) == 2:
            ref = codegen.qualified_run_ref(str(self.workbook.Name),
                                            parts[0], parts[1])
        else:
            escaped = str(self.workbook.Name).replace("'", "''")
            ref = f"'{escaped}'!{parts[0]}"
        return self.app.Run(ref, *args)

    # ----- ranges ----------------------------------------------------------

    @_wrap_com("read range")
    def read_range(self, sheet: str, ref: str) -> list[list[Any]]:
        self._require_workbook()
        value = self.workbook.Worksheets(sheet).Range(ref).Value2
        if isinstance(value, tuple):
            return [list(row) if isinstance(row, tuple) else [row]
                    for row in value]
        return [[value]]

    @_wrap_com("write range")
    def write_range(self, sheet: str, start_cell: str,
                    data: list[list[Any]]) -> dict[str, Any]:
        """Write a 2-D block, in bounded chunks.

        Two live-verified constraints shape this (2026-07-25):

        The target range is built from explicit corner cells rather than
        ``Range(...).Resize(rows, cols)``. Under late-bound dispatch a
        property whose parameters are all optional (Resize, Offset) is
        invoked with no arguments, and the trailing call becomes the returned
        range's default Item indexer, so Resize(2, 2) yields the single cell
        B2 instead of A1:B2.

        The block is split into MAX_WRITE_CELLS_PER_CHUNK pieces because a
        single oversized Value2 assignment wedges Excel after any macro has
        run in the workbook. See that constant for the measurements.
        """
        self._require_workbook()
        try:
            width = validate_block(data)
        except ValueError as err:
            raise HostError(f"write_range: {err}") from err
        worksheet = self.workbook.Worksheets(sheet)
        anchor = worksheet.Range(start_cell)
        chunks = self._write_block(worksheet, int(anchor.Row),
                                   int(anchor.Column), data)
        return {"rows": len(data), "columns": width, "chunks": chunks}

    def _write_block(self, worksheet: Any, base_row: int, base_column: int,
                     data: list[list[Any]]) -> int:
        """Chunked Value2 writes; shared by write_range and batch staging."""
        width = len(data[0])
        chunks = 0
        for chunk in plan_write_chunks(len(data), width,
                                       MAX_WRITE_CELLS_PER_CHUNK):
            piece = tuple(
                tuple(row[chunk.column_start:chunk.column_end])
                for row in data[chunk.row_start:chunk.row_end])
            first = worksheet.Cells(base_row + chunk.row_start,
                                    base_column + chunk.column_start)
            last = worksheet.Cells(base_row + chunk.row_end - 1,
                                   base_column + chunk.column_end - 1)
            worksheet.Range(first, last).Value2 = piece
            chunks += 1
        return chunks

    @_wrap_com("reset sheets")
    def reset_sheets(self) -> dict[str, Any]:
        """Clear every worksheet's cells while keeping injected modules.

        A cheap between-tests reset: new_workbook costs a workbook plus
        module reinjection; this costs a few Clear calls.
        """
        self._require_workbook()
        sheets = self.workbook.Worksheets
        cleared = 0
        for index in range(1, int(sheets.Count) + 1):
            sheets.Item(index).Cells.Clear()
            cleared += 1
        return {"cleared_sheets": cleared}

    # ----- batch execution -------------------------------------------------

    @_wrap_com("run VBA batch")
    def run_batch(self, calls: list[dict[str, Any]]) -> str:
        """Stage encoded calls, run the generated batch dispatcher once.

        Returns the dispatcher's JSON array string. Staging goes through the
        very hidden batch sheet with type-prefixed text cells (exact
        round-trip fidelity; see codegen.encode_batch_arg) and the chunked
        writer (the post-macro Value2 wedge applies to staging too).
        """
        self._require_workbook()
        if len(calls) > codegen.MAX_BATCH_CALLS:
            raise HostError(
                f"run_batch supports at most {codegen.MAX_BATCH_CALLS} "
                f"calls, got {len(calls)}.")
        entries: list[tuple[str, str, vbasig.ProcedureSignature, int]] = []
        rows: list[list[Any]] = []
        for index, call in enumerate(calls):
            target = str(call["target"])
            encoded_args = [str(a) for a in call.get("args", [])]
            codegen.validate_run_target(target)
            if len(encoded_args) > codegen.MAX_RUN_ARGS:
                raise HostError(
                    f"Batch call {index} has {len(encoded_args)} arguments; "
                    f"the limit is {codegen.MAX_RUN_ARGS}.")
            module, proc, signature = self._resolve_target(target)
            if not signature.accepts(len(encoded_args)):
                raise HostError(
                    f"Batch call {index}: {module}.{proc} takes "
                    f"{signature.arity_text()} argument(s); "
                    f"{len(encoded_args)} were supplied.")
            entries.append((module, proc, signature, len(encoded_args)))
            row: list[Any] = [f"{module}.{proc}", len(encoded_args)]
            row.extend(encoded_args)
            row.extend([""] * (codegen.MAX_RUN_ARGS - len(encoded_args)))
            rows.append(row)

        self._ensure_support()
        key_set = frozenset(
            codegen.batch_call_key(m, p, n) for m, p, _s, n in entries)
        if self._batch_signature != key_set:
            self._write_module(codegen.BATCH_MODULE_NAME,
                               codegen.batch_module_source(entries),
                               "standard")
            self._batch_signature = key_set
        self._ensure_progress_path()
        self._ensure_coverage()
        sheet = self._batch_sheet()
        self._write_block(sheet, 1, 1, rows)
        ref = codegen.qualified_run_ref(
            str(self.workbook.Name), codegen.BATCH_MODULE_NAME,
            codegen.BATCH_ENTRY)
        return str(self.app.Run(ref, len(calls)))

    def _batch_sheet(self) -> Any:
        sheets = self.workbook.Worksheets
        for index in range(1, int(sheets.Count) + 1):
            candidate = sheets.Item(index)
            if str(candidate.Name) == codegen.BATCH_SHEET_NAME:
                return candidate
        sheet = sheets.Add()
        sheet.Name = codegen.BATCH_SHEET_NAME
        sheet.Visible = 2  # xlSheetVeryHidden: invisible even in the UI list
        return sheet

    # ----- module export and coverage --------------------------------------

    @_wrap_com("export modules")
    def export_modules(self, directory: str) -> dict[str, Any]:
        """Export every non-harness component via VBIDE Export."""
        self._require_workbook()
        extensions = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}
        reserved = {n.lower() for n in codegen.HARNESS_MODULE_NAMES}
        exported: list[str] = []
        components = self.workbook.VBProject.VBComponents
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            name = str(component.Name)
            if name.lower() in reserved:
                continue
            extension = extensions.get(int(component.Type))
            if extension is None:
                continue
            target = str(Path(directory) / f"{name}{extension}")
            component.Export(target)
            exported.append(target)
        return {"files": exported}

    def coverage_init(self, module_count: int, max_line: int) -> None:
        self._ensure_support()
        self._cov_dims = (module_count, max_line)
        self._cov_set = False
        self._ensure_coverage()

    def coverage_report(self) -> str:
        self._ensure_support()
        return str(self.run_support("PyVbaCovReportJson", []))

    # ----- visibility and compile ------------------------------------------

    @_wrap_com("set visibility")
    def set_visible(self, visible: bool) -> dict[str, Any]:
        self.app.Visible = visible
        return {"visible": bool(self.app.Visible)}

    @_wrap_com("list procedures")
    def list_procs(self, module: str) -> list[dict[str, Any]]:
        self._require_workbook()
        component = self._find_component(
            self.workbook.VBProject.VBComponents, module)
        if component is None:
            raise HostError(
                f"Module {module!r} does not exist in this workbook's "
                "VBA project.")
        found = []
        for signature in vbasig.list_procedures(
                self._module_source(component)):
            found.append({
                "name": signature.name,
                "kind": signature.kind,
                "required": signature.required,
                "optional": signature.optional,
                "param_array": signature.has_param_array,
            })
        return found

    @_wrap_com("compile project")
    def start_compile(self) -> str:
        """Make the VBE visible and execute its Compile command control.

        Excel must be visible for the Compile error dialog to surface as a
        detectable window (XLIDE oracle lesson: hidden hosts convert compile
        rejections into silent false accepts). Returns "already-compiled"
        when the Compile control is disabled, which the VBE does exactly when
        the project is fully compiled; otherwise fires the compile and
        returns "fired". The caller owns the watch window and the verdict.
        """
        self._require_workbook()
        self.app.Visible = True
        vbe = self.app.VBE
        vbe.MainWindow.Visible = True
        vbe.ActiveVBProject = self.workbook.VBProject
        control = self._find_compile_control(vbe)
        if control is None:
            raise HostError(
                "Could not locate the VBE Compile command (control id 578).")
        if not control.Enabled:
            return "already-compiled"
        control.Execute()
        return "fired"

    def compile_control_disabled(self) -> bool:
        """True when the Compile control has gone disabled.

        The VBE disables its Compile command precisely when the project is
        fully compiled, so after a fired compile this is a POSITIVE success
        signal: the accept verdict no longer has to wait out the full
        dialog-watch window (which existed only to prove a negative).
        """
        try:
            control = self._compile_control
            return control is not None and not control.Enabled
        except com_error:
            return False

    def _find_compile_control(self, vbe: Any) -> Any:
        if self._compile_control is not None:
            return self._compile_control
        bars = vbe.CommandBars
        for bar_index in range(1, int(bars.Count) + 1):
            found = self._search_controls(bars.Item(bar_index).Controls, 0)
            if found is not None:
                self._compile_control = found
                return found
        return None

    def _search_controls(self, controls: Any, depth: int) -> Any:
        if depth > 3:
            return None
        for index in range(1, int(controls.Count) + 1):
            control = controls.Item(index)
            try:
                if int(control.Id) == VBE_COMPILE_CONTROL_ID:
                    return control
                if int(control.Type) == MSO_CONTROL_POPUP:
                    found = self._search_controls(control.Controls, depth + 1)
                    if found is not None:
                        return found
            except com_error:
                continue
        return None

    @_wrap_com("hide VBE")
    def end_compile(self) -> None:
        try:
            self.app.VBE.MainWindow.Visible = False
        except com_error:
            pass
        self.app.Visible = False
