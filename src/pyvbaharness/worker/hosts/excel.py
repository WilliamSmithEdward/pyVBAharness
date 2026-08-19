"""Excel COM host: workbooks, worksheet ranges, and batch staging."""
from __future__ import annotations

from typing import Any

import pythoncom
from pywintypes import com_error

from ... import codegen, vbasig
from ...ranges import plan_write_chunks, validate_block
from .base import MSO_AUTOMATION_SECURITY_LOW, HostError, OfficeHost, _wrap_com

XL_FORMAT_BY_EXTENSION = {
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12
}

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

XL_SHEET_VERY_HIDDEN = 2


class ExcelHost(OfficeHost):
    app_key = "excel"
    has_enable_events = True

    # ----- lifecycle -------------------------------------------------------

    def _configure_app(self) -> None:
        app = self.app
        app.Visible = False
        app.DisplayAlerts = False
        self._set_quietly("EnableEvents", False)
        self._set_quietly("ScreenUpdating", False)
        # Prevents the "this workbook contains links" prompt, which has no
        # Win32 buttons and so cannot be clicked away after the fact.
        self._set_quietly("AskToUpdateLinks", False)
        # msoFeatureInstallNone: raise an error rather than showing the
        # Windows Installer "configuring Office" dialog, which is modal,
        # carries no Win32 buttons, and can sit for minutes.
        self._set_quietly("FeatureInstall", 0)
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        # AutoRecover saves fire on a timer and can raise a save dialog in
        # the middle of an otherwise healthy run.
        try:
            app.AutoRecover.Enabled = False
        except com_error:
            pass

    def _resuppress_alerts(self) -> None:
        try:
            self.app.DisplayAlerts = False
        except com_error:
            pass

    @_wrap_com("close workbook", tolerate_disconnect=True)
    def close_document(self) -> None:
        if self.document is not None:
            self._resuppress_alerts()
            self.document.Close(False)
            self.document = None
            self._reset_injection_state()

    # ----- workbooks -------------------------------------------------------

    @_wrap_com("create workbook")
    def new_document(self) -> dict[str, Any]:
        self.close_document()
        self.document = self.app.Workbooks.Add()
        self._reset_injection_state()
        return {"name": str(self.document.Name), "path": "", "unsaved": True}

    @_wrap_com("open workbook")
    def open_document(self, path: str, read_only: bool) -> dict[str, Any]:
        self.close_document()
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
        self.document = opened
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
        self._require_document()
        suffix = ("." + path.rsplit(".", 1)[-1].lower()) if "." in path else ""
        file_format = XL_FORMAT_BY_EXTENSION.get(suffix)
        if file_format is None:
            raise HostError(
                f"save_as supports .xlsm and .xlsb, not {suffix or path!r}: "
                "other formats silently drop VBA under suppressed alerts.")
        self._resuppress_alerts()
        self.document.SaveAs(path, file_format)
        return {"path": str(self.document.FullName)}

    # ----- VBA project -----------------------------------------------------

    def _components(self) -> Any:
        return self.document.VBProject.VBComponents

    def _activate_vbproject(self, vbe: Any) -> None:
        vbe.ActiveVBProject = self.document.VBProject

    def _run_ref(self, module: str, proc: str) -> str:
        return codegen.qualified_run_ref(self._document_name(), module, proc)

    def _bare_run_ref(self, proc: str) -> str:
        escaped = self._document_name().replace("'", "''")
        return f"'{escaped}'!{proc}"

    # ----- ranges ----------------------------------------------------------

    @_wrap_com("read range")
    def read_range(self, sheet: str, ref: str) -> list[list[Any]]:
        self._require_document()
        value = self.document.Worksheets(sheet).Range(ref).Value2
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
        self._require_document()
        try:
            width = validate_block(data)
        except ValueError as err:
            raise HostError(f"write_range: {err}") from err
        worksheet = self.document.Worksheets(sheet)
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
        self._require_document()
        sheets = self.document.Worksheets
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
        self._require_document()
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
        return str(self._invoke_run(
            self._run_ref(codegen.BATCH_MODULE_NAME, codegen.BATCH_ENTRY),
            [len(calls)]))

    def _batch_sheet(self) -> Any:
        sheets = self.document.Worksheets
        for index in range(1, int(sheets.Count) + 1):
            candidate = sheets.Item(index)
            if str(candidate.Name) == codegen.BATCH_SHEET_NAME:
                return candidate
        sheet = sheets.Add()
        sheet.Name = codegen.BATCH_SHEET_NAME
        sheet.Visible = XL_SHEET_VERY_HIDDEN  # invisible even in the UI list
        return sheet
