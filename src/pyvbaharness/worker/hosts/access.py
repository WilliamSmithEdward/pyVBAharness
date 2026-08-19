"""Access COM host.

Access differs from the document-based hosts in three measured ways
(2026-08-18), and each one shapes the code below:

- ``Application.Run`` cannot be called through pywin32's dynamic dispatch
  wrapper. Every call raises DISP_E_PARAMNOTOPTIONAL (scode 0x8002000E),
  padding the 30 optional arguments with ``pythoncom.Missing`` does not help,
  and the same call through a raw ``IDispatch::Invoke`` succeeds. So this
  host invokes Run itself.
- Run resolves bare procedure names only. ``PyVbaMod.PyVbaRun`` fails with
  "Microsoft Access cannot find the procedure".
- There is no unsaved database. A database is a file that Access writes to
  continuously, so ``new_document`` creates one on disk and ``save_as`` has
  no meaning.

Access also has no "trust access to the VBA project object model" option: the
VBE object model is always reachable, which is why this host has no VBOM
preflight.
"""
from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Any

import pythoncom
from pywintypes import com_error

from .base import MSO_AUTOMATION_SECURITY_LOW, HostError, OfficeHost, _wrap_com

AC_QUIT_SAVE_NONE = 2


class AccessHost(OfficeHost):
    app_key = "access"
    vbom_settings_path = "(Access does not gate the VBA project object model)"

    def __init__(self) -> None:
        super().__init__()
        self._run_dispid: int | None = None
        self._scratch_dir: str = ""

    # ----- lifecycle -------------------------------------------------------

    def _configure_app(self) -> None:
        app = self.app
        app.Visible = False
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        self._set_quietly("FeatureInstall", 0)
        # DoCmd.SetWarnings is deliberately NOT called here. Measured
        # 2026-08-18: with no database open it fails with "The command or
        # action 'SetWarnings' isn't available now", so it is applied in
        # _open_finished once a database exists.

    def _app_hwnd(self) -> int:
        """Access exposes its main window as a method, not a property."""
        try:
            return int(self.app.hWndAccessApp())
        except (com_error, AttributeError, TypeError, ValueError):
            return 0

    def _resuppress_alerts(self) -> None:
        if self.document is None:
            return
        try:
            self.app.DoCmd.SetWarnings(False)
        except com_error:
            pass

    def _quit_app(self) -> None:
        self.app.Quit(AC_QUIT_SAVE_NONE)

    def release(self) -> None:
        super().release()
        if self._scratch_dir:
            # Access holds the .accdb and its .laccdb lock until the process
            # exits, so this only succeeds once it has; failure is harmless
            # because the directory is under the OS temp root.
            shutil.rmtree(self._scratch_dir, ignore_errors=True)
            self._scratch_dir = ""

    # ----- databases -------------------------------------------------------

    @_wrap_com("create database")
    def new_document(self) -> dict[str, Any]:
        """Create a scratch database.

        Access has no in-memory equivalent of an unsaved workbook, so the
        harness makes one in a private temp directory and deletes it at
        teardown.
        """
        self.close_document()
        self._scratch_dir = tempfile.mkdtemp(prefix="pyvbaharness-access-")
        path = str(Path(self._scratch_dir) / "harness.accdb")
        self.app.NewCurrentDatabase(path)
        self._open_finished()
        return {"name": self._document_name(), "path": path, "unsaved": False}

    @_wrap_com("open database")
    def open_document(self, path: str, read_only: bool) -> dict[str, Any]:
        """Open an existing database read-write, only when asked explicitly.

        Access has no read-only automation mode, and injecting a module
        writes into the database file immediately rather than at save time.
        Silently opening a caller's database read-write when they asked for
        read-only would exceed what they authorized, so this refuses instead.
        """
        if read_only:
            raise HostError(
                "Access cannot open a database read-only for automation, and "
                "injecting VBA modifies the file immediately. Pass "
                "read_only=False to confirm you intend to write to "
                f"{path!r}, ideally against a copy.")
        self.close_document()
        self.app.OpenCurrentDatabase(path)
        self._open_finished()
        return {
            "name": self._document_name(),
            "path": str(self.app.CurrentProject.FullName),
            "read_only": False,
            "display_alerts": False,
        }

    def _open_finished(self) -> None:
        self.document = self.app.CurrentProject
        self._reset_injection_state()
        # Now that a database exists, this call is available and turns off
        # the action-query and object-delete confirmations.
        self._resuppress_alerts()

    def save_as(self, path: str) -> dict[str, Any]:
        raise HostError(
            "Access writes changes to the database file continuously, so "
            "save_as has no meaning. The database is already at "
            f"{self._current_path()!r}; copy that file if you need a "
            "snapshot.")

    def _current_path(self) -> str:
        try:
            return str(self.app.CurrentProject.FullName)
        except com_error:
            return ""

    @_wrap_com("close database", tolerate_disconnect=True)
    def close_document(self) -> None:
        """Close the database after removing everything the harness added.

        A module added through the VBE is unsaved, and Access raises a modal
        "Save As / Module Name" prompt for each one when the database closes.
        That prompt blocked Quit outright during development on 2026-08-18
        and left Access wedged. Deleting the components first means the
        prompt never exists, which is stronger than being able to dismiss it.
        """
        if self.document is None:
            return
        self._discard_injected_modules()
        self.app.CloseCurrentDatabase()
        self.document = None
        self._reset_injection_state()

    def _discard_injected_modules(self) -> None:
        """Remove exactly the components this session added.

        Scoped to what the harness injected, never the whole project: an
        opened database's own modules belong to the caller.
        """
        try:
            components = self._components()
        except com_error:
            return
        for name in sorted(self._injected):
            try:
                component = self._find_component(components, name)
                if component is not None:
                    components.Remove(component)
            except com_error:
                continue

    # ----- VBA project -----------------------------------------------------

    def _components(self) -> Any:
        """Access keeps VBA in the application-wide project, not on a
        document object."""
        return self.app.VBE.ActiveVBProject.VBComponents

    def _activate_vbproject(self, vbe: Any) -> None:
        """Already active: Access has exactly one VBA project per database."""

    def _run_ref(self, module: str, proc: str) -> str:
        """Bare procedure name. Measured 2026-08-18: every qualified form
        (``Mod.Proc``, ``'db.accdb'!Mod.Proc``, ``db!Mod.Proc``) fails with
        "Microsoft Access cannot find the procedure"."""
        return proc

    def _invoke_run(self, ref: str, args: list[Any]) -> Any:
        """Call Application.Run through raw IDispatch.

        pywin32's dynamic wrapper cannot invoke this method on Access: every
        call raises DISP_E_PARAMNOTOPTIONAL regardless of how many trailing
        ``Missing`` values are supplied, while the identical call through
        ``IDispatch::Invoke`` returns normally (both measured 2026-08-18).
        """
        oleobj = self.app._oleobj_
        if self._run_dispid is None:
            self._run_dispid = oleobj.GetIDsOfNames("Run")
        return oleobj.Invoke(self._run_dispid, 0, pythoncom.DISPATCH_METHOD,
                             True, ref, *args)
