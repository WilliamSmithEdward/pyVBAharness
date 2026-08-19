"""PowerPoint COM host.

PowerPoint is the one host the harness cannot hide. ``Application.Visible =
False`` raises "Invalid request. Hiding the application window is not
allowed." (measured 2026-08-18), and there is no supported way around it, so
runs happen on a visible desktop. Everything else about ownership and
teardown is unchanged: the instance is still private to the session and still
dies with it.
"""
from __future__ import annotations

from typing import Any

from pywintypes import com_error

from .base import MSO_AUTOMATION_SECURITY_LOW, HostError, OfficeHost, _wrap_com

PP_ALERTS_NONE = 1
PP_SAVE_AS_BY_EXTENSION = {
    ".pptm": 25,  # ppSaveAsOpenXMLPresentationMacroEnabled
    ".potm": 26,  # ppSaveAsOpenXMLTemplateMacroEnabled
}


class PowerPointHost(OfficeHost):
    app_key = "powerpoint"

    # ----- lifecycle -------------------------------------------------------

    def _configure_app(self) -> None:
        app = self.app
        # No Visible = False here: PowerPoint rejects it. can_hide records
        # that so the supervisor never reports a hidden host it does not have.
        app.DisplayAlerts = PP_ALERTS_NONE
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        self._set_quietly("FeatureInstall", 0)

    def _resuppress_alerts(self) -> None:
        try:
            self.app.DisplayAlerts = PP_ALERTS_NONE
        except com_error:
            pass

    @_wrap_com("close presentation", tolerate_disconnect=True)
    def close_document(self) -> None:
        if self.document is not None:
            self._resuppress_alerts()
            # Saved = True first: an unsaved presentation makes Close raise
            # the "want to save your changes?" prompt, which has no Win32
            # buttons and would wedge teardown. Preventing the prompt beats
            # detecting it.
            try:
                self.document.Saved = True
            except com_error:
                pass
            self.document.Close()
            self.document = None
            self._reset_injection_state()

    # ----- presentations ---------------------------------------------------

    @_wrap_com("create presentation")
    def new_document(self) -> dict[str, Any]:
        self.close_document()
        self.document = self.app.Presentations.Add()
        self._reset_injection_state()
        return {"name": self._document_name(), "path": "", "unsaved": True}

    @_wrap_com("open presentation")
    def open_document(self, path: str, read_only: bool) -> dict[str, Any]:
        self.close_document()
        # Positional: FileName, ReadOnly, Untitled, WithWindow. msoTrue is -1
        # and msoFalse is 0 in the Office object model.
        opened = self.app.Presentations.Open(path, -1 if read_only else 0,
                                             0, -1)
        self.document = opened
        self._reset_injection_state()
        return {
            "name": str(opened.Name),
            "path": str(opened.FullName),
            "read_only": bool(int(opened.ReadOnly) != 0),
            "display_alerts": False,
        }

    @_wrap_com("save presentation")
    def save_as(self, path: str) -> dict[str, Any]:
        self._require_document()
        suffix = ("." + path.rsplit(".", 1)[-1].lower()) if "." in path else ""
        file_format = PP_SAVE_AS_BY_EXTENSION.get(suffix)
        if file_format is None:
            raise HostError(
                f"save_as supports .pptm and .potm, not {suffix or path!r}: "
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
        """``Module.Proc``. PowerPoint also accepts presentation-qualified
        references, but the unqualified form is what every other host uses
        and it resolves inside the single presentation this session owns."""
        return f"{module}.{proc}"
