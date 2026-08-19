"""Word COM host."""
from __future__ import annotations

from typing import Any

from pywintypes import com_error

from .base import MSO_AUTOMATION_SECURITY_LOW, HostError, OfficeHost, _wrap_com

WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_BY_EXTENSION = {
    ".docm": 13,  # wdFormatXMLDocumentMacroEnabled
    ".dotm": 15,  # wdFormatXMLTemplateMacroEnabled
}


class WordHost(OfficeHost):
    app_key = "word"

    # ----- lifecycle -------------------------------------------------------

    def _configure_app(self) -> None:
        app = self.app
        app.Visible = False
        app.DisplayAlerts = WD_ALERTS_NONE
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        self._suppress_feature_install()
        # Word-specific prompt sources, disabled before a document exists.
        # Each of these is a dialog the harness could otherwise only discover
        # after it had already blocked a COM call.
        options = None
        try:
            options = app.Options
        except com_error:
            return
        for name, value in (
            # "Convert File" dialog when opening anything Word does not
            # recognise as native.
            ("ConfirmConversions", False),
            # AutoRecover writes can surface a save dialog mid-run.
            ("SaveInterval", 0),
            ("BackgroundSave", False),
            # Prompts to update linked fields and attached templates.
            ("UpdateLinksAtOpen", False),
            ("WarnBeforeSavingPrintingSendingMarkup", False),
            # Pure throughput: nothing reads the squiggles in a headless run.
            ("CheckSpellingAsYouType", False),
            ("CheckGrammarAsYouType", False),
        ):
            try:
                setattr(options, name, value)
            except com_error:
                continue

    def _suppress_feature_install(self) -> None:
        """msoFeatureInstallNone: fail the call instead of showing the
        Windows Installer progress dialog, which is modal, unkillable by
        message, and can sit for minutes."""
        self._set_quietly("FeatureInstall", 0)

    def _resuppress_alerts(self) -> None:
        try:
            self.app.DisplayAlerts = WD_ALERTS_NONE
        except com_error:
            pass

    @_wrap_com("close document", tolerate_disconnect=True)
    def close_document(self) -> None:
        if self.document is not None:
            self._resuppress_alerts()
            self.document.Close(WD_DO_NOT_SAVE_CHANGES)
            self.document = None
            self._reset_injection_state()

    def _quit_app(self) -> None:
        self.app.Quit(WD_DO_NOT_SAVE_CHANGES)

    # ----- documents -------------------------------------------------------

    @_wrap_com("create document")
    def new_document(self) -> dict[str, Any]:
        self.close_document()
        self.document = self.app.Documents.Add()
        self._reset_injection_state()
        return {"name": self._document_name(), "path": "", "unsaved": True}

    @_wrap_com("open document")
    def open_document(self, path: str, read_only: bool) -> dict[str, Any]:
        self.close_document()
        # Positional: FileName, ConfirmConversions, ReadOnly,
        # AddToRecentFiles. ConfirmConversions=False is load-bearing: it is
        # what stops the modal "Convert File" dialog.
        opened = self.app.Documents.Open(path, False, read_only, False)
        self.document = opened
        self._reset_injection_state()
        return {
            "name": str(opened.Name),
            "path": str(opened.FullName),
            "read_only": bool(opened.ReadOnly),
            "display_alerts": False,
        }

    @_wrap_com("save document")
    def save_as(self, path: str) -> dict[str, Any]:
        self._require_document()
        suffix = ("." + path.rsplit(".", 1)[-1].lower()) if "." in path else ""
        file_format = WD_FORMAT_BY_EXTENSION.get(suffix)
        if file_format is None:
            raise HostError(
                f"save_as supports .docm and .dotm, not {suffix or path!r}: "
                "other formats silently drop VBA under suppressed alerts.")
        self._resuppress_alerts()
        self.document.SaveAs2(path, file_format)
        return {"path": str(self.document.FullName)}

    # ----- VBA project -----------------------------------------------------

    def _components(self) -> Any:
        return self.document.VBProject.VBComponents

    def _activate_vbproject(self, vbe: Any) -> None:
        vbe.ActiveVBProject = self.document.VBProject

    def _run_ref(self, module: str, proc: str) -> str:
        """Word takes ``Module.Proc`` and rejects document qualification.

        Measured 2026-08-18: ``'Document1'!Mod.Proc`` and
        ``Document1!Mod.Proc`` both fail, while ``Mod.Proc`` and a bare
        ``Proc`` succeed. The harness owns its Word instance and keeps
        exactly one document open, so an unqualified reference cannot
        resolve into a document the caller did not mean.
        """
        return f"{module}.{proc}"
