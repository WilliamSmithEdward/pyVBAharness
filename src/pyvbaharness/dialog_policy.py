"""Pure dialog classification and dismissal policy.

Ported from XLIDE's XlideTestModalWatcher with ROneCOne's conservatism: a
dialog offering a real decision (Yes/No/Cancel/Retry/Debug/...) is never
guessed at; it is reported as blocking so the supervisor can kill the owned
Excel. Only two dismissals are ever considered safe:

- the End button on a VBE runtime-error dialog (plus OK when the dialog has
  no decision buttons), and
- OK on a purely informational dialog (only OK/Help buttons).

Compile-error dialogs are never dismissed: they mean the project cannot run
and dismissing them just lets the host wedge in a half-broken state; the
session must recycle instead. This module is windowing-free and fully unit
tested; the worker's watcher feeds it captured window data.
"""
from __future__ import annotations

from dataclasses import dataclass, field

# Win32 standard dialog control IDs.
IDOK = 1
IDCANCEL = 2
IDABORT = 3
IDRETRY = 4
IDIGNORE = 5
IDYES = 6
IDNO = 7
IDCLOSE = 8
IDHELP = 9

_DECISION_IDS = {IDCANCEL, IDABORT, IDRETRY, IDIGNORE, IDYES, IDNO, IDCLOSE}
_DECISION_CAPTIONS = {
    "cancel", "abort", "retry", "ignore", "yes", "no", "close", "debug",
}

CLASS_COMPILE_ERROR = "compile-error"
CLASS_RUNTIME_ERROR = "runtime-error"
CLASS_VBA_MODAL = "vba-modal"
CLASS_EXCEL_MODAL = "excel-modal"


@dataclass
class ButtonCapture:
    text: str
    control_id: int = 0


@dataclass
class DialogCapture:
    """What the watcher observed about one #32770 dialog."""

    title: str
    texts: list[str] = field(default_factory=list)
    buttons: list[ButtonCapture] = field(default_factory=list)

    @property
    def message(self) -> str:
        return max(self.texts, key=len, default="")


@dataclass
class DialogDecision:
    classification: str
    safe_to_dismiss: bool
    button: ButtonCapture | None
    reason: str


def normalize_caption(text: str) -> str:
    return (text or "").replace("&", "").strip().lower()


def classify(dialog: DialogCapture) -> str:
    haystack = " ".join([dialog.title, *dialog.texts]).lower()
    if "compile error" in haystack:
        return CLASS_COMPILE_ERROR
    if "run-time error" in haystack or "runtime error" in haystack:
        return CLASS_RUNTIME_ERROR
    if "microsoft visual basic" in haystack:
        return CLASS_VBA_MODAL
    return CLASS_EXCEL_MODAL


def _is_ok(button: ButtonCapture) -> bool:
    return button.control_id == IDOK or normalize_caption(button.text) == "ok"


def _is_help(button: ButtonCapture) -> bool:
    return button.control_id == IDHELP or normalize_caption(button.text) == "help"


def _is_decision(button: ButtonCapture) -> bool:
    return (button.control_id in _DECISION_IDS
            or normalize_caption(button.text) in _DECISION_CAPTIONS)


def _find_caption(dialog: DialogCapture, caption: str) -> ButtonCapture | None:
    wanted = normalize_caption(caption)
    for button in dialog.buttons:
        if normalize_caption(button.text) == wanted:
            return button
    return None


def _find_ok(dialog: DialogCapture) -> ButtonCapture | None:
    for button in dialog.buttons:
        if _is_ok(button):
            return button
    return None


def decide(dialog: DialogCapture) -> DialogDecision:
    classification = classify(dialog)

    if classification == CLASS_COMPILE_ERROR:
        return DialogDecision(classification, False, None,
                              "compile-error-blocks-runner")

    if classification in (CLASS_RUNTIME_ERROR, CLASS_VBA_MODAL):
        end_button = _find_caption(dialog, "End")
        if end_button is None and not any(_is_decision(b) for b in dialog.buttons):
            end_button = _find_ok(dialog)
        if end_button is not None:
            return DialogDecision(classification, True, end_button,
                                  "dismiss-vbe-error-end")

    ok_button = _find_ok(dialog)
    only_informational = (
        bool(dialog.buttons)
        and ok_button is not None
        and all(_is_ok(b) or _is_help(b) for b in dialog.buttons)
    )
    if only_informational:
        return DialogDecision(classification, True, ok_button,
                              "dismiss-informational-ok")

    return DialogDecision(classification, False, None,
                          "decision-or-unknown-dialog")


def fingerprint(dialog: DialogCapture, handle: int) -> str:
    """Stable dedup key so each unique dialog surfaces exactly once."""
    return "|".join([
        str(handle),
        dialog.title,
        dialog.message,
        *(b.text for b in dialog.buttons),
    ])
