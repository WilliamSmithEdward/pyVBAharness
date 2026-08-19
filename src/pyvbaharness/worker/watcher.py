"""Dialog watcher thread: Win32 window scanning for the owned host PID.

Runs inside the worker process as a daemon thread. It never calls COM (the
main thread's STA may be blocked inside Excel at any moment); everything here
is user32 via ctypes, so it keeps observing and dismissing dialogs while
the main thread is stuck in ``Application.Run``.

Behavior (references: XlideTestModalWatcher.cs, watch_vbe_dialogs.ps1):

- every 250 ms, enumerate visible top-level windows of the owned PID
- capture #32770 dialogs (title, child texts, buttons with control IDs)
- classify and act through dialog_policy.decide: safe dismissals are clicked
  with SendMessageTimeout(BM_CLICK, SMTO_ABORTIFHUNG, 500 ms) so a hung
  window cannot hang the watcher; everything else is reported blocked
- report the VBE main window (class wndclass_desked_gsk) entering BREAK
  mode, which during a hidden run means the debugger took over. A visible
  VBE alone is not that signal: a VBE add-in can raise the window in design
  mode after a compile check, and killing the session for that is a false
  positive (measured 2026-08-18 with XLIDE loaded, title
  "XLIDE [design] - ..." reappearing when runs resumed)
- deduplicate by dialog fingerprint so each unique surface reports once
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import threading
import time
from dataclasses import dataclass, field
from typing import Callable

from ..dialog_policy import (
    CLASS_COMPILE_ERROR,
    ButtonCapture,
    DialogCapture,
    DialogDecision,
    decide,
    fingerprint,
)
from ..dialog_policy import _find_ok  # shared OK-button lookup
from ..screenshot import MAIN_WINDOW_CLASSES

_user32 = ctypes.windll.user32

_EnumWindowsProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

_BM_CLICK = 0x00F5
_SMTO_ABORTIFHUNG = 0x0002
_VBE_WINDOW_CLASS = "wndclass_desked_gsk"
_DIALOG_CLASS = "#32770"

# Office dialogs that are NOT classic Win32 #32770 windows. Modern Office
# raises "Do you want to save your changes?" as a NUIDialog whose controls
# are drawn inside a NetUIHWND surface: there are no Win32 Button children
# to enumerate or click, and GetWindowText returns nothing for any of them.
# They can be detected and reported, but not read or dismissed. Every app
# also has its own legacy "bosa_sdm_<app>" dialog class, matched by prefix
# rather than listed, so no host is missing one.
_OPAQUE_DIALOG_CLASSES = ("NUIDialog", "MsoDialog")
_OPAQUE_DIALOG_PREFIX = "bosa_sdm_"

# A titled main window is disabled exactly while something modal owns the
# application, whatever class that something is. Measured 2026-07-25: zero
# disabled samples across 15 scans of a CPU-pegged 3 s macro, against 26 of
# 28 scans while a save prompt was up. Held for this many consecutive scans
# (250 ms each) before reporting, which leaves the dismissal path time to
# clear a #32770 it can actually handle.
_MODAL_CONFIRM_SCANS = 4

# Upper bound on how long a clicked dialog may take to disappear before the
# generic modal check stops holding its fire. The check itself keys on
# whether the dialog window still exists, which is exact; this only stops a
# click that silently failed from suppressing reporting forever.
_DISMISS_MAX_S = 10.0


@dataclass
class WatcherRecord:
    """One observed dialog (or VBE window) plus what was done about it."""

    at: float
    kind: str  # dialog | vbe-window
    capture: DialogCapture
    decision: DialogDecision | None = None
    action: str = "none"  # none | click:<caption> | blocked:<reason>
    handle: int = 0
    extra: dict = field(default_factory=dict)


def _window_text(hwnd: int) -> str:
    length = _user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(max(1, length + 1))
    copied = _user32.GetWindowTextW(hwnd, buffer, len(buffer))
    return buffer.value if copied > 0 else ""


def _window_class(hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    length = _user32.GetClassNameW(hwnd, buffer, len(buffer))
    return buffer.value if length > 0 else ""


def _window_pid(hwnd: int) -> int:
    pid = wt.DWORD(0)
    _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def _click_button(hwnd: int) -> bool:
    result = wt.DWORD(0)
    ok = _user32.SendMessageTimeoutW(
        hwnd, _BM_CLICK, 0, 0, _SMTO_ABORTIFHUNG, 500, ctypes.byref(result))
    return bool(ok)


def _is_opaque_dialog(class_name: str) -> bool:
    return (class_name in _OPAQUE_DIALOG_CLASSES
            or class_name.startswith(_OPAQUE_DIALOG_PREFIX))


def _enum_top_level(pid: int) -> list[int]:
    """Every top-level window of the process, visible or not.

    Visibility is filtered per use, not here: the harness runs its host
    hidden, so the main windows are invisible, and a visible-only
    enumeration would hide the disabled-window signal that reveals a modal
    dialog.
    """
    handles: list[int] = []

    def callback(hwnd: int, _lparam: int) -> bool:
        if _window_pid(hwnd) == pid:
            handles.append(hwnd)
        return True

    _user32.EnumWindows(_EnumWindowsProc(callback), 0)
    return handles


def _capture_dialog(hwnd: int) -> DialogCapture:
    capture = DialogCapture(title=_window_text(hwnd).strip())
    buttons: list[tuple[int, ButtonCapture]] = []

    def callback(child: int, _lparam: int) -> bool:
        child_class = _window_class(child)
        text = _window_text(child).strip()
        if child_class.lower() == "button":
            control_id = _user32.GetDlgCtrlID(child)
            buttons.append((child, ButtonCapture(text=text,
                                                 control_id=control_id)))
        elif text and text not in capture.texts:
            capture.texts.append(text)
        return True

    _user32.EnumChildWindows(hwnd, _EnumWindowsProc(callback), 0)
    capture.buttons = [b for _, b in buttons]
    # Keep child handles reachable for the click step.
    capture_button_handles = {id(b): h for h, b in buttons}
    capture.__dict__["_button_handles"] = capture_button_handles
    return capture


def _button_handle(capture: DialogCapture, button: ButtonCapture) -> int:
    return capture.__dict__.get("_button_handles", {}).get(id(button), 0)


class DialogWatcher:
    """Polls the owned host PID and applies the dismissal policy."""

    def __init__(self, pid: int,
                 on_record: Callable[[WatcherRecord], None],
                 interval_s: float = 0.25,
                 artifacts_dir: str = "",
                 app: str = "excel") -> None:
        self.pid = pid
        self.on_record = on_record
        self.interval_s = interval_s
        self.artifacts_dir = artifacts_dir
        self.app = app
        self.app_label = app.capitalize()
        # Empty for an unknown app, which disables the disabled-main-window
        # signal rather than matching the wrong class.
        self.main_class = MAIN_WINDOW_CLASSES.get(app, "")
        self._records: list[WatcherRecord] = []
        self._records_lock = threading.Lock()
        self._seen: set[str] = set()
        self._stop = threading.Event()
        self._suppress_vbe = threading.Event()
        self._compile_mode = threading.Event()
        self._vbe_reported = False
        self._vbe_noted = 0
        self._modal_scans = 0
        self._modal_reported = False
        self._dismissed: tuple[int, float] = (0, 0.0)
        self._thread = threading.Thread(target=self._loop, daemon=True,
                                        name="pyvba-dialog-watcher")

    def start(self) -> None:
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        self._thread.join(timeout=2.0)

    def suppress_vbe_reporting(self, suppress: bool) -> None:
        """Compile checks legitimately show the VBE window."""
        if suppress:
            self._suppress_vbe.set()
        else:
            self._suppress_vbe.clear()
            self._vbe_reported = False
            self._vbe_noted = 0

    def set_compile_mode(self, active: bool) -> None:
        """During a compile check, compile-error dialogs ARE the evidence:
        capture them, then dismiss with OK so the compile call can return.
        Outside compile mode they stay blocking (the project cannot run and
        the session must recycle)."""
        if active:
            self._compile_mode.set()
        else:
            self._compile_mode.clear()

    def mark(self) -> int:
        """Snapshot for records_since: index of the next record."""
        with self._records_lock:
            return len(self._records)

    def records_since(self, mark: int) -> list[WatcherRecord]:
        with self._records_lock:
            return list(self._records[mark:])

    # ----- internals -------------------------------------------------------

    def _emit(self, record: WatcherRecord) -> None:
        with self._records_lock:
            self._records.append(record)
        try:
            self.on_record(record)
        except Exception:
            # The watcher must survive a reporting failure; it is the last
            # line of observation while the main thread is blocked.
            pass

    def _loop(self) -> None:
        while not self._stop.is_set():
            try:
                self._scan()
            except Exception:
                pass
            self._stop.wait(self.interval_s)

    def _scan(self) -> None:
        opaque: list[tuple[int, str]] = []
        main_disabled = False
        for hwnd in _enum_top_level(self.pid):
            class_name = _window_class(hwnd)
            if self.main_class and class_name == self.main_class:
                # Checked whether or not it is visible: the harness hides its
                # host, and a disabled main window is what modality looks
                # like from outside.
                if (_window_text(hwnd)
                        and not _user32.IsWindowEnabled(hwnd)):
                    main_disabled = True
                continue
            if not _user32.IsWindowVisible(hwnd):
                continue
            if class_name == _DIALOG_CLASS:
                self._handle_dialog(hwnd)
            elif class_name == _VBE_WINDOW_CLASS:
                self._handle_vbe_window(hwnd)
            elif _is_opaque_dialog(class_name):
                opaque.append((hwnd, class_name))
        self._check_opaque_modal(main_disabled, opaque)

    def _check_opaque_modal(self, main_disabled: bool,
                            opaque: list[tuple[int, str]]) -> None:
        """Report a modal this watcher can see but cannot classify.

        Office's own prompts (save changes, and other NetUI dialogs) carry
        no Win32 buttons, so the dismissal policy cannot act on them. Without
        this check they are invisible to the harness and the run degrades
        into a bare timeout, which sends the reader hunting for an infinite
        loop that does not exist. Detecting them turns that into
        modal-blocked plus a screenshot.
        """
        if self._compile_mode.is_set():
            return  # compile checks legitimately raise dialogs of their own
        if not main_disabled:
            self._modal_scans = 0
            self._modal_reported = False
            return
        if self._dismissal_in_flight():
            return  # a dismissible dialog is closing; that is not a block
        self._modal_scans += 1
        if self._modal_scans < _MODAL_CONFIRM_SCANS or self._modal_reported:
            return
        self._modal_reported = True
        classes = sorted({name for _hwnd, name in opaque})
        detail = (f"A {self.app_label} dialog of class {', '.join(classes)} "
                  "is blocking the run." if classes else
                  f"{self.app_label} is blocked by a modal window.")
        capture = DialogCapture(
            title=f"{self.app_label} modal dialog",
            texts=[detail,
                   "Its controls are not Win32 buttons, so the harness "
                   "cannot read or dismiss it safely."])
        record = WatcherRecord(
            at=time.time(), kind="opaque-modal", capture=capture,
            handle=opaque[0][0] if opaque else 0,
            action="blocked:opaque-modal-dialog",
            extra={"window_classes": classes})
        self._attach_screenshot(record, record.handle or self._main_window())
        self._emit(record)

    def _dismissal_in_flight(self) -> bool:
        """True while a dialog this watcher clicked is still on screen.

        The signal is the window handle going invalid, which happens exactly
        when the dialog closes, so a fast dismissal costs nothing and a slow
        one is waited out precisely. The time bound only covers a click that
        reported success without closing anything.
        """
        hwnd, clicked_at = self._dismissed
        if not hwnd:
            return False
        if (_user32.IsWindow(hwnd)
                and time.time() - clicked_at < _DISMISS_MAX_S):
            return True
        self._dismissed = (0, 0.0)
        return False

    def vbe_window_visible(self) -> bool:
        """True while a VBE window of the owned process is on screen.

        Pure ctypes, so it can be asked while the main thread is inside a COM
        call. The compile path uses it to know that hiding the VBE actually
        took effect before it re-arms VBE reporting.
        """
        for hwnd in _enum_top_level(self.pid):
            if (_window_class(hwnd) == _VBE_WINDOW_CLASS
                    and _user32.IsWindowVisible(hwnd)):
                return True
        return False

    def _main_window(self) -> int:
        if not self.main_class:
            return 0
        for hwnd in _enum_top_level(self.pid):
            if _window_class(hwnd) == self.main_class and _window_text(hwnd):
                return hwnd
        return 0

    def _handle_dialog(self, hwnd: int) -> None:
        capture = _capture_dialog(hwnd)
        key = fingerprint(capture, hwnd)
        if key in self._seen:
            return
        self._seen.add(key)
        decision = decide(capture)
        record = WatcherRecord(at=time.time(), kind="dialog", capture=capture,
                               decision=decision, handle=hwnd)
        if (self._compile_mode.is_set()
                and decision.classification == CLASS_COMPILE_ERROR):
            ok_button = _find_ok(capture)
            handle = _button_handle(capture, ok_button) if ok_button else 0
            if handle and _click_button(handle):
                record.action = "click:OK(compile-mode)"
            else:
                record.action = "blocked:compile-dialog-dismiss-failed"
            self._emit(record)
            return
        if decision.safe_to_dismiss and decision.button is not None:
            handle = _button_handle(capture, decision.button)
            clicked = handle and _click_button(handle)
            if clicked:
                record.action = f"click:{decision.button.text or 'OK'}"
                # Suppresses the generic modal check until this window is
                # actually gone, rather than for a guessed interval.
                self._dismissed = (hwnd, time.time())
            else:
                record.action = "blocked:low-level-dismiss-failed"
        else:
            record.action = f"blocked:{decision.reason}"
        if record.action.startswith("blocked:"):
            self._attach_screenshot(record, hwnd)
        self._emit(record)

    def _attach_screenshot(self, record: WatcherRecord, hwnd: int) -> None:
        """Capture a blocking dialog before it is reported (and the host is
        killed): the picture is often the fastest answer to "what dialog?"."""
        if not self.artifacts_dir:
            return
        try:
            from ..screenshot import capture_window_safely

            stamp = time.strftime("%Y%m%d-%H%M%S")
            # Bounded: the dialog's owner process is often wedged, and the
            # watcher must keep scanning no matter what GDI does.
            path = capture_window_safely(
                hwnd, f"{self.artifacts_dir}\\dialog-{stamp}-{hwnd}.png",
                timeout_s=2.0)
            if path:
                record.extra["screenshot"] = path
        except Exception:
            pass  # a failed capture must never change dialog handling

    def _handle_vbe_window(self, hwnd: int) -> None:
        """Report the VBE only when it is in break mode.

        The debugger taking over a hidden run shows the VBE with "[break]"
        in its caption, and that genuinely blocks Application.Run forever.
        A VBE window in design mode does not block anything: an add-in such
        as XLIDE re-raises it after a compile check, and treating that as a
        takeover killed healthy sessions several runs later. Design-mode
        sightings are still recorded, as an observation rather than a block,
        so the trace keeps the evidence. A break whose caption is localized
        falls through to the command deadline, which reports a timeout with
        a screenshot instead.
        """
        if self._suppress_vbe.is_set() or self._vbe_reported:
            return
        title = _window_text(hwnd)
        break_mode = "[break]" in title.lower()
        if not break_mode and hwnd == self._vbe_noted:
            return  # already noted this window in design mode
        self._vbe_reported = break_mode
        self._vbe_noted = hwnd
        record = WatcherRecord(
            at=time.time(), kind="vbe-window" if break_mode else "vbe-design",
            capture=DialogCapture(title=title), handle=hwnd,
            action=("blocked:vbe-break-mode" if break_mode
                    else "note:vbe-window-visible"),
            extra={"break_mode": break_mode})
        if break_mode:
            self._attach_screenshot(record, hwnd)
        self._emit(record)
