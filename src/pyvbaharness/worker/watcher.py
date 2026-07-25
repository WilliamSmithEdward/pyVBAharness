"""Dialog watcher thread: Win32 window scanning for the owned Excel PID.

Runs inside the worker process as a daemon thread. It never calls COM (the
main thread's STA may be blocked inside Excel at any moment); everything here
is user32 via ctypes, so it keeps observing and dismissing dialogs while the
main thread is stuck in ``Application.Run``.

Behavior (references: XlideTestModalWatcher.cs, watch_vbe_dialogs.ps1):

- every 250 ms, enumerate visible top-level windows of the owned PID
- capture #32770 dialogs (title, child texts, buttons with control IDs)
- classify and act through dialog_policy.decide: safe dismissals are clicked
  with SendMessageTimeout(BM_CLICK, SMTO_ABORTIFHUNG, 500 ms) so a hung
  window cannot hang the watcher; everything else is reported blocked
- report the VBE main window (class wndclass_desked_gsk) becoming visible,
  which during a hidden run means the debugger took over (break mode)
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

_user32 = ctypes.windll.user32

_EnumWindowsProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

_BM_CLICK = 0x00F5
_SMTO_ABORTIFHUNG = 0x0002
_VBE_WINDOW_CLASS = "wndclass_desked_gsk"
_DIALOG_CLASS = "#32770"


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


def _enum_top_level(pid: int) -> list[int]:
    handles: list[int] = []

    def callback(hwnd: int, _lparam: int) -> bool:
        if _user32.IsWindowVisible(hwnd) and _window_pid(hwnd) == pid:
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
    """Polls the owned Excel PID and applies the dismissal policy."""

    def __init__(self, pid: int,
                 on_record: Callable[[WatcherRecord], None],
                 interval_s: float = 0.25,
                 artifacts_dir: str = "") -> None:
        self.pid = pid
        self.on_record = on_record
        self.interval_s = interval_s
        self.artifacts_dir = artifacts_dir
        self._records: list[WatcherRecord] = []
        self._records_lock = threading.Lock()
        self._seen: set[str] = set()
        self._stop = threading.Event()
        self._suppress_vbe = threading.Event()
        self._compile_mode = threading.Event()
        self._vbe_reported = False
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
        for hwnd in _enum_top_level(self.pid):
            class_name = _window_class(hwnd)
            if class_name == _DIALOG_CLASS:
                self._handle_dialog(hwnd)
            elif class_name == _VBE_WINDOW_CLASS:
                self._handle_vbe_window(hwnd)

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
            else:
                record.action = "blocked:low-level-dismiss-failed"
        else:
            record.action = f"blocked:{decision.reason}"
        if record.action.startswith("blocked:"):
            self._attach_screenshot(record, hwnd)
        self._emit(record)

    def _attach_screenshot(self, record: WatcherRecord, hwnd: int) -> None:
        """Capture a blocking dialog before it is reported (and Excel is
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
        if self._suppress_vbe.is_set() or self._vbe_reported:
            return
        self._vbe_reported = True
        title = _window_text(hwnd)
        record = WatcherRecord(
            at=time.time(), kind="vbe-window",
            capture=DialogCapture(title=title), handle=hwnd,
            action="blocked:vbe-window-visible",
            extra={"break_mode": "[break]" in title.lower()})
        self._attach_screenshot(record, hwnd)
        self._emit(record)
