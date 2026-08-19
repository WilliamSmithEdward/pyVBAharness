"""Machine-wide mutexes.

Three named mutexes coordinate Office automation across processes:

- The SESSION mutex serializes whole sessions, and is per application: one
  app's UI is its own surface, so an exclusive Word session has no reason to
  block an exclusive Excel one. Office UI automation is mostly a single-user
  surface (XLIDE oracle README: "run oracle commands sequentially"), so
  exclusive sessions hold this for their lifetime. SessionPool members opt
  out (``exclusive=False``): hidden macro runs and range IO are safe across
  separate instances because dialog handling is PID-scoped and
  message-based, not focus-based.
- The CREATE mutex serializes instance creation. Proving a new Office
  process belongs to this session is a snapshot / create / diff sequence, and
  two sessions creating at once would each see the other's new process in the
  diff. Held only for the seconds an instance takes to start.
- The COMPILE mutex serializes compile checks only, machine-wide across
  every app. A compile check makes the host and the VBE visible and drives
  the VBE command bar; the VBE really is a shared UI surface, so it stays
  one-at-a-time even for pool sessions.

All are abandoned-safe: if a holder dies, the OS hands the mutex to the
next waiter with WAIT_ABANDONED, which we treat as acquired.
"""
from __future__ import annotations

import ctypes

from .results import SessionLockHeld

SESSION_MUTEX_NAME = "Global\\pyvbaharness-excel-session"
COMPILE_MUTEX_NAME = "Global\\pyvbaharness-compile-check"
CREATE_MUTEX_NAME = "Global\\pyvbaharness-app-create"


def session_mutex_name(app: str) -> str:
    """Per-app session mutex name.

    Excel keeps the 1.x name so a session started by an older version still
    excludes one started by a newer one.
    """
    if app == "excel":
        return SESSION_MUTEX_NAME
    return f"Global\\pyvbaharness-{app}-session"


_WAIT_OBJECT_0 = 0x0
_WAIT_ABANDONED = 0x80
_WAIT_TIMEOUT = 0x102


class SessionLock:
    def __init__(self, timeout_s: float = 0.0,
                 name: str = SESSION_MUTEX_NAME,
                 purpose: str = "session") -> None:
        self.timeout_s = timeout_s
        self.name = name
        self.purpose = purpose
        self._handle = None

    def acquire(self) -> None:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, self.name)
        if not handle:
            raise SessionLockHeld(
                f"Could not create the {self.purpose} mutex.")
        wait_ms = int(self.timeout_s * 1000)
        result = kernel32.WaitForSingleObject(handle, wait_ms)
        if result in (_WAIT_OBJECT_0, _WAIT_ABANDONED):
            self._handle = handle
            return
        kernel32.CloseHandle(handle)
        if self.purpose == "create":
            raise SessionLockHeld(
                "Timed out waiting to create an Office instance. Another "
                "session has been starting one for too long; check for a "
                "stuck Office process.")
        if self.purpose == "compile":
            raise SessionLockHeld(
                "Another compile check is still running. Compile checks are "
                "serialized machine-wide because they drive the visible VBE.")
        raise SessionLockHeld(
            "Another pyvbaharness session is running on this machine for "
            "this application. Office automation is sequential by contract; "
            "wait for it to finish, use SessionPool for parallel work, or "
            "pass exclusive=False to opt out.")

    def release(self) -> None:
        if self._handle is None:
            return
        kernel32 = ctypes.windll.kernel32
        kernel32.ReleaseMutex(self._handle)
        kernel32.CloseHandle(self._handle)
        self._handle = None

    def __enter__(self) -> "SessionLock":
        self.acquire()
        return self

    def __exit__(self, *_exc: object) -> None:
        self.release()
