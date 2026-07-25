"""Machine-wide session lock.

Excel UI automation is a single-user surface: parallel harness sessions
contend on modal dialogs, VBE state, and process cleanup, producing timeouts
and misleading results (XLIDE oracle README, "run oracle commands
sequentially"). Sessions therefore hold a named Windows mutex by default.

The mutex is abandoned-safe: if a holder dies, the OS hands the mutex to the
next waiter with WAIT_ABANDONED, which we treat as acquired.
"""
from __future__ import annotations

import ctypes

from .results import SessionLockHeld

_MUTEX_NAME = "Global\\pyvbaharness-excel-session"
_WAIT_OBJECT_0 = 0x0
_WAIT_ABANDONED = 0x80
_WAIT_TIMEOUT = 0x102


class SessionLock:
    def __init__(self, timeout_s: float = 0.0) -> None:
        self.timeout_s = timeout_s
        self._handle = None

    def acquire(self) -> None:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, _MUTEX_NAME)
        if not handle:
            raise SessionLockHeld("Could not create the session mutex.")
        wait_ms = int(self.timeout_s * 1000)
        result = kernel32.WaitForSingleObject(handle, wait_ms)
        if result in (_WAIT_OBJECT_0, _WAIT_ABANDONED):
            self._handle = handle
            return
        kernel32.CloseHandle(handle)
        raise SessionLockHeld(
            "Another pyvbaharness session is running on this machine. "
            "Excel automation is sequential by contract; wait for it to "
            "finish, or pass exclusive=False to opt out (unsupported).")

    def release(self) -> None:
        if self._handle is None:
            return
        kernel32 = ctypes.windll.kernel32
        kernel32.ReleaseMutex(self._handle)
        kernel32.CloseHandle(self._handle)
        self._handle = None
