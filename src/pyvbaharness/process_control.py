"""Owned-process bookkeeping and termination.

Excel launched via COM is not a child of the worker process (DCOM starts it),
so killing the worker tree never kills Excel. Both reference harnesses solve
this the same way: discover the Excel PID from the Application window handle,
record it in a manifest the supervisor can read even when the worker is hung,
and kill exactly that PID on breach (``taskkill /PID <pid> /T /F``).

PID reuse is guarded by recording the process creation time next to the PID;
a sweep or kill only acts when both still match. Manifests from crashed runs
are swept at the next session start so orphaned hidden Excel processes do not
accumulate.
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import json
import os
import subprocess
import time
from pathlib import Path

_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259
_PROCESS_SET_QUOTA = 0x0100
_PROCESS_TERMINATE = 0x0001
_JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
_JOB_OBJECT_EXTENDED_LIMIT_INFORMATION = 9

MANIFEST_DIR = Path(os.environ.get("LOCALAPPDATA", ".")) / "pyvbaharness" / "sessions"
STALE_AFTER_S = 30 * 60


def process_creation_time(pid: int) -> int | None:
    """Kernel FILETIME creation stamp for a PID, or None if unqueryable.

    A process that has exited still answers this while any handle to it
    remains open, so this is identity, not liveness: pair it with
    ``is_process_alive``.
    """
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        creation = wt.FILETIME()
        exit_time = wt.FILETIME()
        kernel = wt.FILETIME()
        user = wt.FILETIME()
        ok = kernel32.GetProcessTimes(
            handle, ctypes.byref(creation), ctypes.byref(exit_time),
            ctypes.byref(kernel), ctypes.byref(user))
        if not ok:
            return None
        return (creation.dwHighDateTime << 32) | creation.dwLowDateTime
    finally:
        kernel32.CloseHandle(handle)


class _IoCounters(ctypes.Structure):
    _fields_ = [(name, ctypes.c_uint64) for name in (
        "ReadOperationCount", "WriteOperationCount", "OtherOperationCount",
        "ReadTransferCount", "WriteTransferCount", "OtherTransferCount")]


class _JobBasicLimits(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", ctypes.c_int64),
        ("PerJobUserTimeLimit", ctypes.c_int64),
        ("LimitFlags", wt.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wt.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wt.DWORD),
        ("SchedulingClass", wt.DWORD),
    ]


class _JobExtendedLimits(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", _JobBasicLimits),
        ("IoInfo", _IoCounters),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryUsed", ctypes.c_size_t),
        ("PeakJobMemoryUsed", ctypes.c_size_t),
    ]


class KillOnCloseJob:
    """A kernel job that force-kills its member processes when the last
    handle to it closes.

    The worker puts its owned Excel in one of these and simply holds the
    handle. If the worker dies for ANY reason (taskkill, crash, the whole
    Python tree being terminated), the kernel closes the handle and kills
    Excel with no cooperation from anyone. This removes the orphan-Excel
    window entirely for worker-death cases; the manifest sweep remains only
    for the pathological path where assignment itself failed.
    """

    def __init__(self) -> None:
        self._kernel32 = ctypes.windll.kernel32
        self._handle = self._kernel32.CreateJobObjectW(None, None)
        self.active = False
        if not self._handle:
            return
        limits = _JobExtendedLimits()
        limits.BasicLimitInformation.LimitFlags = (
            _JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE)
        ok = self._kernel32.SetInformationJobObject(
            self._handle, _JOB_OBJECT_EXTENDED_LIMIT_INFORMATION,
            ctypes.byref(limits), ctypes.sizeof(limits))
        if not ok:
            self._kernel32.CloseHandle(self._handle)
            self._handle = None

    def assign(self, pid: int) -> bool:
        """Add a process to the job. False (not an exception) on failure:
        the caller degrades to manifest-based cleanup."""
        if not self._handle or pid <= 0:
            return False
        process = self._kernel32.OpenProcess(
            _PROCESS_SET_QUOTA | _PROCESS_TERMINATE, False, pid)
        if not process:
            return False
        try:
            ok = bool(self._kernel32.AssignProcessToJobObject(
                self._handle, process))
            self.active = self.active or ok
            return ok
        finally:
            self._kernel32.CloseHandle(process)


def process_ids_by_image(image_name: str) -> set[int]:
    """PIDs whose executable file name matches (case-insensitive).

    Used to prove a freshly created Excel really is new: if the PID behind a
    new Application object already existed, the COM call attached to someone
    else's Excel and the harness must refuse it.
    """
    psapi = ctypes.windll.psapi
    kernel32 = ctypes.windll.kernel32
    wanted = image_name.lower()
    capacity = 1024
    while True:
        buffer = (wt.DWORD * capacity)()
        needed = wt.DWORD(0)
        if not psapi.EnumProcesses(ctypes.byref(buffer),
                                   ctypes.sizeof(buffer),
                                   ctypes.byref(needed)):
            return set()
        if needed.value < ctypes.sizeof(buffer):
            count = needed.value // ctypes.sizeof(wt.DWORD)
            break
        capacity *= 2
    found: set[int] = set()
    for index in range(count):
        pid = int(buffer[index])
        if pid <= 0:
            continue
        handle = kernel32.OpenProcess(
            _PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            continue
        try:
            size = wt.DWORD(260)
            name = ctypes.create_unicode_buffer(size.value)
            if kernel32.QueryFullProcessImageNameW(
                    handle, 0, name, ctypes.byref(size)):
                if name.value.rsplit("\\", 1)[-1].lower() == wanted:
                    found.add(pid)
        finally:
            kernel32.CloseHandle(handle)
    return found


def is_process_alive(pid: int) -> bool:
    """True only while the process is still running.

    GetExitCodeProcess is the liveness signal; OpenProcess succeeding is not,
    because the kernel keeps an exited process object alive as long as any
    handle to it exists.
    """
    if pid <= 0:
        return False
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return False
    try:
        code = wt.DWORD(0)
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
            return False
        return code.value == _STILL_ACTIVE
    finally:
        kernel32.CloseHandle(handle)


def pid_matches(pid: int, expected_creation: int | None) -> bool:
    """True when the PID is alive and is still the recorded process."""
    if not is_process_alive(pid):
        return False
    actual = process_creation_time(pid)
    if actual is None:
        return False
    return expected_creation is None or actual == expected_creation


def kill_pid_tree(pid: int, expected_creation: int | None = None) -> bool:
    """Force-kill a PID (and children) when it still matches the manifest.

    Returns True when a kill was issued. A PID that no longer matches its
    recorded creation time belongs to someone else now and is left alone.
    """
    if pid <= 0 or not pid_matches(pid, expected_creation):
        return False
    completed = subprocess.run(
        ["taskkill.exe", "/PID", str(pid), "/T", "/F"],
        capture_output=True,
        timeout=15,
        check=False,
    )
    return completed.returncode == 0


class OwnedProcessManifest:
    """One JSON file per session recording the PIDs the session owns."""

    def __init__(self, session_id: str, directory: Path | None = None) -> None:
        self.directory = directory or MANIFEST_DIR
        self.path = self.directory / f"{session_id}.json"
        self._data: dict = {"session": session_id, "written_at": time.time()}

    def record(self, role: str, pid: int) -> None:
        self._data[role] = {"pid": pid, "creation": process_creation_time(pid)}
        self._data["written_at"] = time.time()
        self.directory.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self._data), encoding="utf-8")
        os.replace(tmp, self.path)

    def entry(self, role: str) -> tuple[int, int | None] | None:
        record = self._data.get(role)
        if not isinstance(record, dict):
            return None
        return int(record.get("pid", 0)), record.get("creation")

    def kill_role(self, role: str) -> bool:
        entry = self.entry(role)
        if entry is None:
            return False
        pid, creation = entry
        return kill_pid_tree(pid, creation)

    def remove(self) -> None:
        try:
            self.path.unlink(missing_ok=True)
        except OSError:
            pass


def sweep_stale_manifests(directory: Path | None = None,
                          stale_after_s: float = STALE_AFTER_S) -> list[str]:
    """Kill exact-match orphans from manifests older than the threshold.

    Only processes whose PID and creation time still match the manifest are
    touched; everything else is stale bookkeeping and just gets deleted.
    Returns human-readable notes about what was done.
    """
    directory = directory or MANIFEST_DIR
    notes: list[str] = []
    if not directory.exists():
        return notes
    now = time.time()
    for path in directory.glob("*.json"):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if now - float(data.get("written_at", 0)) < stale_after_s:
            continue
        for role in ("excel", "worker"):
            record = data.get(role)
            if not isinstance(record, dict):
                continue
            pid = int(record.get("pid", 0))
            creation = record.get("creation")
            if pid > 0 and kill_pid_tree(pid, creation):
                notes.append(
                    f"killed orphaned {role} process {pid} from stale "
                    f"session {data.get('session', '?')}")
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
    return notes
