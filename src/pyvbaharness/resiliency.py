"""Office crash-recovery bookkeeping.

When an Office application is terminated rather than quit, it records what it
had open under::

    HKCU\\Software\\Microsoft\\Office\\<ver>\\<App>\\Resiliency\\DocumentRecovery

and the next launch greets the user with a Document Recovery pane, or offers
safe mode once enough terminations accumulate. The harness kills its host by
design when a run hangs, so without this every killed run leaves the user a
recovery prompt naming a file the harness was working on.

Measured 2026-09-22 on Office 365 x64. Killing a session whose only document
was an unsaved scratch workbook wrote nothing. Killing one that had opened a
workbook on disk wrote an entry naming that workbook's full path, and
``Workbook.EnableAutoRecover = False`` did not prevent it: the entry comes
from crash detection, not from AutoRecover saving. So this cleans up rather
than prevents.

Removal is scoped to documents the calling session owned. On the machine
where this was written the same tree held two entries from an unrelated
project, so a blanket wipe of DocumentRecovery would have discarded a
user's pending recovery data. An entry is deleted only when a path the
session actually opened appears inside it.

Pure winreg: no COM, no pywin32, so the supervisor can call it while the
worker is already gone.
"""
from __future__ import annotations

import re
import winreg
from dataclasses import dataclass

from . import apps

_OFFICE_ROOT = r"Software\Microsoft\Office"
_VERSION_RE = re.compile(r"^\d{1,2}\.\d$")
_RECOVERY_SUBKEY = "Resiliency\\DocumentRecovery"


@dataclass(frozen=True)
class RemovedEntry:
    """One recovery entry this session cleaned up."""

    key_path: str
    matched_path: str


def _subkeys(root, path: str) -> list[str]:
    try:
        key = winreg.OpenKey(root, path)
    except OSError:
        return []
    names = []
    try:
        index = 0
        while True:
            try:
                names.append(winreg.EnumKey(key, index))
            except OSError:
                break
            index += 1
    finally:
        winreg.CloseKey(key)
    return names


def office_versions() -> list[str]:
    """Office version keys present for this user, newest first.

    Enumerated rather than assumed, so a machine with more than one Office
    version, or one the fallback does not name, is still covered.
    """
    found = [name for name in _subkeys(winreg.HKEY_CURRENT_USER, _OFFICE_ROOT)
             if _VERSION_RE.match(name)]
    if not found:
        found = [apps.OFFICE_VERSION_FALLBACK]
    return sorted(found, key=lambda v: float(v), reverse=True)


def _entry_text(root, path: str) -> str:
    """Every string in an entry's values, concatenated.

    The recovery blob is undocumented binary that carries the document path
    as UTF-16. Decoding the whole thing and searching it is enough to tell
    whose document an entry is about, and does not depend on the layout.
    """
    try:
        key = winreg.OpenKey(root, path)
    except OSError:
        return ""
    parts: list[str] = []
    try:
        index = 0
        while True:
            try:
                _name, data, _kind = winreg.EnumValue(key, index)
            except OSError:
                break
            index += 1
            if isinstance(data, bytes):
                parts.append(data.decode("utf-16-le", "ignore"))
            elif isinstance(data, str):
                parts.append(data)
    finally:
        winreg.CloseKey(key)
    return "\x00".join(parts)


def _delete_tree(root, path: str) -> None:
    for child in _subkeys(root, path):
        _delete_tree(root, f"{path}\\{child}")
    winreg.DeleteKey(root, path)


def clear_document_recovery(app: str, paths: list[str]) -> list[RemovedEntry]:
    """Delete recovery entries naming any of ``paths``.

    ``paths`` are the documents the session owned. An entry matching none of
    them belongs to someone else and is left alone, so this can never
    discard recovery data the harness did not create. Returns what was
    removed, for the trace.

    Registry failures are swallowed: this runs on the teardown path, where
    it must never turn a handled abort into an unhandled error.
    """
    wanted = [p.lower() for p in paths if p]
    if not wanted:
        return []
    detail = apps.info(app)
    removed: list[RemovedEntry] = []
    for version in office_versions():
        base = (f"{_OFFICE_ROOT}\\{version}\\{detail.registry_key}"
                f"\\{_RECOVERY_SUBKEY}")
        for entry in _subkeys(winreg.HKEY_CURRENT_USER, base):
            entry_path = f"{base}\\{entry}"
            text = _entry_text(winreg.HKEY_CURRENT_USER, entry_path).lower()
            if not text:
                continue
            match = next((p for p in wanted if p in text), "")
            if not match:
                continue
            try:
                _delete_tree(winreg.HKEY_CURRENT_USER, entry_path)
            except OSError:
                continue
            removed.append(RemovedEntry(key_path=entry_path,
                                        matched_path=match))
    return removed
