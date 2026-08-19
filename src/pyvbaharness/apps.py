"""Per-application capabilities, measured rather than assumed.

Imported by the supervisor, the pool and the CLI, so it must stay free of
COM and of pywin32: it is plain data plus lookups.

Every field here was established by running the thing on Office 365 x64
(2026-08-18); see docs/IMPLEMENTATION_GUIDE.md for the measurements behind
each one.
"""
from __future__ import annotations

from dataclasses import dataclass

#: Office major version key used under HKCU\Software\Microsoft\Office.
OFFICE_VERSION_FALLBACK = "16.0"


@dataclass(frozen=True)
class AppInfo:
    key: str
    label: str
    progid: str
    image_name: str
    document_noun: str
    #: Registry subkey under Software\Microsoft\Office\<ver>\ for Trust
    #: Center settings. Empty when the app has no VBOM gate at all.
    security_key: str
    #: False when a second CoCreateInstance returns the running process
    #: instead of starting a new one. A single-instance app cannot be
    #: pooled, cannot run two concurrent sessions, and cannot be used at all
    #: while the user has their own copy open, because the harness would
    #: have to take ownership of a process it does not own.
    multi_instance: bool = True
    #: False when Application.Visible = False is refused.
    can_hide: bool = True
    #: Whether "Trust access to the VBA project object model" applies.
    needs_vbom: bool = True
    #: File extensions save_as accepts (macro-enabled formats only).
    save_formats: tuple[str, ...] = ()
    #: Worksheet grid: read_range / write_range / run_batch.
    has_grid: bool = False


APPS: dict[str, AppInfo] = {
    "excel": AppInfo(
        key="excel", label="Excel", progid="Excel.Application",
        image_name="EXCEL.EXE", document_noun="workbook",
        security_key="Excel", save_formats=(".xlsm", ".xlsb"),
        has_grid=True),
    "word": AppInfo(
        key="word", label="Word", progid="Word.Application",
        image_name="WINWORD.EXE", document_noun="document",
        security_key="Word", save_formats=(".docm", ".dotm")),
    "powerpoint": AppInfo(
        key="powerpoint", label="PowerPoint",
        progid="PowerPoint.Application", image_name="POWERPNT.EXE",
        document_noun="presentation", security_key="PowerPoint",
        # Measured: two CoCreateInstance calls returned one process.
        multi_instance=False,
        # Measured: "Invalid request. Hiding the application window is not
        # allowed."
        can_hide=False,
        save_formats=(".pptm", ".potm")),
    "access": AppInfo(
        key="access", label="Access", progid="Access.Application",
        image_name="MSACCESS.EXE", document_noun="database",
        # Access exposes no "trust access to the VBA project" option; the
        # VBE object model is always reachable.
        security_key="", needs_vbom=False,
        # Access writes the database continuously; there is nothing to
        # save_as.
        save_formats=()),
}

APP_KEYS = tuple(APPS)


def info(app: str) -> AppInfo:
    try:
        return APPS[app]
    except KeyError:
        raise ValueError(
            f"Unknown app {app!r}; expected one of {', '.join(APP_KEYS)}."
        ) from None


def single_instance_reason(app: str) -> str:
    """Why this app cannot have two sessions at once."""
    detail = info(app)
    return (
        f"{detail.label} is single-instance: a second COM activation returns "
        f"the {detail.image_name} process that is already running instead of "
        "starting a new one. The harness kills the process it owns on a "
        f"hang, so it will not share one. Close {detail.label} and run a "
        "single session at a time.")
