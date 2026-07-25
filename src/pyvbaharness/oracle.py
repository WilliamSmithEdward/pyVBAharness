"""Trace oracle: validates a session's event trace against the lifecycle
contract.

Ported from XLIDE's vbaTestHostOracle. The oracle is pure: it consumes the
recorded event dicts (protocol.py vocabulary) and returns issues. Unit tests
replay synthetic traces without Excel; ExcelSession validates its own trace
at teardown and surfaces violations, so a regression in the harness's own
safety behavior fails loudly instead of silently.

Invariants:

- exactly one owned Excel instance per session; attaching is forbidden
- alerts suppressed and events disabled on the owned instance
- every command carries a positive timeout
- no command starts after the owned Excel was killed
- a timeout / modal-blocked / hang outcome is followed by excel-killed
- a normal session closes workbooks without saving (unless an explicit
  save_as command ran) and quits Excel
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from . import protocol as p


@dataclass
class OracleIssue:
    code: str
    message: str
    event_index: int | None = None


def validate_session_trace(events: list[dict[str, Any]]) -> list[OracleIssue]:
    """Validate a session trace that may span worker recycles.

    Each excel-created starts a new lifecycle generation; every generation
    must independently satisfy the contract. Without the split, a recycle
    would falsely trip single-owned-excel-instance and
    no-commands-after-kill.
    """
    segments: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    for event in events:
        if event.get("kind") == p.EV_EXCEL_CREATED and current:
            segments.append(current)
            current = []
        current.append(event)
    if current:
        segments.append(current)
    issues: list[OracleIssue] = []
    for segment in segments:
        issues.extend(validate_trace(segment))
    return issues


def validate_trace(events: list[dict[str, Any]]) -> list[OracleIssue]:
    issues: list[OracleIssue] = []
    if not events:
        return [OracleIssue("empty-trace",
                            "A session trace must record the Excel lifecycle.")]

    created = _indexed(events, p.EV_EXCEL_CREATED)
    killed = _indexed(events, p.EV_EXCEL_KILLED)
    quit_events = _indexed(events, p.EV_EXCEL_QUIT)
    started = _indexed(events, p.EV_COMMAND_STARTED)
    finished = _indexed(events, p.EV_COMMAND_FINISHED)
    closed = _indexed(events, p.EV_WORKBOOK_CLOSED)

    if len(created) != 1:
        issues.append(OracleIssue(
            "single-owned-excel-instance",
            "A session must create exactly one owned Excel instance.",
            created[0][0] if created else None))
    for index, event in created:
        if event.get("attached"):
            issues.append(OracleIssue(
                "attached-excel-instance",
                "The harness must never attach to a user Excel instance.",
                index))
        if event.get("display_alerts") is not False:
            issues.append(OracleIssue(
                "suppress-alerts",
                "The owned Excel instance must run with DisplayAlerts=False.",
                index))
        if event.get("enable_events") is not False:
            issues.append(OracleIssue(
                "suppress-events",
                "The owned Excel instance must run with EnableEvents=False.",
                index))

    first_kill = killed[0][0] if killed else None
    for index, event in started:
        timeout_ms = event.get("timeout_ms")
        if not isinstance(timeout_ms, (int, float)) or timeout_ms <= 0:
            issues.append(OracleIssue(
                "command-timeout",
                "Every command must carry a positive timeout so hangs are "
                "bounded.",
                index))
        if first_kill is not None and index > first_kill:
            issues.append(OracleIssue(
                "no-commands-after-kill",
                "No command may start after the owned Excel was killed.",
                index))

    hang_outcomes = {"timeout", "modal-blocked", "hung"}
    first_hang = next(
        ((i, e) for i, e in finished if e.get("outcome") in hang_outcomes),
        None)
    if first_hang is not None:
        hang_index = first_hang[0]
        if not any(i > hang_index for i, _ in killed):
            issues.append(OracleIssue(
                "hang-cleanup",
                "A timeout or blocked modal must be followed by killing the "
                "owned Excel instance.",
                hang_index))
        return issues

    saved_explicitly = any(
        e.get("command") == p.CMD_SAVE_AS and e.get("outcome") == "passed"
        for _, e in finished)
    for index, event in closed:
        if event.get("save_changes") and not saved_explicitly:
            issues.append(OracleIssue(
                "close-without-saving",
                "Workbooks must close without saving unless an explicit "
                "save command ran.",
                index))

    if not killed and not quit_events:
        issues.append(OracleIssue(
            "normal-cleanup",
            "A session must quit (or kill) its owned Excel instance.",
            None))

    return issues


def _indexed(events: list[dict[str, Any]],
             kind: str) -> list[tuple[int, dict[str, Any]]]:
    return [(i, e) for i, e in enumerate(events) if e.get("kind") == kind]
