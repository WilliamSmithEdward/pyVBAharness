"""Wire protocol between supervisor and worker.

Commands travel as one JSON object per line on the worker's stdin. Events
travel as one prefixed JSON object per line on the worker's stdout; any
unprefixed stdout line is passthrough logging. The prefix keeps event parsing
unambiguous even if third-party code prints to stdout inside the worker.

Every event carries ``session`` (the session id) and ``kind``. Command events
carry ``cid`` (command id) so the supervisor can match results to requests.
"""
from __future__ import annotations

import json
from typing import Any

EVENT_PREFIX = "PYVBA_EVENT|"

# Event kinds (the trace vocabulary; see oracle.py for the invariants).
EV_READY = "worker-ready"
EV_EXCEL_CREATED = "excel-created"
EV_WORKBOOK_OPENED = "workbook-opened"
EV_WORKBOOK_CREATED = "workbook-created"
EV_COMMAND_STARTED = "command-started"
EV_COMMAND_FINISHED = "command-finished"
EV_MODAL_DETECTED = "modal-detected"
EV_MODAL_DISMISSED = "modal-dismissed"
EV_MODAL_BLOCKED = "modal-blocked"
EV_VBE_WINDOW = "vbe-window-visible"
EV_WORKBOOK_CLOSED = "workbook-closed"
EV_EXCEL_QUIT = "excel-quit"
EV_EXCEL_KILLED = "excel-killed"
EV_PHASE = "host-phase"

# Command names understood by the worker.
CMD_PING = "ping"
CMD_NEW_WORKBOOK = "new_workbook"
CMD_OPEN_WORKBOOK = "open_workbook"
CMD_ADD_MODULE = "add_module"
CMD_REMOVE_MODULE = "remove_module"
CMD_RUN = "run"
CMD_RUN_RAW = "run_raw"
CMD_READ_RANGE = "read_range"
CMD_WRITE_RANGE = "write_range"
CMD_SAVE_AS = "save_as"
CMD_LIST_PROCS = "list_procs"
CMD_COMPILE = "compile"
CMD_SET_VISIBLE = "set_visible"
CMD_SHUTDOWN = "shutdown"


def encode_command(cid: int, name: str, params: dict[str, Any]) -> str:
    return json.dumps({"cid": cid, "cmd": name, "params": params},
                      ensure_ascii=False)


def decode_command(line: str) -> dict[str, Any]:
    parsed = json.loads(line)
    if (not isinstance(parsed, dict)
            or not isinstance(parsed.get("cmd"), str)
            or not isinstance(parsed.get("cid"), int)):
        raise ValueError(f"malformed command line: {line!r}")
    params = parsed.get("params")
    if params is None:
        parsed["params"] = {}
    elif not isinstance(params, dict):
        raise ValueError(f"malformed command params: {line!r}")
    return parsed


def encode_event(payload: dict[str, Any]) -> str:
    return EVENT_PREFIX + json.dumps(payload, ensure_ascii=False,
                                     default=_json_default)


def decode_event(line: str) -> dict[str, Any] | None:
    """Parse one worker stdout line; None for ordinary (non-event) output.

    A structurally wrong but syntactically valid JSON payload is rejected
    (returns None) rather than ingested into the session state machine.
    """
    if not line.startswith(EVENT_PREFIX):
        return None
    parsed = json.loads(line[len(EVENT_PREFIX):])
    if not isinstance(parsed, dict) or not isinstance(parsed.get("kind"), str):
        return None
    return parsed


def _json_default(obj: Any) -> Any:
    """Degrade unserializable COM values to strings instead of crashing."""
    try:
        import datetime
        if isinstance(obj, (datetime.datetime, datetime.date, datetime.time)):
            return obj.isoformat()
    except Exception:  # pragma: no cover - datetime import cannot fail
        pass
    if isinstance(obj, (bytes, bytearray)):
        return obj.decode("utf-8", "replace")
    return str(obj)
