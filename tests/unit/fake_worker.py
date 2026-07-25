"""Scripted fake worker for supervisor state-machine tests.

Speaks the pyvbaharness pipe protocol with canned behaviors keyed by the run
target, so ExcelSession's watchdog/abort/recycle logic can be exercised with
no Excel and no COM. Self-contained on purpose: stdlib only.
"""
import json
import os
import sys
import time

PREFIX = "PYVBA_EVENT|"
SESSION = sys.argv[sys.argv.index("--session") + 1] if "--session" in sys.argv else "?"


def emit(kind, **payload):
    body = {"kind": kind, "session": SESSION, "at": time.time()}
    body.update(payload)
    print(PREFIX + json.dumps(body), flush=True)


def main():
    emit("excel-created", pid=os.getpid(), attached=False,
         display_alerts=False, enable_events=False)
    emit("worker-ready", pid=os.getpid())
    opened = False
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        command = json.loads(line)
        cid = command["cid"]
        name = command["cmd"]
        params = command.get("params") or {}
        timeout_ms = params.pop("_timeout_ms", None)
        emit("command-started", cid=cid, command=name, timeout_ms=timeout_ms)
        if name == "shutdown":
            emit("command-finished", cid=cid, command=name, outcome="passed",
                 duration_ms=0)
            break
        if name == "new_workbook":
            opened = True
            emit("workbook-created", name="FakeBook")
            emit("command-finished", cid=cid, command=name, outcome="passed",
                 data={"name": "FakeBook"}, duration_ms=1)
            continue
        if name == "run":
            target = params.get("target", "")
            if target == "Hang.Forever":
                time.sleep(600)  # never answers; the supervisor must kill us
            elif target == "Modal.Blocked":
                emit("modal-blocked", title="Microsoft Excel",
                     message="Overwrite?", texts=["Overwrite?"],
                     buttons=["Yes", "No"], button_ids=[6, 7],
                     classification="excel-modal",
                     reason="decision-or-unknown-dialog",
                     action="blocked:decision-or-unknown-dialog")
                time.sleep(600)  # stuck behind the modal
            elif target == "Die.Now":
                sys.exit(3)
            else:
                emit("command-finished", cid=cid, command=name,
                     outcome="passed", value=42, output=["hi"],
                     duration_ms=1)
            continue
        emit("command-finished", cid=cid, command=name, outcome="passed",
             data={}, duration_ms=1)
    if opened:
        emit("workbook-closed", save_changes=False, duration_ms=1)
    emit("excel-quit", duration_ms=1)


if __name__ == "__main__":
    main()
