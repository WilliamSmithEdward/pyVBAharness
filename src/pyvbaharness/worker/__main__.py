"""Worker process entry point: ``python -m pyvbaharness.worker``.

Reads one JSON command per stdin line, emits prefixed JSON events on stdout
(see protocol.py). All COM lives here, in the main thread's STA; the dialog
watcher runs as a ctypes-only daemon thread. The worker never enforces run
timeouts itself: a hung COM call cannot be interrupted from inside, so the
supervisor watches the event stream and kills this process (and the recorded
Excel PID) from outside.
"""
from __future__ import annotations

import argparse
import faulthandler
import json
import os
import sys
import tempfile
import threading
import time
from typing import Any

from .. import codegen, protocol
from ..results import PASSED, RUNNER_ERROR, VBA_ERROR
from .excel_host import ExcelHost, HostError
from .watcher import DialogWatcher, WatcherRecord


class ProgressTail(threading.Thread):
    """Tails the VBA progress file and emits vba-progress events.

    VBA cannot call back into Python mid-run (the COM thread is blocked in
    Application.Run), so PyVbaProgress appends lines to a file and this
    thread converts them into events. VBA's Print # writes in the ANSI code
    page, hence the mbcs decode; brief sharing violations while VBA holds
    the file open are ridden out by simply retrying next tick.
    """

    def __init__(self, path: str) -> None:
        super().__init__(daemon=True, name="pyvba-progress-tail")
        self.path = path
        self._offset = 0
        self._stop = threading.Event()

    def stop(self) -> None:
        self._stop.set()

    def drain_now(self) -> None:
        """Flush pending progress lines immediately.

        Called just before a command result is emitted: the poll interval
        would otherwise drop the last heartbeat, because the supervisor
        stops reading progress once command-finished arrives.
        """
        try:
            self._drain()
        except OSError:
            pass

    def run(self) -> None:
        while not self._stop.is_set():
            try:
                self._drain()
            except OSError:
                pass
            self._stop.wait(0.2)

    def _drain(self) -> None:
        if not os.path.exists(self.path):
            return
        with open(self.path, "rb") as handle:
            handle.seek(self._offset)
            chunk = handle.read()
        if not chunk:
            return
        self._offset += len(chunk)
        for raw in chunk.decode("mbcs", "replace").splitlines():
            raw = raw.strip()
            if not raw:
                continue
            fraction_text, _sep, message = raw.partition("|")
            try:
                fraction = float(fraction_text)
            except ValueError:
                fraction = 0.0
                message = raw
            _emit(protocol.EV_PROGRESS,
                  {"fraction": fraction, "message": message})

_emit_lock = threading.Lock()
_session_id = "unknown"


def _emit(kind: str, payload: dict[str, Any] | None = None) -> None:
    body = {"kind": kind, "session": _session_id, "at": time.time()}
    if payload:
        body.update(payload)
    line = protocol.encode_event(body)
    with _emit_lock:
        print(line, flush=True)


def _record_payload(record: WatcherRecord) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "title": record.capture.title,
        "message": record.capture.message,
        "texts": record.capture.texts,
        "buttons": [b.text for b in record.capture.buttons],
        "button_ids": [b.control_id for b in record.capture.buttons],
        "action": record.action,
    }
    if record.decision is not None:
        payload["classification"] = record.decision.classification
        payload["reason"] = record.decision.reason
    payload.update(record.extra)
    return payload


def _on_watcher_record(record: WatcherRecord) -> None:
    if record.kind == "vbe-window":
        _emit(protocol.EV_VBE_WINDOW, _record_payload(record))
    elif record.action.startswith("blocked:"):
        _emit(protocol.EV_MODAL_BLOCKED, _record_payload(record))
    elif record.action.startswith("click:"):
        _emit(protocol.EV_MODAL_DISMISSED, _record_payload(record))
    else:
        _emit(protocol.EV_MODAL_DETECTED, _record_payload(record))


class Worker:
    def __init__(self, artifacts_dir: str = "") -> None:
        self.host = ExcelHost()
        self.watcher: DialogWatcher | None = None
        self.progress_tail: ProgressTail | None = None
        self.artifacts_dir = artifacts_dir

    # ----- startup / shutdown ---------------------------------------------

    def start(self) -> None:
        self.host.progress_path = os.path.join(
            tempfile.gettempdir(), f"pyvba-progress-{os.getpid()}.log")
        info = self.host.create()
        _emit(protocol.EV_EXCEL_CREATED, info)
        self.watcher = DialogWatcher(self.host.pid, _on_watcher_record,
                                     artifacts_dir=self.artifacts_dir)
        self.watcher.start()
        self.progress_tail = ProgressTail(self.host.progress_path)
        self.progress_tail.start()
        _emit(protocol.EV_READY, {"pid": self.host.pid})

    def shutdown(self) -> None:
        watcher = self.watcher
        if watcher is not None:
            watcher.stop()
        if self.progress_tail is not None:
            self.progress_tail.stop()
        try:
            os.remove(self.host.progress_path)
        except OSError:
            pass
        started = time.time()
        try:
            had_workbook = self.host.workbook is not None
            self.host.close_workbook()
            if had_workbook:
                _emit(protocol.EV_WORKBOOK_CLOSED, {
                    "save_changes": False,
                    "duration_ms": int((time.time() - started) * 1000),
                })
        except HostError as err:
            _emit(protocol.EV_PHASE, {"phase": "workbook-close",
                                      "outcome": "failed",
                                      "message": str(err)})
        started = time.time()
        try:
            self.host.quit()
            _emit(protocol.EV_EXCEL_QUIT, {
                "duration_ms": int((time.time() - started) * 1000)})
        except HostError as err:
            _emit(protocol.EV_PHASE, {"phase": "excel-quit",
                                      "outcome": "failed",
                                      "message": str(err)})
        self.host.release()
        _emit(protocol.EV_PHASE, {"phase": "com-release", "outcome": "passed"})

    # ----- command execution ----------------------------------------------

    def execute(self, name: str, params: dict[str, Any]) -> dict[str, Any]:
        """Returns the command-finished payload (outcome + data)."""
        watcher = self.watcher
        mark = watcher.mark() if watcher else 0
        try:
            data = self._dispatch(name, params)
        except HostError as err:
            return self._finish(RUNNER_ERROR, mark, {
                "message": str(err),
                "hresult": err.hresult,
            })
        except (ValueError, KeyError, TypeError) as err:
            return self._finish(RUNNER_ERROR, mark, {"message": str(err)})
        finally:
            # Emit trailing progress lines BEFORE the result: the
            # supervisor stops consuming progress once command-finished
            # arrives, so the poll interval would drop the last heartbeat.
            if self.progress_tail is not None:
                self.progress_tail.drain_now()
        if name == protocol.CMD_RUN:
            return self._finish_run(data, mark)
        return self._finish(PASSED, mark, {"data": data})

    def _finish(self, outcome: str, mark: int,
                extra: dict[str, Any]) -> dict[str, Any]:
        payload = {"outcome": outcome, **extra}
        watcher = self.watcher
        if watcher is not None:
            records = watcher.records_since(mark)
            if records:
                payload["dialogs"] = [_record_payload(r) for r in records]
        return payload

    def _finish_run(self, raw: Any, mark: int) -> dict[str, Any]:
        """Interpret the runner module's JSON string."""
        try:
            parsed = json.loads(str(raw))
            outcome = parsed.get("outcome")
        except (TypeError, ValueError):
            parsed = None
            outcome = None
        if outcome == "passed":
            return self._finish(PASSED, mark, {
                "value": parsed.get("value"),
                "output": parsed.get("output", []),
            })
        if outcome == "vba-error":
            return self._finish(VBA_ERROR, mark, {
                "number": parsed.get("number", 0),
                "source": parsed.get("source", ""),
                "description": parsed.get("description", ""),
                "line": parsed.get("line", 0),
                "stack": parsed.get("stack", []),
                "output": parsed.get("output", []),
            })
        return self._finish(RUNNER_ERROR, mark, {
            "message": "The injected runner returned an unreadable result: "
                       + repr(raw)[:500],
        })

    def _dispatch(self, name: str, params: dict[str, Any]) -> Any:
        host = self.host
        if name == protocol.CMD_PING:
            return {"pong": True, "pid": host.pid}
        if name == protocol.CMD_NEW_WORKBOOK:
            data = host.new_workbook()
            _emit(protocol.EV_WORKBOOK_CREATED, data)
            return data
        if name == protocol.CMD_OPEN_WORKBOOK:
            data = host.open_workbook(str(params["path"]),
                                      bool(params.get("read_only", True)))
            _emit(protocol.EV_WORKBOOK_OPENED, data)
            return data
        if name == protocol.CMD_ADD_MODULE:
            return host.add_module(str(params["name"]), str(params["source"]),
                                   str(params.get("kind", "standard")))
        if name == protocol.CMD_REMOVE_MODULE:
            return host.remove_module(str(params["name"]))
        if name == protocol.CMD_RUN:
            return host.run(str(params["target"]),
                            list(params.get("args", [])))
        if name == protocol.CMD_RUN_RAW:
            value = host.run_raw(str(params["target"]),
                                 list(params.get("args", [])))
            return {"value": value}
        if name == protocol.CMD_RUN_BATCH:
            raw = host.run_batch(list(params.get("calls", [])))
            try:
                items = json.loads(raw)
            except (TypeError, ValueError) as err:
                raise HostError(
                    "The batch dispatcher returned an unreadable result: "
                    + repr(raw)[:300]) from err
            if not isinstance(items, list):
                raise HostError("The batch dispatcher returned a non-list.")
            return {"results": items}
        if name == protocol.CMD_RESET_SHEETS:
            return host.reset_sheets()
        if name == protocol.CMD_EXPORT_MODULES:
            return host.export_modules(str(params["dir"]))
        if name == protocol.CMD_COV_INIT:
            host.coverage_init(int(params["modules"]),
                               int(params["max_line"]))
            return {"initialized": True}
        if name == protocol.CMD_COV_REPORT:
            raw = host.coverage_report()
            try:
                hits = json.loads(raw)
            except (TypeError, ValueError) as err:
                raise HostError("Unreadable coverage report: "
                                + repr(raw)[:300]) from err
            return {"hits": hits}
        if name == protocol.CMD_READ_RANGE:
            return {"data": host.read_range(str(params["sheet"]),
                                            str(params["ref"]))}
        if name == protocol.CMD_WRITE_RANGE:
            return host.write_range(str(params["sheet"]),
                                    str(params["start_cell"]),
                                    list(params["data"]))
        if name == protocol.CMD_SAVE_AS:
            return host.save_as(str(params["path"]))
        if name == protocol.CMD_LIST_PROCS:
            return {"procs": host.list_procs(str(params["module"]))}
        if name == protocol.CMD_SET_VISIBLE:
            return host.set_visible(bool(params.get("visible", False)))
        if name == protocol.CMD_COMPILE:
            if params.get("include_support"):
                self.host.ensure_support_module()
            return self._compile(float(params.get("watch_seconds", 10.0)))
        raise ValueError(f"Unknown command: {name}")

    def _compile(self, watch_seconds: float) -> dict[str, Any]:
        """Fire VBE Compile and watch for a compile-error dialog.

        No dialog within the watch window means accepted; the window is the
        price of a negative signal (XLIDE oracle lesson: accept cases wait
        the full window). A found dialog is captured, dismissed by the
        watcher's compile-mode override, and reported as the verdict.
        """
        watcher = self.watcher
        if watcher is None:
            raise HostError("Watcher unavailable; cannot verify compile.")
        watcher.suppress_vbe_reporting(True)
        watcher.set_compile_mode(True)
        mark = watcher.mark()
        try:
            fired = self.host.start_compile()
            if fired == "already-compiled":
                return {"compile": "accepted", "signal": "already-compiled"}
            deadline = time.time() + watch_seconds
            while time.time() < deadline:
                for record in watcher.records_since(mark):
                    decision = record.decision
                    if (decision is not None
                            and decision.classification == "compile-error"):
                        return {
                            "compile": "rejected",
                            "dialog": _record_payload(record),
                        }
                # Positive completion: the VBE disables Compile once the
                # project is fully compiled, so a clean compile returns in
                # milliseconds instead of waiting out the whole window. A
                # failed compile leaves the control enabled, but sweep the
                # records once more so a dialog can never be outrun.
                if self.host.compile_control_disabled():
                    for record in watcher.records_since(mark):
                        decision = record.decision
                        if (decision is not None and
                                decision.classification == "compile-error"):
                            return {
                                "compile": "rejected",
                                "dialog": _record_payload(record),
                            }
                    return {"compile": "accepted",
                            "signal": "control-disabled"}
                time.sleep(0.05)
            return {"compile": "accepted", "signal": "watch-window-elapsed",
                    "watch_seconds": watch_seconds}
        finally:
            try:
                self.host.end_compile()
            except HostError:
                pass
            watcher.set_compile_mode(False)
            watcher.suppress_vbe_reporting(False)


def main() -> int:
    global _session_id
    parser = argparse.ArgumentParser()
    parser.add_argument("--session", default="unknown")
    parser.add_argument("--artifacts", default="")
    args = parser.parse_args()
    _session_id = args.session

    for stream in (sys.stdin, sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8")
            except (ValueError, OSError):
                pass

    # Diagnostic seam: when a command wedges, the supervisor kills this
    # process from outside and the reason is invisible. With this set, every
    # thread's stack is dumped to stderr (which the supervisor tails) at the
    # given interval, so a hang can be located instead of guessed at.
    debug_interval = os.environ.get("PYVBAHARNESS_STACK_DUMP_S")
    if debug_interval:
        try:
            faulthandler.dump_traceback_later(float(debug_interval),
                                              repeat=True, file=sys.stderr)
        except (ValueError, RuntimeError):
            pass

    worker = Worker(artifacts_dir=args.artifacts)
    try:
        worker.start()
    except Exception as err:  # noqa: BLE001 - fatal startup must be reported
        print(f"PYVBA_WORKER_FATAL|{err}", file=sys.stderr, flush=True)
        return 1

    exit_code = 0
    try:
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            try:
                command = protocol.decode_command(line)
            except ValueError as err:
                print(f"PYVBA_WORKER_FATAL|{err}", file=sys.stderr, flush=True)
                exit_code = 1
                break
            cid = command["cid"]
            name = command["cmd"]
            params = command["params"]
            timeout_ms = params.pop("_timeout_ms", None)
            _emit(protocol.EV_COMMAND_STARTED, {
                "cid": cid, "command": name, "timeout_ms": timeout_ms})
            started = time.time()
            if name == protocol.CMD_SHUTDOWN:
                _emit(protocol.EV_COMMAND_FINISHED, {
                    "cid": cid, "command": name, "outcome": PASSED,
                    "duration_ms": 0})
                break
            payload = worker.execute(name, params)
            payload.update({
                "cid": cid,
                "command": name,
                "duration_ms": int((time.time() - started) * 1000),
            })
            _emit(protocol.EV_COMMAND_FINISHED, payload)
    finally:
        worker.shutdown()
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
