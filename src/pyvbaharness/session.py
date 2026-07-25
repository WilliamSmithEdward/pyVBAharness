"""ExcelSession: the supervisor.

Runs no COM itself. It spawns the worker process, feeds it commands, consumes
its event stream, and enforces the three watchdog windows from outside the
COM apartment (XLIDE session pattern):

- startup: spawn until worker-ready
- per-command: every command carries a positive timeout
- cleanup: shutdown grace, because Close/Quit/COM-release can hang after the
  useful work finished

All failures funnel through one abort path: kill the recorded Excel PID,
kill the worker, record excel-killed, mark the session dead. A dead session
issues no further commands (oracle rule). A timeout or blocked modal is
reported as infrastructure state, never as evidence about the VBA code.
"""
from __future__ import annotations

import hashlib
import json
import os
import queue
import subprocess
import sys
import threading
import time
import uuid
import warnings
import weakref
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import codegen, protocol, vbasig
from .numbering import instrument_error_lines
from .lock import COMPILE_MUTEX_NAME, SessionLock
from .oracle import OracleIssue, validate_session_trace
from .process_control import (
    OwnedProcessManifest,
    is_process_alive,
    sweep_stale_manifests,
)
from .results import (
    COMPILE_ACCEPTED,
    COMPILE_INFRA_FAILURE,
    COMPILE_REJECTED,
    MODAL_BLOCKED,
    PASSED,
    RUNNER_ERROR,
    TIMEOUT,
    VBA_ERROR,
    CompileResult,
    DialogRecord,
    HarnessError,
    RunResult,
    SessionDead,
    TestCaseResult,
    VbaError,
    WorkerProtocolError,
)

_CREATE_NO_WINDOW = 0x08000000
USER_MODULE_NAME = "PyVbaUserCode"


@dataclass
class HarnessConfig:
    startup_timeout_s: float = 60.0
    default_timeout_s: float = 30.0
    cleanup_grace_s: float = 5.0
    compile_watch_s: float = 10.0
    exclusive: bool = True
    lock_wait_s: float = 0.0
    auto_recycle: bool = True
    manifest_dir: Path | None = None
    python_executable: str = sys.executable
    extra_env: dict[str, str] = field(default_factory=dict)
    # Test seam: replaces "<python> -m pyvbaharness.worker" so the supervisor
    # state machine can be exercised against a scripted fake worker.
    worker_argv: list[str] | None = None


class ExcelSession:
    """One owned, watchdogged Excel instance reusable across many runs."""

    def __init__(self, config: HarnessConfig | None = None) -> None:
        self.config = config or HarnessConfig()
        self.events: list[dict[str, Any]] = []
        self.oracle_issues: list[OracleIssue] = []
        self._lock: SessionLock | None = None
        self._proc: subprocess.Popen | None = None
        self._manifest: OwnedProcessManifest | None = None
        self._queue: "queue.Queue[tuple[str, Any]]" = queue.Queue()
        self._stderr_tail: deque[str] = deque(maxlen=200)
        self._cid = 0
        self._dead = True
        self._closed = False
        self._injected: dict[str, str] = {}
        self._readers: list[threading.Thread] = []
        self._finalizer: weakref.finalize | None = None
        self.session_id = uuid.uuid4().hex[:12]
        if self.config.exclusive:
            self._lock = SessionLock(timeout_s=self.config.lock_wait_s)
            self._lock.acquire()
        try:
            for note in sweep_stale_manifests(self.config.manifest_dir):
                warnings.warn(f"pyvbaharness sweep: {note}", stacklevel=2)
            self._start()
        except BaseException:
            if self._lock is not None:
                self._lock.release()
            raise

    # ----- lifecycle -------------------------------------------------------

    def _start(self) -> None:
        # One manifest per logical session, reused across recycles, so a
        # recycle cannot orphan a manifest file that the sweep would later
        # have to clean up.
        if not self.session_id:
            self.session_id = uuid.uuid4().hex[:12]
        if self._manifest is None:
            self._manifest = OwnedProcessManifest(self.session_id,
                                                  self.config.manifest_dir)
        env = dict(os.environ)
        src_root = str(Path(__file__).resolve().parents[1])
        env["PYTHONPATH"] = os.pathsep.join(
            p for p in (src_root, env.get("PYTHONPATH")) if p)
        env["PYTHONIOENCODING"] = "utf-8"
        env.update(self.config.extra_env)
        argv = (list(self.config.worker_argv) if self.config.worker_argv
                else [self.config.python_executable, "-m",
                      "pyvbaharness.worker"])
        self._proc = subprocess.Popen(
            argv + ["--session", self.session_id],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=_CREATE_NO_WINDOW,
            env=env,
        )
        self._manifest.record("worker", self._proc.pid)
        self._queue = queue.Queue()
        self._stderr_tail.clear()
        self._injected.clear()
        self._readers = [
            threading.Thread(target=self._read_stdout, daemon=True,
                             name="pyvba-stdout"),
            threading.Thread(target=self._read_stderr, daemon=True,
                             name="pyvba-stderr"),
        ]
        for reader in self._readers:
            reader.start()
        self._dead = False
        self._closed = False
        # Crash safety net: if this session object is garbage collected or
        # the interpreter exits without close(), kill the recorded processes
        # via the manifest. Detached on a clean close. weakref.finalize must
        # not capture self, so it works from the manifest alone.
        if self._finalizer is not None:
            self._finalizer.detach()
        self._finalizer = weakref.finalize(
            self, _emergency_cleanup, self._manifest)
        self._await_ready()

    def _read_stdout(self) -> None:
        proc = self._proc
        assert proc is not None and proc.stdout is not None
        try:
            for line in proc.stdout:
                line = line.rstrip("\r\n")
                if not line:
                    continue
                try:
                    event = protocol.decode_event(line)
                except ValueError:
                    event = None
                if event is not None:
                    self._queue.put(("event", event))
                else:
                    self._queue.put(("log", line))
        except (OSError, ValueError):
            # A hard-killed worker breaks the pipe mid-read; that IS the
            # exit signal, not an error worth surfacing.
            pass
        self._queue.put(("exit", proc.wait()))

    def _read_stderr(self) -> None:
        proc = self._proc
        assert proc is not None and proc.stderr is not None
        try:
            for line in proc.stderr:
                self._stderr_tail.append(line.rstrip("\r\n"))
        except (OSError, ValueError):
            pass

    def _await_ready(self) -> None:
        deadline = time.monotonic() + self.config.startup_timeout_s
        while True:
            item = self._next_item(deadline)
            if item is None:
                self._abort("timeout")
                raise HarnessError(
                    f"Worker startup exceeded "
                    f"{self.config.startup_timeout_s:.0f}s. Cold Office "
                    "first-run can be slow once; retry before diagnosing.")
            kind, payload = item
            if kind == "exit":
                self._abort("runner-error")
                raise HarnessError(
                    "Worker exited during startup: " + self._stderr_summary())
            if kind == "event":
                self._ingest(payload)
                if payload.get("kind") == protocol.EV_READY:
                    return
                if payload.get("kind") in (protocol.EV_MODAL_BLOCKED,
                                           protocol.EV_VBE_WINDOW):
                    self._abort(MODAL_BLOCKED)
                    raise HarnessError(
                        "Blocked by a modal dialog during startup: "
                        + str(payload.get("title", "")))

    def close(self) -> None:
        """Cooperative shutdown with a hard cleanup deadline."""
        if self._closed:
            return
        self._closed = True
        try:
            if not self._dead and self._proc is not None:
                try:
                    self._write_command(
                        protocol.CMD_SHUTDOWN, {},
                        int(self.config.cleanup_grace_s * 1000))
                except (OSError, ValueError):
                    self._abort("runner-error")
                else:
                    grace = self.config.cleanup_grace_s
                    deadline = time.monotonic() + grace + 10.0
                    exited = self._drain_until_exit(deadline)
                    if not exited:
                        self._abort("cleanup-failed")
                self._dead = True
        finally:
            self._ensure_owned_excel_gone()
            # Join the pipe readers so a session's file objects are not
            # garbage collected while a thread still reads them (noisy
            # OSError at interpreter teardown otherwise).
            for reader in self._readers:
                reader.join(timeout=2.0)
            self._close_pipes()
            if self._finalizer is not None:
                self._finalizer.detach()
                self._finalizer = None
            if self._manifest is not None:
                self._manifest.remove()
            if self._lock is not None:
                self._lock.release()
                self._lock = None
            self.oracle_issues = validate_session_trace(self.events)
            for issue in self.oracle_issues:
                warnings.warn(
                    f"pyvbaharness trace oracle: {issue.code}: "
                    f"{issue.message}", stacklevel=2)

    def _ensure_owned_excel_gone(self) -> None:
        """Verify the owned Excel process actually ended; kill it if not.

        A successful Quit() is not proof of termination. Excel that has been
        made visible at any point (a compile check does exactly that) treats
        itself as user-launched and keeps running after its last automation
        client disconnects, leaving a hidden orphan. Observed on 2026-07-25:
        a live run ended with every harness process gone and one responding
        Excel still alive.
        """
        if self._manifest is None:
            return
        entry = self._manifest.entry("excel")
        if entry is None:
            return
        pid, _creation = entry
        deadline = time.monotonic() + self.config.cleanup_grace_s
        while time.monotonic() < deadline:
            if not is_process_alive(pid):
                return
            time.sleep(0.2)
        if self._manifest.kill_role("excel"):
            self.events.append({
                "kind": protocol.EV_EXCEL_KILLED,
                "session": self.session_id,
                "reason": "quit-did-not-terminate",
                "killed": True,
                "at": time.time(),
            })

    def _drain_until_exit(self, deadline: float) -> bool:
        while True:
            item = self._next_item(deadline)
            if item is None:
                return False
            kind, payload = item
            if kind == "exit":
                return True
            if kind == "event":
                self._ingest(payload)

    def recycle(self) -> None:
        """Kill whatever remains and start a fresh worker + Excel.

        State (open workbook, injected modules) is NOT replayed; callers
        reopen and reinject. Recycling exists so one hang cannot poison
        subsequent runs.
        """
        if not self._dead:
            self._abort("recycled")
        self._start()

    def __enter__(self) -> "ExcelSession":
        return self

    def __exit__(self, *_exc: Any) -> None:
        self.close()

    # ----- abort path ------------------------------------------------------

    def _abort(self, reason: str) -> None:
        """Single failure path: kill owned Excel, kill worker, mark dead."""
        if self._dead:
            return
        self._dead = True
        killed = False
        if self._manifest is not None:
            killed = self._manifest.kill_role("excel")
        proc = self._proc
        if proc is not None and proc.poll() is None:
            proc.kill()
        self.events.append({
            "kind": protocol.EV_EXCEL_KILLED,
            "session": self.session_id,
            "reason": reason,
            "killed": killed,
            "at": time.time(),
        })

    def _stderr_summary(self) -> str:
        return "\n".join(list(self._stderr_tail)[-10:]) or "(no stderr)"

    def _close_pipes(self) -> None:
        """Close the worker's pipe objects with errors suppressed.

        After a hard-killed worker, a TextIOWrapper flushing at garbage
        collection raises OSError 22 on the broken pipe; closing eagerly
        keeps that out of interpreter-teardown noise.
        """
        proc = self._proc
        if proc is None:
            return
        for stream in (proc.stdin, proc.stdout, proc.stderr):
            if stream is None:
                continue
            try:
                stream.close()
            except (OSError, ValueError):
                pass

    # ----- command plumbing ------------------------------------------------

    def _ingest(self, event: dict[str, Any]) -> None:
        self.events.append(event)
        if (event.get("kind") == protocol.EV_EXCEL_CREATED
                and self._manifest is not None):
            pid = int(event.get("pid") or 0)
            if pid > 0:
                self._manifest.record("excel", pid)

    def _next_item(self, deadline: float) -> tuple[str, Any] | None:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return None
        try:
            return self._queue.get(timeout=remaining)
        except queue.Empty:
            return None

    def _write_command(self, name: str, params: dict[str, Any],
                       timeout_ms: int | None) -> int:
        proc = self._proc
        if proc is None or proc.stdin is None:
            raise SessionDead("No worker process.")
        self._cid += 1
        payload = dict(params)
        payload["_timeout_ms"] = timeout_ms
        proc.stdin.write(protocol.encode_command(self._cid, name, payload)
                         + "\n")
        proc.stdin.flush()
        return self._cid

    def _command(self, name: str, params: dict[str, Any],
                 timeout_s: float | None) -> dict[str, Any]:
        """Send one command and wait for its result under the watchdog.

        Returns the command-finished payload. Timeout / modal-block / worker
        death do not raise here; they return a synthesized payload with the
        matching infrastructure outcome, and the session is already dead.
        """
        if self._dead:
            if self.config.auto_recycle:
                self._start()
            else:
                raise SessionDead(
                    "Session was killed after a previous failure; call "
                    "recycle() or enable auto_recycle.")
        timeout_s = timeout_s or self.config.default_timeout_s
        try:
            cid = self._write_command(name, params, int(timeout_s * 1000))
        except OSError as err:
            result = self._synthesize(
                name, RUNNER_ERROR, f"Could not reach the worker: {err}")
            self._abort("runner-error")
            return result
        deadline = time.monotonic() + timeout_s
        while True:
            item = self._next_item(deadline)
            if item is None:
                # Terminal result first, then the kill, so the trace reads
                # "hang reported -> owned Excel killed" (oracle ordering).
                result = self._synthesize(
                    name, TIMEOUT, f"Timed out after {timeout_s:.0f}s.")
                self._abort(TIMEOUT)
                return result
            kind, payload = item
            if kind == "exit":
                result = self._synthesize(
                    name, RUNNER_ERROR,
                    "Worker exited mid-command: " + self._stderr_summary())
                self._abort("runner-error")
                return result
            if kind == "log":
                continue
            event = payload
            self._ingest(event)
            event_kind = event.get("kind")
            if event_kind == protocol.EV_MODAL_BLOCKED and name != protocol.CMD_COMPILE:
                result = self._synthesize(
                    name, MODAL_BLOCKED,
                    "Blocked by a modal dialog: "
                    + str(event.get("title") or event.get("message") or ""),
                    dialogs=[event])
                self._abort(MODAL_BLOCKED)
                return result
            if event_kind == protocol.EV_VBE_WINDOW:
                result = self._synthesize(
                    name, MODAL_BLOCKED,
                    "The VBE window took over during a hidden run "
                    "(debugger break).", dialogs=[event])
                self._abort(MODAL_BLOCKED)
                return result
            if (event_kind == protocol.EV_COMMAND_FINISHED
                    and event.get("cid") == cid):
                return event

    def _synthesize(self, command: str, outcome: str, message: str,
                    dialogs: list[dict[str, Any]] | None = None
                    ) -> dict[str, Any]:
        event = {
            "kind": protocol.EV_COMMAND_FINISHED,
            "session": self.session_id,
            "command": command,
            "outcome": outcome,
            "message": message,
            "synthesized": True,
            "at": time.time(),
        }
        if dialogs:
            event["dialogs"] = dialogs
        self.events.append(event)
        return event

    @staticmethod
    def _dialog_records(payload: dict[str, Any]) -> list[DialogRecord]:
        records = []
        for raw in payload.get("dialogs", []) or []:
            records.append(DialogRecord(
                title=str(raw.get("title", "")),
                message=str(raw.get("message", "")),
                texts=[str(t) for t in raw.get("texts", []) or []],
                buttons=[str(b) for b in raw.get("buttons", []) or []],
                button_ids=[int(i) for i in raw.get("button_ids", []) or []],
                classification=str(raw.get("classification", "excel-modal")),
                action=str(raw.get("action", "none")),
            ))
        return records

    def _expect_passed(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Infrastructure commands must pass; anything else is an error."""
        outcome = payload.get("outcome")
        if outcome == PASSED:
            return payload.get("data") or {}
        message = str(payload.get("message", "")) or str(outcome)
        if outcome in (TIMEOUT, MODAL_BLOCKED):
            raise SessionDead(
                f"{payload.get('command')}: {outcome}: {message}")
        raise HarnessError(f"{payload.get('command')}: {message}")

    # ----- public API ------------------------------------------------------

    def ping(self) -> dict[str, Any]:
        return self._expect_passed(
            self._command(protocol.CMD_PING, {}, 10.0))

    def new_workbook(self) -> dict[str, Any]:
        self._injected.clear()
        return self._expect_passed(
            self._command(protocol.CMD_NEW_WORKBOOK, {}, None))

    def open_workbook(self, path: str | Path,
                      read_only: bool = True,
                      timeout: float | None = None) -> dict[str, Any]:
        resolved = str(Path(path).resolve())
        self._injected.clear()
        return self._expect_passed(self._command(
            protocol.CMD_OPEN_WORKBOOK,
            {"path": resolved, "read_only": read_only}, timeout))

    def add_module(self, name: str, source: str, kind: str = "standard",
                   line_numbers: bool = False) -> dict[str, Any]:
        """Create or replace a module.

        With ``line_numbers=True`` the executable body lines are numbered so
        a VBA error reports its source line (``result.error.line``). Injection
        is skipped when this session already injected identical source into
        the current workbook, which makes repeat run_vba calls cost a warm
        run instead of a module replacement.
        """
        codegen.validate_module_name(name)
        if kind not in ("standard", "class"):
            raise ValueError("kind must be 'standard' or 'class'")
        if line_numbers:
            source = instrument_error_lines(source)
        digest = hashlib.sha256(
            f"{kind}\x00{source}".encode("utf-8", "surrogatepass")
        ).hexdigest()
        cache_key = name.lower()
        if self._injected.get(cache_key) == digest:
            return {"name": name, "kind": kind, "cached": True}
        data = self._expect_passed(self._command(
            protocol.CMD_ADD_MODULE,
            {"name": name, "source": source, "kind": kind}, None))
        self._injected[cache_key] = digest
        return data

    def remove_module(self, name: str) -> dict[str, Any]:
        self._injected.pop(name.lower(), None)
        return self._expect_passed(self._command(
            protocol.CMD_REMOVE_MODULE, {"name": name}, None))

    def list_procedures(self, module: str) -> list[dict[str, Any]]:
        """Callable procedures declared in a module: name, kind, arity."""
        data = self._expect_passed(self._command(
            protocol.CMD_LIST_PROCS, {"module": module}, None))
        return data.get("procs", [])

    def run_macro(self, target: str, *args: Any,
                  timeout: float | None = None) -> RunResult:
        codegen.validate_run_target(target)
        started = time.monotonic()
        payload = self._command(protocol.CMD_RUN,
                                {"target": target, "args": list(args)},
                                timeout)
        duration = time.monotonic() - started
        outcome = str(payload.get("outcome", RUNNER_ERROR))
        error = None
        if outcome == VBA_ERROR:
            raw_line = int(payload.get("line", 0) or 0)
            error = VbaError(
                number=int(payload.get("number", 0)),
                source=str(payload.get("source", "")),
                description=str(payload.get("description", "")),
                line=raw_line if raw_line > 0 else None)
        return RunResult(
            outcome=outcome,
            duration_s=duration,
            value=payload.get("value"),
            output=[str(x) for x in payload.get("output", []) or []],
            error=error,
            message=str(payload.get("message", "")),
            dialogs=self._dialog_records(payload),
        )

    def run_vba(self, source: str, proc: str = "Main", args: tuple = (),
                timeout: float | None = None,
                module_name: str = USER_MODULE_NAME,
                line_numbers: bool = True) -> RunResult:
        """Inject ``source`` as a module (replacing any prior one) and run
        ``proc`` from it. If no workbook is open, an unsaved in-memory
        workbook is created.

        Line numbering is on by default so ``result.error.line`` reports the
        failing source line; identical source is not reinjected, so calling
        this in a loop costs a warm run, not a module replacement.
        """
        if not self._workbook_known():
            self.new_workbook()
        self.add_module(module_name, source, line_numbers=line_numbers)
        return self.run_macro(f"{module_name}.{proc}", *args, timeout=timeout)

    def eval(self, expression: str, timeout: float | None = None) -> Any:
        """Evaluate one VBA expression and return its value.

        The expression is wrapped in a generated Function, so anything legal
        on the right-hand side of an assignment works:
        ``excel.eval("WorksheetFunction.Sum(1, 2, 3)")`` returns 6.0.
        Raises HarnessError when the expression raises a VBA error.
        """
        if "\n" in expression or "\r" in expression:
            raise ValueError("eval takes a single-line VBA expression; "
                             "use run_vba for statements.")
        source = ("Public Function PyVbaEvalResult() As Variant\n"
                  f"    PyVbaEvalResult = ({expression})\n"
                  "End Function\n")
        result = self.run_vba(source, proc="PyVbaEvalResult", timeout=timeout,
                              module_name="PyVbaEvalCode",
                              line_numbers=False)
        if result.outcome == PASSED:
            return result.value
        if result.error is not None:
            raise HarnessError(f"eval({expression!r}): {result.error}")
        raise SessionDead(f"eval({expression!r}): {result.outcome}: "
                          f"{result.message}")

    def run_tests(self, source: str | None = None,
                  module: str | None = None, prefix: str = "Test",
                  timeout: float | None = None,
                  tests: list[str] | None = None) -> list[TestCaseResult]:
        """Discover and run VBA test procedures.

        Tests are zero-argument Subs or Functions whose names start with
        ``prefix``, either in ``source`` (injected as ``module``, default
        PyVbaTests) or in an existing module. ``tests`` names an explicit
        subset instead of prefix discovery (SessionPool sharding uses this).
        Inside tests, ``PyVbaAssert`` and ``PyVbaAssertEqual`` raise
        structured failures, and ``PyVbaLog`` output comes back per test.

        A hanging test costs that test, not the suite: when the session dies
        (timeout or blocked modal) and ``source`` is available, the session
        recycles, the module is reinjected, and the remaining tests still
        run. Without ``source`` to reinject, the remainder is reported as
        runner-error "not run" rather than silently dropped.
        """
        if source is not None:
            module = module or "PyVbaTests"
            self._install_test_source(module, source)
            available = vbasig.discover_tests(source, prefix="")
        elif module is None:
            raise ValueError("run_tests needs source or a module name")
        else:
            available = [
                p["name"] for p in self.list_procedures(module)
                if p["kind"] in ("sub", "function") and p["required"] == 0
            ]
        if tests is None:
            wanted = prefix.lower()
            selected = [n for n in available
                        if n.lower().startswith(wanted)]
        else:
            lookup = {n.lower(): n for n in available}
            missing = [t for t in tests if t.lower() not in lookup]
            if missing:
                raise ValueError(
                    f"Unknown test procedure(s): {', '.join(missing)}")
            selected = [lookup[t.lower()] for t in tests]

        results: list[TestCaseResult] = []
        for index, name in enumerate(selected):
            run = self.run_macro(f"{module}.{name}", timeout=timeout)
            results.append(TestCaseResult(name=name, result=run))
            if run.outcome not in (TIMEOUT, MODAL_BLOCKED, RUNNER_ERROR):
                continue
            if run.outcome == RUNNER_ERROR and not self._dead:
                # An isolated harness-level failure with a healthy session
                # (for example a target that stopped existing) does not end
                # the suite.
                continue
            if source is not None:
                try:
                    self._install_test_source(module, source)
                    continue  # recovered: keep running the remaining tests
                except (HarnessError, SessionDead):
                    pass
            for skipped in selected[index + 1:]:
                results.append(TestCaseResult(
                    name=skipped,
                    result=RunResult(
                        outcome=RUNNER_ERROR, duration_s=0.0,
                        message=f"Not run: the session died on {name} "
                                f"({run.outcome}).")))
            break
        return results

    def _install_test_source(self, module: str, source: str) -> None:
        if not self._workbook_known():
            self.new_workbook()
        self.add_module(module, source, line_numbers=True)

    def save_trace(self, path: str | Path) -> Path:
        """Write the session's event trace as JSON lines for postmortems."""
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("w", encoding="utf-8") as handle:
            for event in self.events:
                handle.write(json.dumps(event, ensure_ascii=False,
                                        default=str) + "\n")
        return target

    def _workbook_known(self) -> bool:
        for event in reversed(self.events):
            kind = event.get("kind")
            if kind in (protocol.EV_WORKBOOK_CREATED,
                        protocol.EV_WORKBOOK_OPENED):
                return True
            if kind in (protocol.EV_WORKBOOK_CLOSED, protocol.EV_EXCEL_KILLED,
                        protocol.EV_EXCEL_CREATED):
                return False
        return False

    def run_raw(self, target: str, *args: Any,
                timeout: float | None = None) -> Any:
        """Direct Application.Run without the injected runner.

        No in-VBA error trapping: a VBA error surfaces as a COM error and
        raises HarnessError here. Prefer run_macro for managed execution.
        """
        codegen.validate_run_target(target)
        data = self._expect_passed(self._command(
            protocol.CMD_RUN_RAW, {"target": target, "args": list(args)},
            timeout))
        return data.get("value")

    @property
    def is_dead(self) -> bool:
        """True while the session has no live worker (recycle pending)."""
        return self._dead

    @property
    def excel_pid(self) -> int:
        """PID of the owned Excel instance (0 before excel-created)."""
        for event in reversed(self.events):
            if event.get("kind") == protocol.EV_EXCEL_CREATED:
                return int(event.get("pid") or 0)
        return 0

    def read_range(self, sheet: str, ref: str,
                   timeout: float | None = None) -> list[list[Any]]:
        data = self._expect_passed(self._command(
            protocol.CMD_READ_RANGE, {"sheet": sheet, "ref": ref}, timeout))
        return data.get("data", [])

    def write_range(self, sheet: str, start_cell: str,
                    data: list[list[Any]],
                    timeout: float | None = None) -> dict[str, Any]:
        return self._expect_passed(self._command(
            protocol.CMD_WRITE_RANGE,
            {"sheet": sheet, "start_cell": start_cell, "data": data},
            timeout))

    def save_as(self, path: str | Path,
                timeout: float | None = None) -> dict[str, Any]:
        return self._expect_passed(self._command(
            protocol.CMD_SAVE_AS, {"path": str(Path(path).resolve())},
            timeout))

    def compile_project(self, watch_seconds: float | None = None,
                        include_harness_support: bool = False
                        ) -> CompileResult:
        """VBE compile check. Excel becomes visible for its duration.

        A clean compile usually returns fast: the VBE disables its Compile
        command once the project is compiled, which is a positive completion
        signal. The watch window is only waited out when neither that signal
        nor a compile-error dialog appears; size it generously on cold
        machines (the first compile in a fresh VBE can exceed 10 s).

        ``include_harness_support=True`` injects the harness support module
        first, so code that calls PyVbaLog or the assert helpers compiles the
        way it will actually run. Leave it False to check a workbook exactly
        as-is.

        Compile checks are serialized machine-wide (dedicated mutex), even
        for pool sessions: they drive the visible VBE, which is a genuinely
        shared UI surface. Hidden runs in other sessions continue in
        parallel.
        """
        watch = watch_seconds or self.config.compile_watch_s
        started = time.monotonic()
        compile_lock = SessionLock(timeout_s=watch + 120.0,
                                   name=COMPILE_MUTEX_NAME,
                                   purpose="compile")
        compile_lock.acquire()
        try:
            payload = self._command(
                protocol.CMD_COMPILE,
                {"watch_seconds": watch,
                 "include_support": include_harness_support},
                watch + 30.0)
        finally:
            compile_lock.release()
        duration = time.monotonic() - started
        outcome = payload.get("outcome")
        if outcome == PASSED:
            data = payload.get("data") or {}
            dialog_payload = data.get("dialog")
            if data.get("compile") == COMPILE_REJECTED:
                dialogs = self._dialog_records({"dialogs": [dialog_payload]})
                return CompileResult(
                    outcome=COMPILE_REJECTED, duration_s=duration,
                    dialog=dialogs[0] if dialogs else None,
                    message=str((dialog_payload or {}).get("message", "")))
            return CompileResult(outcome=COMPILE_ACCEPTED,
                                 duration_s=duration)
        return CompileResult(
            outcome=COMPILE_INFRA_FAILURE, duration_s=duration,
            message=str(payload.get("message", ""))
            or "Compile check could not run to completion; the verdict is "
               "unknown (timeout is never evidence).")


def _emergency_cleanup(manifest: OwnedProcessManifest) -> None:
    """Finalizer body: kill recorded processes when a session was never
    closed (garbage collection or interpreter exit). Must not reference the
    session object."""
    try:
        manifest.kill_role("excel")
        manifest.kill_role("worker")
        manifest.remove()
    except Exception:
        pass


def run_vba(source: str, proc: str = "Main", args: tuple = (),
            timeout: float | None = None,
            config: HarnessConfig | None = None) -> RunResult:
    """One-shot convenience: fresh session, unsaved workbook, inject, run,
    tear down. Full isolation at the cost of Excel startup (1-3 s warm)."""
    with ExcelSession(config) as session:
        return session.run_vba(source, proc=proc, args=args, timeout=timeout)
