"""SessionPool: parallel VBA execution across multiple owned Excel
instances.

Each member is a full ExcelSession: its own worker process, its own new
EXCEL.EXE (PID-verified, kill-on-close job), its own dialog watcher and
manifest. Members run with ``exclusive=False`` (the machine-wide session
mutex exists to stop ACCIDENTAL concurrency; the pool is deliberate
concurrency) and ``auto_recycle=True`` (a hang costs one member one recycle,
never the pool).

What stays serialized, on purpose:

- compile checks, machine-wide, via the dedicated compile mutex inside
  ``compile_project`` - they drive the visible VBE, a shared UI surface;
- each member session, internally - one command at a time per session, which
  the pool guarantees by checking a session out to exactly one task.

Hidden macro runs and range IO parallelize cleanly: dialog handling is
PID-scoped window enumeration plus messages sent directly to button handles,
so one instance's dialogs cannot interfere with another's.

Budget roughly 150-300 MB of RAM per member; see
benchmarks/run_pool_benchmarks.py for where throughput flattens on a given
machine.
"""
from __future__ import annotations

import dataclasses
import threading
import queue
from concurrent.futures import Future, ThreadPoolExecutor
from typing import Any, Callable

from . import vbasig
from .results import HarnessError, RunResult, TestCaseResult
from .session import ExcelSession, HarnessConfig


class SessionPool:
    """N concurrently usable ExcelSessions behind one work queue.

    ``submit(task)`` schedules ``task(session)`` on a checked-out member and
    returns a Future; ``run_vba`` is the common case; ``run_tests`` shards a
    test suite across all members and merges the results in discovery order.
    """

    def __init__(self, size: int = 2,
                 config: HarnessConfig | None = None) -> None:
        if size < 1:
            raise ValueError("SessionPool size must be at least 1")
        base = config or HarnessConfig()
        member_config = dataclasses.replace(
            base, exclusive=False, auto_recycle=True)
        self.size = size
        self._closed = False
        self._executor = ThreadPoolExecutor(
            max_workers=size, thread_name_prefix="pyvba-pool")
        self._idle: "queue.Queue[ExcelSession]" = queue.Queue()
        self._sessions: list[ExcelSession] = []

        # Start members concurrently: Excel launches are independent
        # processes, and N sequential warm starts would cost N x 0.5-3 s.
        startups = [self._executor.submit(ExcelSession, member_config)
                    for _ in range(size)]
        failures: list[BaseException] = []
        for startup in startups:
            try:
                self._sessions.append(startup.result())
            except BaseException as err:  # noqa: BLE001 - collected, re-raised
                failures.append(err)
        if failures:
            for session in self._sessions:
                try:
                    session.close()
                except Exception:
                    pass
            self._executor.shutdown(wait=False)
            raise HarnessError(
                f"{len(failures)} of {size} pool sessions failed to start: "
                f"{failures[0]}") from failures[0]
        for session in self._sessions:
            self._idle.put(session)

    # ----- scheduling ------------------------------------------------------

    def submit(self, task: Callable[[ExcelSession], Any]) -> Future:
        """Run ``task(session)`` on the next free member; returns a Future.

        The session is exclusively checked out to the task for its duration,
        so the task may call any session method, including multi-step flows
        (open a workbook, inject, run, read ranges).
        """
        if self._closed:
            raise HarnessError("The pool is closed.")
        return self._executor.submit(self._run_checked_out, task)

    def _run_checked_out(self, task: Callable[[ExcelSession], Any]) -> Any:
        session = self._idle.get()
        try:
            return task(session)
        finally:
            self._idle.put(session)

    def run_vba(self, source: str, proc: str = "Main", args: tuple = (),
                timeout: float | None = None,
                line_numbers: bool = True) -> "Future[RunResult]":
        """Parallel run_vba; returns a Future of the RunResult."""
        return self.submit(lambda session: session.run_vba(
            source, proc=proc, args=args, timeout=timeout,
            line_numbers=line_numbers))

    def run_tests(self, source: str, module: str = "PyVbaTests",
                  prefix: str = "Test",
                  timeout: float | None = None) -> list[TestCaseResult]:
        """Shard a VBA test suite across all members and merge the results.

        Discovery is pure Python (vbasig), so every member injects the same
        module and runs its slice. Results come back in discovery order. A
        hanging test costs its member one recycle-and-reinject; the other
        members keep running throughout.
        """
        names = vbasig.discover_tests(source, prefix=prefix)
        if not names:
            return []
        shards = [names[index::self.size] for index in range(self.size)]
        shards = [shard for shard in shards if shard]
        futures = [
            self.submit(lambda session, shard=shard: session.run_tests(
                source=source, module=module, tests=shard, timeout=timeout))
            for shard in shards
        ]
        by_name: dict[str, TestCaseResult] = {}
        errors: list[BaseException] = []
        for future in futures:
            try:
                for case in future.result():
                    by_name[case.name] = case
            except BaseException as err:  # noqa: BLE001 - surfaced below
                errors.append(err)
        if errors:
            raise HarnessError(
                f"{len(errors)} test shard(s) failed: {errors[0]}"
            ) from errors[0]
        return [by_name[name] for name in names if name in by_name]

    # ----- introspection ---------------------------------------------------

    @property
    def excel_pids(self) -> list[int]:
        return [session.excel_pid for session in self._sessions]

    # ----- lifecycle -------------------------------------------------------

    def close(self) -> None:
        """Finish in-flight tasks, then close every member concurrently."""
        if self._closed:
            return
        self._closed = True
        self._executor.shutdown(wait=True)
        closers = [threading.Thread(target=session.close, daemon=True)
                   for session in self._sessions]
        for closer in closers:
            closer.start()
        for closer in closers:
            closer.join(timeout=60.0)

    def __enter__(self) -> "SessionPool":
        return self

    def __exit__(self, *_exc: Any) -> None:
        self.close()
