"""Live SessionPool tests: real parallel Excel instances.

Run with:  python -m pytest tests/live -m live -o addopts="" -v

The busy-wait workloads prove genuine concurrency by wall clock: two 2 s
VBA runs finishing in well under 4 s can only have run in parallel.
"""
import time
import warnings

import pytest

from pyvbaharness import (
    COMPILE_ACCEPTED,
    PASSED,
    TIMEOUT,
    HarnessConfig,
    SessionPool,
)
from pyvbaharness.process_control import is_process_alive

pytestmark = pytest.mark.live


def busy_source(seconds: float) -> str:
    return f"""
Public Sub Busy()
    Dim started As Single
    started = Timer
    Do While Timer - started < {seconds}!
    Loop
End Sub
"""


class TestParallelExecution:
    def test_two_busy_runs_overlap(self):
        with SessionPool(2) as pool:
            started = time.monotonic()
            futures = [pool.run_vba(busy_source(2.0), proc="Busy",
                                    timeout=30.0) for _ in range(2)]
            outcomes = [f.result(timeout=60).outcome for f in futures]
            elapsed = time.monotonic() - started
        assert outcomes == [PASSED, PASSED]
        # Serial execution would take >= 4 s; parallel finishes in ~2 s.
        assert elapsed < 3.4, f"no overlap: {elapsed:.1f}s for two 2s runs"

    def test_members_use_distinct_excel_processes(self):
        with SessionPool(3) as pool:
            pids = pool.excel_pids
            assert len(set(pids)) == 3
            assert all(pid > 0 for pid in pids)
            futures = [pool.submit(lambda s: s.eval("2 + 2"))
                       for _ in range(6)]
            assert [f.result(timeout=60) for f in futures] == [4] * 6

    def test_hang_in_one_member_does_not_stall_the_other(self):
        hang = """
Public Sub Freeze()
    Do
    Loop
End Sub
"""
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with SessionPool(2) as pool:
                hung = pool.run_vba(hang, proc="Freeze", timeout=6.0)
                started = time.monotonic()
                quick = pool.run_vba(busy_source(0.2), proc="Busy",
                                     timeout=30.0)
                assert quick.result(timeout=30).outcome == PASSED
                assert time.monotonic() - started < 5.0
                assert hung.result(timeout=60).outcome == TIMEOUT
                # The hung member recycled; the pool remains at capacity.
                after = [pool.run_vba(busy_source(0.1), proc="Busy",
                                      timeout=30.0) for _ in range(4)]
                assert all(f.result(timeout=60).outcome == PASSED
                           for f in after)


class TestShardedSuite:
    def test_suite_with_hang_completes_everywhere(self):
        suite = """
Public Sub TestAlpha()
    PyVbaAssertEqual 2, 1 + 1
End Sub

Public Sub TestHangs()
    Do
    Loop
End Sub

Public Sub TestBeta()
    PyVbaAssertEqual 6, 2 * 3
End Sub

Public Sub TestGamma()
    PyVbaLog "gamma ran"
End Sub
"""
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with SessionPool(2) as pool:
                results = pool.run_tests(suite, timeout=6.0)
        by_name = {r.name: r for r in results}
        assert sorted(by_name) == ["TestAlpha", "TestBeta", "TestGamma",
                                   "TestHangs"]
        assert by_name["TestHangs"].result.outcome == TIMEOUT
        # Recovery: every other test really ran; nothing was marked not-run.
        assert by_name["TestAlpha"].passed
        assert by_name["TestBeta"].passed
        assert by_name["TestGamma"].passed
        assert by_name["TestGamma"].result.output == ["gamma ran"]


class TestCompileSerialization:
    def test_concurrent_compiles_both_complete(self):
        source = """
Option Explicit
Public Function Fine() As Long
    Fine = 1
End Function
"""

        def compile_task(session):
            session.new_workbook()
            session.add_module("GoodMod", source)
            return session.compile_project(watch_seconds=10.0)

        with SessionPool(2) as pool:
            futures = [pool.submit(compile_task) for _ in range(2)]
            results = [f.result(timeout=120) for f in futures]
        assert [r.outcome for r in results] == [COMPILE_ACCEPTED] * 2


class TestPoolTeardown:
    def test_all_excels_die_with_the_pool(self):
        pool = SessionPool(2)
        pids = pool.excel_pids
        assert all(is_process_alive(pid) for pid in pids)
        pool.close()
        deadline = time.monotonic() + 20.0
        while time.monotonic() < deadline and any(
                is_process_alive(pid) for pid in pids):
            time.sleep(0.5)
        assert not any(is_process_alive(pid) for pid in pids)
