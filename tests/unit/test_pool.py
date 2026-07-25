"""SessionPool behavior against the scripted fake worker (no Excel).

Covers scheduling, isolation of a hang to one member, sharded run_tests
with recovery, and lifecycle. Real-Excel parallelism is covered by
tests/live/test_live_pool.py.
"""
import sys
import time
import warnings
from pathlib import Path

import pytest

from pyvbaharness import (
    PASSED,
    RUNNER_ERROR,
    TIMEOUT,
    HarnessConfig,
    HarnessError,
    SessionPool,
)

FAKE_WORKER = str(Path(__file__).with_name("fake_worker.py"))


def fake_config(tmp_path, **overrides) -> HarnessConfig:
    defaults = dict(
        worker_argv=[sys.executable, FAKE_WORKER],
        exclusive=False,
        manifest_dir=tmp_path,
        startup_timeout_s=15.0,
        default_timeout_s=5.0,
        cleanup_grace_s=3.0,
    )
    defaults.update(overrides)
    return HarnessConfig(**defaults)


SUITE = """
Public Sub TestOne()
End Sub

Public Sub TestTwo()
End Sub

Public Sub TestThree()
End Sub

Public Sub TestFour()
End Sub
"""


class TestScheduling:
    def test_submit_runs_and_returns_futures(self, tmp_path):
        with SessionPool(2, fake_config(tmp_path)) as pool:
            futures = [pool.submit(lambda s: s.run_macro("Ok.Now"))
                       for _ in range(4)]
            outcomes = [f.result(timeout=30).outcome for f in futures]
        assert outcomes == [PASSED] * 4

    def test_tasks_do_not_serialize_behind_a_slow_one(self, tmp_path):
        with SessionPool(2, fake_config(tmp_path)) as pool:
            slow = pool.submit(lambda s: s.run_macro("Slow.One"))
            quick = pool.submit(lambda s: s.run_macro("Ok.Now"))
            started = time.monotonic()
            assert quick.result(timeout=5).outcome == PASSED
            assert time.monotonic() - started < 0.9  # not behind the 1s run
            assert slow.result(timeout=5).value == "slow"

    def test_distinct_member_sessions(self, tmp_path):
        with SessionPool(3, fake_config(tmp_path)) as pool:
            pids = pool.excel_pids
            assert len(pids) == 3
            assert len(set(pids)) == 3

    def test_closed_pool_rejects_work(self, tmp_path):
        pool = SessionPool(1, fake_config(tmp_path))
        pool.close()
        with pytest.raises(HarnessError):
            pool.submit(lambda s: None)


class TestHangIsolation:
    def test_hang_times_out_without_blocking_other_member(self, tmp_path):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with SessionPool(2, fake_config(tmp_path)) as pool:
                hang = pool.submit(
                    lambda s: s.run_macro("Hang.Forever", timeout=2.0))
                quick = pool.submit(lambda s: s.run_macro("Ok.Now"))
                assert quick.result(timeout=5).outcome == PASSED
                assert hang.result(timeout=30).outcome == TIMEOUT
                # The hung member recycles; the pool stays fully usable.
                after = [pool.submit(lambda s: s.run_macro("Ok.Now"))
                         for _ in range(3)]
                assert all(f.result(timeout=30).outcome == PASSED
                           for f in after)


class TestShardedRunTests:
    def test_results_merge_in_discovery_order(self, tmp_path):
        with SessionPool(2, fake_config(tmp_path)) as pool:
            results = pool.run_tests(SUITE)
        assert [r.name for r in results] == [
            "TestOne", "TestTwo", "TestThree", "TestFour"]
        assert all(r.passed for r in results)

    def test_empty_suite(self, tmp_path):
        with SessionPool(2, fake_config(tmp_path)) as pool:
            assert pool.run_tests("Public Sub Helper()\nEnd Sub\n") == []

    def test_hanging_test_recovers_and_rest_still_run(self, tmp_path):
        suite = ("Public Sub TestHangOne()\nEnd Sub\n"
                 "Public Sub TestTwo()\nEnd Sub\n"
                 "Public Sub TestThree()\nEnd Sub\n")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with SessionPool(2, fake_config(tmp_path)) as pool:
                results = pool.run_tests(suite, timeout=2.0)
        by_name = {r.name: r for r in results}
        assert sorted(by_name) == ["TestHangOne", "TestThree", "TestTwo"]
        assert by_name["TestHangOne"].result.outcome == TIMEOUT
        assert by_name["TestTwo"].result.outcome == PASSED
        assert by_name["TestThree"].result.outcome == PASSED


class TestSessionRecovery:
    def test_run_tests_recovers_after_hang_with_source(self, tmp_path):
        from pyvbaharness import ExcelSession

        suite = ("Public Sub TestHangOne()\nEnd Sub\n"
                 "Public Sub TestAfter()\nEnd Sub\n")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            session = ExcelSession(fake_config(tmp_path, auto_recycle=True))
            try:
                results = session.run_tests(suite, timeout=2.0)
            finally:
                session.close()
        assert [r.name for r in results] == ["TestHangOne", "TestAfter"]
        assert results[0].result.outcome == TIMEOUT
        assert results[1].result.outcome == PASSED

    def test_run_tests_marks_not_run_without_recycle(self, tmp_path):
        from pyvbaharness import ExcelSession

        suite = ("Public Sub TestHangOne()\nEnd Sub\n"
                 "Public Sub TestAfter()\nEnd Sub\n")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            session = ExcelSession(fake_config(tmp_path, auto_recycle=False))
            try:
                results = session.run_tests(suite, timeout=2.0)
            finally:
                session.close()
        assert results[0].result.outcome == TIMEOUT
        assert results[1].result.outcome == RUNNER_ERROR
        assert "Not run" in results[1].result.message

    def test_explicit_test_subset_validated(self, tmp_path):
        from pyvbaharness import ExcelSession

        session = ExcelSession(fake_config(tmp_path))
        try:
            with pytest.raises(ValueError):
                session.run_tests(SUITE, tests=["TestOne", "NoSuchTest"])
            subset = session.run_tests(SUITE, tests=["testfour", "TestTwo"])
            assert [r.name for r in subset] == ["TestFour", "TestTwo"]
        finally:
            session.close()


class TestLifecycle:
    def test_close_removes_all_manifests(self, tmp_path):
        pool = SessionPool(2, fake_config(tmp_path))
        pool.run_vba("Public Sub Main()\nEnd Sub\n").result(timeout=30)
        pool.close()
        assert list(tmp_path.glob("*.json")) == []

    def test_close_is_idempotent(self, tmp_path):
        pool = SessionPool(1, fake_config(tmp_path))
        pool.close()
        pool.close()

    def test_invalid_size_rejected(self, tmp_path):
        with pytest.raises(ValueError):
            SessionPool(0, fake_config(tmp_path))