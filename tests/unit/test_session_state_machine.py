"""Supervisor state-machine tests against the scripted fake worker.

These validate the hang-resistance contract without Excel: watchdog breach
kills and marks dead, modal-blocked aborts immediately, worker death maps to
runner-error, auto-recycle restores service, and a clean session leaves an
oracle-clean trace.
"""
import sys
import warnings
from pathlib import Path

import pytest

from pyvbaharness import (
    MODAL_BLOCKED,
    PASSED,
    RUNNER_ERROR,
    TIMEOUT,
    ExcelSession,
    HarnessConfig,
    SessionDead,
)
from pyvbaharness.oracle import validate_trace

FAKE_WORKER = str(Path(__file__).with_name("fake_worker.py"))


def fake_config(tmp_path, **overrides) -> HarnessConfig:
    defaults = dict(
        worker_argv=[sys.executable, FAKE_WORKER],
        exclusive=False,
        manifest_dir=tmp_path,
        startup_timeout_s=15.0,
        default_timeout_s=5.0,
        cleanup_grace_s=3.0,
        auto_recycle=False,
    )
    defaults.update(overrides)
    return HarnessConfig(**defaults)


class TestGracefulAbort:
    """A host that is still answering is asked to quit before it is killed.

    Terminating a healthy Office writes the open document into its
    crash-recovery list, and the user meets that as a recovery pane the next
    time they start the application.
    """

    def _last_kill(self, session) -> dict:
        for event in reversed(session.events):
            if event.get("kind") == "app-killed":
                return event
        return {}

    def test_recycle_quits_instead_of_killing(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        try:
            session.run_macro("Mod.Ok")
            session.recycle()
            event = self._last_kill(session)
            assert event.get("reason") == "recycled"
            assert event.get("quit_cleanly") is True
            assert event.get("killed") is False
        finally:
            session.close()

    def test_a_wedged_host_is_killed_without_waiting(self, tmp_path):
        """A timeout means the host is not answering, so asking it to quit
        would only add latency to a decision already made."""
        session = ExcelSession(fake_config(tmp_path))
        try:
            result = session.run_macro("Hang.Forever")
            assert result.outcome == TIMEOUT
            event = self._last_kill(session)
            assert event.get("quit_cleanly") is False
            assert event.get("killed") is True
        finally:
            session.close()

    def test_a_dead_worker_falls_through_to_the_kill(self, tmp_path):
        """Nothing is left to ask, and the host it owned may have outlived
        it, so the kill path still has to run."""
        session = ExcelSession(fake_config(tmp_path))
        try:
            result = session.run_macro("Die.Now")
            assert result.outcome == RUNNER_ERROR
            assert self._last_kill(session).get("quit_cleanly") is False
        finally:
            session.close()


class TestHappyPath:
    def test_run_and_clean_close(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        try:
            session.new_workbook()
            result = session.run_macro("Ok.Now")
            assert result.outcome == PASSED
            assert result.value == 42
            assert result.output == ["hi"]
        finally:
            session.close()
        assert session.oracle_issues == []
        assert validate_trace(session.events) == []

    def test_manifest_removed_after_close(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        session.close()
        assert list(tmp_path.glob("*.json")) == []


class TestTimeout:
    def test_hang_is_killed_and_reported(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        try:
            result = session.run_macro("Hang.Forever", timeout=2.0)
            assert result.outcome == TIMEOUT
            with pytest.raises(SessionDead):
                session.run_macro("Ok.Now")
        finally:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                session.close()
        kinds = [e.get("kind") for e in session.events]
        assert "app-killed" in kinds
        # Oracle stays clean: the hang was followed by a kill.
        assert validate_trace(session.events) == []

    def test_auto_recycle_restores_service(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path, auto_recycle=True))
        try:
            assert session.run_macro("Hang.Forever",
                                     timeout=2.0).outcome == TIMEOUT
            result = session.run_macro("Ok.Now")
            assert result.outcome == PASSED
        finally:
            session.close()


class TestModalBlocked:
    def test_modal_aborts_immediately(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        try:
            result = session.run_macro("Modal.Blocked", timeout=30.0)
            assert result.outcome == MODAL_BLOCKED
            assert result.duration_s < 10.0  # no waiting out the timeout
            assert result.dialogs and result.dialogs[0].title
        finally:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                session.close()
        assert validate_trace(session.events) == []


class TestWorkerDeath:
    def test_death_maps_to_runner_error(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        try:
            result = session.run_macro("Die.Now")
            assert result.outcome == RUNNER_ERROR
        finally:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                session.close()


class TestValidationBeforeCom:
    def test_bad_target_rejected_locally(self, tmp_path):
        session = ExcelSession(fake_config(tmp_path))
        try:
            with pytest.raises(ValueError):
                session.run_macro("'Other.xlsm'!Mod.Proc")
            with pytest.raises(ValueError):
                session.add_module("bad name", "Sub A()\nEnd Sub")
            with pytest.raises(ValueError):
                session.add_module("Mod", "Sub A()\nEnd Sub", kind="form")
        finally:
            session.close()
