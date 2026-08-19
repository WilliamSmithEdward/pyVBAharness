"""App capability table and the refusals it drives.

These run without Office: the point is that the capability facts measured
against live Office are asserted in one place, and that the code paths
gating on them behave.
"""
import pytest

from pyvbaharness import apps
from pyvbaharness.pool import SessionPool
from pyvbaharness.results import HarnessError
from pyvbaharness.session import (
    AccessSession,
    ExcelSession,
    PowerPointSession,
    WordSession,
    session_for,
)
from pyvbaharness.worker.hosts import HOSTS, HostError, build_host


class TestRegistry:
    def test_every_app_has_a_host_and_a_session(self):
        assert set(apps.APP_KEYS) == set(HOSTS)
        for key in apps.APP_KEYS:
            assert session_for.__globals__["SESSIONS"][key].app == key

    def test_hosts_take_their_identity_from_the_table(self):
        """One source of truth: a host must not carry its own ProgID."""
        for key, detail in apps.APPS.items():
            host = build_host(key)
            assert host.progid == detail.progid
            assert host.image_name == detail.image_name
            assert host.can_hide == detail.can_hide
            assert host.document_noun == detail.document_noun

    def test_unknown_app_is_rejected(self):
        with pytest.raises(ValueError, match="Unknown app"):
            apps.info("publisher")
        with pytest.raises(ValueError, match="Unknown app"):
            build_host("publisher")
        with pytest.raises(ValueError, match="Unknown app"):
            session_for("outlook")


class TestMeasuredCapabilities:
    """Facts established by running each application, not by assumption."""

    def test_only_powerpoint_is_single_instance(self):
        single = {k for k, v in apps.APPS.items() if not v.multi_instance}
        assert single == {"powerpoint"}

    def test_only_powerpoint_cannot_be_hidden(self):
        visible = {k for k, v in apps.APPS.items() if not v.can_hide}
        assert visible == {"powerpoint"}

    def test_only_access_has_no_vbom_gate(self):
        ungated = {k for k, v in apps.APPS.items() if not v.needs_vbom}
        assert ungated == {"access"}

    def test_only_excel_has_a_grid(self):
        grids = {k for k, v in apps.APPS.items() if v.has_grid}
        assert grids == {"excel"}

    def test_access_has_no_save_formats(self):
        assert apps.info("access").save_formats == ()


class TestSessionSurface:
    def test_grid_methods_exist_only_on_excel(self):
        for method in ("read_range", "write_range", "reset_sheets",
                       "run_batch"):
            assert hasattr(ExcelSession, method)
            for other in (WordSession, PowerPointSession, AccessSession):
                assert not hasattr(other, method), (
                    f"{other.__name__} should not expose {method}")

    def test_excel_keeps_its_workbook_aliases(self):
        for alias in ("new_workbook", "open_workbook", "has_workbook",
                      "excel_pid"):
            assert hasattr(ExcelSession, alias)

    def test_document_api_is_shared(self):
        for cls in (ExcelSession, WordSession, PowerPointSession,
                    AccessSession):
            for method in ("new_document", "open_document", "run_vba",
                           "run_macro", "add_module", "export_modules",
                           "compile_project"):
                assert hasattr(cls, method), f"{cls.__name__}.{method}"


class TestPoolRefusesSingleInstanceApps:
    def test_powerpoint_pool_larger_than_one_is_refused(self):
        """Refused up front, with the reason, rather than failing on the
        second member after paying for the first."""
        with pytest.raises(HarnessError, match="single-instance"):
            SessionPool(size=2, app="powerpoint")

    def test_message_names_the_process(self):
        reason = apps.single_instance_reason("powerpoint")
        assert "POWERPNT.EXE" in reason
        assert "PowerPoint" in reason


class TestTeardownToleratesADeadHost:
    """win32com's dynamic dispatch raises AttributeError, not com_error, when
    a member cannot be resolved on a disconnected object. A host whose
    process already exited must still close cleanly."""

    def _host(self, app):
        host = build_host(app)

        class Disconnected:
            def __getattr__(self, name):
                raise AttributeError(f"Add.{name}")

        host.app = Disconnected()
        host.document = Disconnected()
        return host

    @pytest.mark.parametrize("app", sorted(apps.APP_KEYS))
    def test_close_reports_instead_of_crashing(self, app):
        host = self._host(app)
        with pytest.raises(HostError, match="no longer reachable"):
            host.close_document()

    @pytest.mark.parametrize("app", sorted(apps.APP_KEYS))
    def test_quit_reports_instead_of_crashing(self, app):
        host = self._host(app)
        with pytest.raises(HostError, match="no longer reachable"):
            host.quit()

    def test_a_real_typo_still_raises_attributeerror(self):
        """The tolerance is scoped to teardown; ordinary code paths must not
        swallow a mistake in this file."""
        host = self._host("excel")
        with pytest.raises(AttributeError):
            host.read_range("Sheet1", "A1")
