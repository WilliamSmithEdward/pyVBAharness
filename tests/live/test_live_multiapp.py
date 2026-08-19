"""Live integration tests for Word, PowerPoint and Access.

Run with:  python -m pytest tests/live -m live -o addopts="" -v

Requires the desktop applications, and for Word and PowerPoint the option
"Trust access to the VBA project object model". Access has no such option.

The shared contract (inject, run, trap errors, report source lines, survive
a hang, tear down without orphans) is parametrized across all three hosts.
Tests below that are specific to one host cover what the probes measured as
different about it: PowerPoint cannot be hidden, Access has no unsaved
database, rejects module-qualified Run references, and needs its injected
modules removed before close or it raises a Save As prompt.
"""
import time

import pytest

from pyvbaharness import (
    MODAL_BLOCKED,
    PASSED,
    TIMEOUT,
    AccessSession,
    HarnessConfig,
    HarnessError,
    PowerPointSession,
    VBA_ERROR,
    WordSession,
)
from pyvbaharness import apps
from pyvbaharness.process_control import is_process_alive

pytestmark = pytest.mark.live

SESSION_CLASSES = {
    "word": WordSession,
    "powerpoint": PowerPointSession,
    "access": AccessSession,
}

# PowerPoint returns the already-running process from a second COM
# activation, so only these can have two sessions alive at once. Anything
# that constructs an extra session is parametrized over this list; the
# PowerPoint equivalents live in TestPowerPointSingleInstance and reuse the
# one module-scoped session.
MULTI_INSTANCE_APPS = sorted(
    k for k in SESSION_CLASSES if apps.info(k).multi_instance)

BASIC_SOURCE = """
Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function

Public Sub Main()
    PyVbaLog "sum=" & Trim$(Str$(AddNums(2, 3)))
End Sub
"""

FAILING_SOURCE = """
Public Sub Boom()
    Dim denominator As Long
    denominator = 0
    PyVbaLog "before"
    Debug.Print 1 / denominator
    PyVbaLog "after"
End Sub
"""


def _waited_for_exit(pid: int, timeout_s: float = 20.0) -> bool:
    """True once the process is gone. Quit() returning is not proof of exit,
    so ownership assertions poll for the real thing."""
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if not is_process_alive(pid):
            return True
        time.sleep(0.2)
    return not is_process_alive(pid)


def _make_database(path) -> None:
    """Create an empty .accdb outside the harness.

    A test that opens a database the harness did not create needs one to
    exist first, and a live scratch database cannot be copied: Access holds
    it exclusively while it is open.
    """
    import pythoncom
    import pywintypes
    import win32com.client.dynamic

    pythoncom.CoInitialize()
    app = win32com.client.dynamic.Dispatch(pythoncom.CoCreateInstance(
        pywintypes.IID("Access.Application"), None,
        pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch))
    try:
        app.Visible = False
        app.NewCurrentDatabase(str(path))
        app.CloseCurrentDatabase()
    finally:
        app.Quit(2)
    # Quit returns before the process does, and the lock file outlives it.
    deadline = time.monotonic() + 20.0
    lock = path.with_suffix(".laccdb")
    while time.monotonic() < deadline and lock.exists():
        time.sleep(0.2)


@pytest.fixture(scope="module", params=sorted(SESSION_CLASSES))
def app_session(request):
    """One module-scoped session per application."""
    config = HarnessConfig(default_timeout_s=60.0, auto_recycle=True)
    with SESSION_CLASSES[request.param](config) as live:
        yield live


class TestSharedContract:
    """Behaviour every host must provide identically."""

    def test_run_with_output(self, app_session):
        result = app_session.run_vba(BASIC_SOURCE, proc="Main")
        assert result.outcome == PASSED, result.message
        assert result.output == ["sum=5"]

    def test_function_return_value_and_args(self, app_session):
        app_session.run_vba(BASIC_SOURCE, proc="Main")
        result = app_session.run_macro("PyVbaUserCode.AddNums", 20, 22)
        assert result.outcome == PASSED, result.message
        assert result.value == 42

    def test_vba_error_is_trapped_not_dialogged(self, app_session):
        """The whole point of the generated dispatcher: an error unwinds into
        a VBA handler instead of raising the host's modal error dialog."""
        result = app_session.run_vba(FAILING_SOURCE, proc="Boom")
        assert result.outcome == VBA_ERROR, result.message
        assert result.error is not None
        assert result.error.number == 11  # division by zero
        assert result.output == ["before"]

    def test_error_reports_the_failing_source_line(self, app_session):
        result = app_session.run_vba(FAILING_SOURCE, proc="Boom")
        assert result.outcome == VBA_ERROR
        assert result.error is not None
        line = FAILING_SOURCE.splitlines()[result.error.line - 1]
        assert "1 / denominator" in line

    def test_unicode_round_trips(self, app_session):
        source = (
            "Public Function Echo(ByVal s As Variant) As Variant\n"
            "    Echo = s & ChrW(233) & ChrW(26085)\n"
            "End Function\n")
        app_session.run_vba(
            source + "\nPublic Sub Main()\nEnd Sub\n", proc="Main")
        result = app_session.run_macro("PyVbaUserCode.Echo", "café ")
        assert result.outcome == PASSED, result.message
        assert result.value == "café é日"

    def test_list_procedures(self, app_session):
        app_session.run_vba(BASIC_SOURCE, proc="Main")
        names = {p["name"] for p in
                 app_session.list_procedures("PyVbaUserCode")}
        assert {"AddNums", "Main"} <= names

    def test_export_modules_writes_files(self, app_session, tmp_path):
        app_session.run_vba(BASIC_SOURCE, proc="Main")
        files = app_session.export_modules(tmp_path)
        exported = {f.rsplit("\\", 1)[-1] for f in files}
        assert "PyVbaUserCode.bas" in exported
        body = (tmp_path / "PyVbaUserCode.bas").read_text(
            encoding="utf-8-sig", errors="replace")
        assert "AddNums" in body

    def test_hang_times_out_and_recycles(self, app_session):
        """A wedged run must cost that run, not the session."""
        source = ("Public Sub Spin()\n"
                  "    Do While True\n"
                  "    Loop\n"
                  "End Sub\n")
        pid = app_session.app_pid
        result = app_session.run_vba(source, proc="Spin", timeout=6.0,
                                     module_name="PyVbaSpin")
        assert result.outcome == TIMEOUT
        assert app_session.is_dead
        # The owned host is killed, by PID, and nothing else is touched.
        assert _waited_for_exit(pid, 15.0), (
            f"the hung host {pid} was not killed")
        # auto_recycle brings the session back for the next test.
        assert app_session.run_vba(BASIC_SOURCE,
                                   proc="Main").outcome == PASSED


class TestOwnership:
    """Each session owns a private instance and takes it down with it."""

    @pytest.mark.parametrize("app", MULTI_INSTANCE_APPS)
    def test_instance_dies_with_the_session(self, app):
        session = SESSION_CLASSES[app]()
        try:
            session.run_vba(BASIC_SOURCE, proc="Main")
            pid = session.app_pid
            assert pid > 0
            assert is_process_alive(pid)
        finally:
            session.close()
        assert _waited_for_exit(pid), (
            f"{app} process {pid} outlived its session")

    @pytest.mark.parametrize("app", MULTI_INSTANCE_APPS)
    def test_two_sessions_never_share_an_instance(self, app):
        first = SESSION_CLASSES[app]()
        second = SESSION_CLASSES[app]()
        try:
            assert first.app_pid > 0 and second.app_pid > 0
            assert first.app_pid != second.app_pid
        finally:
            first.close()
            second.close()

    def test_trace_records_the_app(self, app_session):
        app_session.run_vba(BASIC_SOURCE, proc="Main")
        created = [e for e in app_session.events
                   if e.get("kind") == "app-created"]
        assert created and created[-1]["app"] == app_session.app

    @pytest.mark.parametrize("app", MULTI_INSTANCE_APPS)
    def test_trace_oracle_is_clean(self, app):
        session = SESSION_CLASSES[app]()
        session.run_vba(BASIC_SOURCE, proc="Main")
        session.close()
        assert session.oracle_issues == [], [
            i.message for i in session.oracle_issues]


class TestPowerPointSingleInstance:
    """PowerPoint's two measured limits, both reported rather than hidden."""

    @pytest.fixture
    def powerpoint(self, app_session):
        if app_session.app != "powerpoint":
            pytest.skip("PowerPoint-specific")
        return app_session

    def test_cannot_be_hidden_and_says_so(self, powerpoint):
        """PowerPoint refuses Application.Visible = False. The harness must
        report that honestly rather than claim a hidden host it lacks."""
        powerpoint.run_vba(BASIC_SOURCE, proc="Main")
        created = [e for e in powerpoint.events
                   if e.get("kind") == "app-created"][-1]
        assert created["can_hide"] is False
        assert created["visible"] is True

    def test_second_session_is_refused_not_shared(self, powerpoint):
        """A second activation returns the running process. Taking ownership
        of it would put someone else's presentation inside the blast radius
        of a timeout kill, so the harness refuses instead."""
        assert powerpoint.app_pid > 0
        with pytest.raises(HarnessError, match="single-instance"):
            PowerPointSession(HarnessConfig(exclusive=False)).close()
        # The refusal must not have disturbed the session that owns it.
        assert powerpoint.run_vba(BASIC_SOURCE,
                                  proc="Main").outcome == PASSED


class TestWordSpecifics:
    def test_save_and_reopen_runs_injected_code(self, tmp_path):
        target = tmp_path / "harness.docm"
        with WordSession() as session:
            session.run_vba(BASIC_SOURCE, proc="Main")
            session.save_as(target)
        assert target.exists()
        with WordSession() as session:
            session.open_document(target, read_only=False)
            result = session.run_macro("PyVbaUserCode.AddNums", 1, 2)
            assert result.outcome == PASSED, result.message
            assert result.value == 3

    def test_save_as_rejects_a_macro_free_format(self, tmp_path):
        with WordSession() as session:
            session.run_vba(BASIC_SOURCE, proc="Main")
            with pytest.raises(HarnessError, match="docm"):
                session.save_as(tmp_path / "dropped.docx")


class TestAccessSpecifics:
    def test_new_document_creates_a_scratch_database(self):
        with AccessSession() as session:
            info = session.new_document()
            assert info["path"].endswith(".accdb")
            assert session.run_vba(BASIC_SOURCE,
                                   proc="Main").outcome == PASSED

    def test_open_refuses_read_only(self, tmp_path):
        """Injecting into Access writes the file immediately, so a
        read_only=True request must be refused, not silently honoured."""
        with AccessSession() as session:
            session.new_document()
            with pytest.raises(HarnessError, match="read_only=False"):
                session.open_document(tmp_path / "nothing.accdb",
                                      read_only=True)

    def test_save_as_is_refused_with_an_explanation(self):
        with AccessSession() as session:
            session.new_document()
            with pytest.raises(HarnessError, match="continuously"):
                session.save_as("ignored.accdb")

    def test_close_does_not_wedge_on_unsaved_modules(self):
        """The 2026-08-18 wedge: an unsaved VBE module makes Access raise a
        modal 'Save As / Module Name' prompt at close, which blocked Quit.
        Teardown removes injected modules first, so it cannot appear."""
        session = AccessSession()
        session.run_vba(BASIC_SOURCE, proc="Main")
        pid = session.app_pid
        session.close()
        assert _waited_for_exit(pid), "Access outlived its session"
        assert session.oracle_issues == []


class TestTeardownModals:
    """Closing is where Office prompts, so it is where the harness must be
    watching. The watcher used to be stopped before teardown, which hid a
    close-time dialog completely and turned it into a 15 second wait for the
    cleanup deadline (measured 2026-08-19)."""

    STRAY = """
Public Sub MakeStray()
    Dim c As Object
    Set c = Application.VBE.ActiveVBProject.VBComponents.Add(1)
    c.Name = "StrayModule"
    c.CodeModule.AddFromString "Public Sub Nothing1()" & vbCrLf & "End Sub"
End Sub
"""

    def test_untracked_access_module_does_not_prompt(self):
        """A module the harness did not inject still has to be discarded in
        a scratch database, or Access raises Save As and blocks Quit. The
        harness created this database, so it owns every component in it."""
        session = AccessSession(HarnessConfig(cleanup_grace_s=5.0))
        assert session.run_vba(self.STRAY,
                               proc="MakeStray").outcome == PASSED
        pid = session.app_pid
        started = time.monotonic()
        session.close()
        elapsed = time.monotonic() - started
        assert _waited_for_exit(pid), "Access outlived its session"
        assert elapsed < 10.0, (
            f"teardown took {elapsed:.1f}s; a save prompt most likely "
            "blocked it")
        kinds = [e.get("kind") for e in session.events]
        assert "app-quit" in kinds, "expected a clean quit, not a kill"
        assert session.oracle_issues == []

    def test_close_time_prompt_is_reported_not_waited_out(self, tmp_path):
        """In a database the caller opened, the harness must not delete
        their components, so the prompt can happen. It must then be seen and
        acted on rather than sat through."""
        target = tmp_path / "caller.accdb"
        _make_database(target)
        assert target.exists()

        session = AccessSession(HarnessConfig(cleanup_grace_s=5.0))
        pid = session.app_pid
        session.open_document(target, read_only=False)
        assert session.run_vba(self.STRAY,
                               proc="MakeStray").outcome == PASSED
        started = time.monotonic()
        session.close()
        elapsed = time.monotonic() - started
        assert _waited_for_exit(pid), "Access outlived its session"
        assert elapsed < 10.0, (
            f"teardown took {elapsed:.1f}s; the prompt was waited out "
            "instead of detected")
        blocked = [e for e in session.events
                   if e.get("kind") == "modal-blocked"]
        assert blocked, "the close-time dialog was never reported"
        assert any("save" in str(e.get("title", "")).lower()
                   for e in blocked), [e.get("title") for e in blocked]


class TestExcelOnlyCommandsAreRefused:
    """Worksheet commands must fail with an explanation, not an AttributeError."""

    def test_read_range_is_not_offered(self, app_session):
        assert not hasattr(app_session, "read_range")

    def test_batch_is_not_offered(self, app_session):
        assert not hasattr(app_session, "run_batch")


class TestModalWedges:
    """Each host's modal surfaces, deliberately triggered.

    A MsgBox is the canonical VBA wedge: it blocks Application.Run inside a
    COM call that cannot be interrupted from within the worker. The watcher
    dismisses what it can read and reports what it cannot, and either way
    the session must come back.
    """

    def test_msgbox_ok_is_dismissed(self, app_session):
        source = """
Public Sub Hello()
    MsgBox "hello from the harness"
    PyVbaLog "survived"
End Sub
"""
        result = app_session.run_vba(source, proc="Hello", timeout=30.0,
                                     module_name="PyVbaModal")
        assert result.outcome == PASSED, result.message
        assert result.output == ["survived"]
        assert any(d.action.startswith("click:") for d in result.dialogs)

    def test_msgbox_needing_a_decision_blocks_and_kills(self, app_session):
        """A dialog the harness must not answer for the user is reported,
        not guessed at, and the host is killed rather than left waiting."""
        source = """
Public Sub Ask()
    Dim answer As Long
    answer = MsgBox("continue?", vbYesNo)
End Sub
"""
        pid = app_session.app_pid
        started = time.monotonic()
        result = app_session.run_vba(source, proc="Ask", timeout=90.0,
                                     module_name="PyVbaModal")
        elapsed = time.monotonic() - started
        assert result.outcome == MODAL_BLOCKED, result.outcome
        assert elapsed < 45.0, "reported by detection, not by waiting out the timeout"
        assert result.dialogs
        assert _waited_for_exit(pid), "the blocked host was not killed"
        assert app_session.run_vba(BASIC_SOURCE,
                                   proc="Main").outcome == PASSED


class TestTeardownPrompts:
    """Closing must never raise a prompt, on any host.

    These are the prompts that are prevented rather than dismissed: Office's
    own save dialogs carry no Win32 buttons, so a watcher can see them but
    never click them. The test is that teardown is clean with unsaved
    changes present.
    """

    @pytest.mark.parametrize("app", MULTI_INSTANCE_APPS)
    def test_unsaved_changes_do_not_block_teardown(self, app):
        session = SESSION_CLASSES[app]()
        # Dirty the document as well as the VBA project, so both of the
        # "do you want to save?" paths are live at close.
        session.run_vba(BASIC_SOURCE, proc="Main")
        pid = session.app_pid
        session.close()
        assert _waited_for_exit(pid), (
            f"{app} did not exit; a save prompt most likely blocked Quit")
        assert session.oracle_issues == []
