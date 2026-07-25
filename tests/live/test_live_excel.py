"""Live integration tests: real Excel, real VBA, real hangs.

Run with:  python -m pytest tests/live -m live -o addopts="" -v

Requires desktop Excel and "Trust access to the VBA project object model".
One module-scoped session is shared for speed; each test that needs clean
workbook state calls new_workbook(). Hang and modal tests kill the session
deliberately; auto_recycle restores it for the next test.
"""
import json
import subprocess
import time
import warnings

import pytest

from pyvbaharness import (
    COMPILE_ACCEPTED,
    COMPILE_REJECTED,
    MODAL_BLOCKED,
    PASSED,
    RUNNER_ERROR,
    TIMEOUT,
    VBA_ERROR,
    ExcelSession,
    HarnessConfig,
    HarnessError,
)
from pyvbaharness.process_control import is_process_alive

pytestmark = pytest.mark.live


@pytest.fixture(scope="module")
def session():
    config = HarnessConfig(default_timeout_s=30.0, auto_recycle=True)
    with ExcelSession(config) as live:
        yield live


BASIC_SOURCE = """
Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function

Public Sub Main()
    PyVbaLog "sum=" & Trim$(Str$(AddNums(2, 3)))
End Sub
"""


class TestManagedRuns:
    def test_trivial_run_with_output(self, session):
        result = session.run_vba(BASIC_SOURCE, proc="Main")
        assert result.outcome == PASSED
        assert result.output == ["sum=5"]
        assert result.value is None  # Sub returns Empty -> null

    def test_function_return_value_and_args(self, session):
        session.run_vba(BASIC_SOURCE, proc="Main")
        result = session.run_macro("PyVbaUserCode.AddNums", 20, 22)
        assert result.outcome == PASSED
        assert result.value == 42

    def test_two_d_array_return(self, session):
        source = """
Public Function Grid() As Variant
    Dim v(1 To 2, 1 To 2) As Variant
    v(1, 1) = 1
    v(1, 2) = "x"
    v(2, 1) = True
    v(2, 2) = 2.5
    Grid = v
End Function
"""
        result = session.run_vba(source, proc="Grid")
        assert result.outcome == PASSED
        assert result.value == [[1, "x"], [True, 2.5]]

    def test_vba_runtime_error_captured(self, session):
        source = """
Public Sub Boom()
    Err.Raise 513, "TestSource", "custom failure"
End Sub
"""
        result = session.run_vba(source, proc="Boom")
        assert result.outcome == VBA_ERROR
        assert result.error is not None
        assert result.error.number == 513
        assert result.error.source == "TestSource"
        assert "custom failure" in result.error.description

    def test_unicode_roundtrip(self, session):
        source = """
Public Function Echo(ByVal text As String) As String
    Echo = text & " <- back"
End Function
"""
        result = session.run_vba(source, proc="Echo",
                                 args=("café über",))
        assert result.outcome == PASSED
        assert result.value == "café über <- back"


class TestHangResistance:
    def test_infinite_loop_times_out_and_recycles(self, session):
        source = """
Public Sub Freeze()
    Do
    Loop
End Sub
"""
        result = session.run_vba(source, proc="Freeze", timeout=8.0)
        assert result.outcome == TIMEOUT
        # auto_recycle: the next run must work in a fresh generation.
        after = session.run_vba(BASIC_SOURCE, proc="Main")
        assert after.outcome == PASSED

    def test_msgbox_ok_is_dismissed(self, session):
        source = """
Public Sub Hello()
    MsgBox "hello from the harness"
    PyVbaLog "survived"
End Sub
"""
        result = session.run_vba(source, proc="Hello", timeout=20.0)
        assert result.outcome == PASSED
        assert result.output == ["survived"]
        assert any(d.action.startswith("click:") for d in result.dialogs)

    def test_msgbox_yesno_blocks_and_kills(self, session):
        source = """
Public Sub Ask()
    Dim answer As Long
    answer = MsgBox("continue?", vbYesNo)
End Sub
"""
        started = time.monotonic()
        result = session.run_vba(source, proc="Ask", timeout=60.0)
        elapsed = time.monotonic() - started
        assert result.outcome == MODAL_BLOCKED
        assert elapsed < 30.0  # no waiting out the timeout
        assert result.dialogs


class TestErrorLines:
    def test_error_reports_original_source_line(self, session):
        source = """
Public Sub Main()
    Dim x As Long
    x = 1
    Err.Raise 513, "Here", "boom"
    x = 2
End Sub
"""
        result = session.run_vba(source, proc="Main")
        assert result.outcome == VBA_ERROR
        # The Err.Raise sits on line 5 of the injected source (leading
        # newline counts as line 1).
        assert result.error.line == 5
        assert "at line 5" in str(result.error)

    def test_error_in_loop_reports_the_looping_line(self, session):
        source = """
Public Sub Main()
    Dim i As Long
    Dim v As Variant
    For i = 1 To 5
        v = Array(1, 2, 3)(i)
    Next i
End Sub
"""
        result = session.run_vba(source, proc="Main")
        assert result.outcome == VBA_ERROR
        assert result.error.line == 6

    def test_numbering_can_be_disabled(self, session):
        source = """
Public Sub Main()
    Err.Raise 513, "Here", "boom"
End Sub
"""
        result = session.run_vba(source, proc="Main", line_numbers=False)
        assert result.outcome == VBA_ERROR
        assert result.error.line is None


class TestInjectionCache:
    def test_repeat_run_vba_skips_reinjection(self, session):
        source = """
Public Function Stamp() As Long
    Stamp = 7
End Function
"""
        session.new_workbook()
        first = session.run_vba(source, proc="Stamp")
        assert first.outcome == PASSED
        adds_before = sum(
            1 for e in session.events
            if e.get("kind") == "command-finished"
            and e.get("command") == "add_module")
        again = session.run_vba(source, proc="Stamp")
        assert again.outcome == PASSED
        adds_after = sum(
            1 for e in session.events
            if e.get("kind") == "command-finished"
            and e.get("command") == "add_module")
        assert adds_after == adds_before  # cache hit: no new add_module

    def test_changed_source_reinjects(self, session):
        session.new_workbook()
        v1 = session.run_vba("Public Function F() As Long\n    F = 1\n"
                             "End Function\n", proc="F")
        v2 = session.run_vba("Public Function F() As Long\n    F = 2\n"
                             "End Function\n", proc="F")
        assert (v1.value, v2.value) == (1, 2)


class TestEval:
    def test_expression_value(self, session):
        session.new_workbook()
        assert session.eval("1 + 2 * 3") == 7
        assert session.eval("WorksheetFunction.Sum(1, 2, 3)") == 6.0
        assert session.eval('UCase$("abc")') == "ABC"

    def test_expression_error_raises(self, session):
        session.new_workbook()
        with pytest.raises(HarnessError) as caught:
            session.eval("1 / 0")
        assert "11" in str(caught.value)  # VBA error 11: division by zero

    def test_multiline_rejected_locally(self, session):
        with pytest.raises(ValueError):
            session.eval("1 +\n2")


class TestVbaTestRunner:
    SUITE = """
Public Sub TestAddition()
    PyVbaAssertEqual 4, 2 + 2
End Sub

Public Sub TestFailing()
    PyVbaAssertEqual 5, 2 + 2, "arithmetic is broken"
End Sub

Public Sub TestErroring()
    Err.Raise 91
End Sub

Public Sub HelperNotATest()
End Sub

Public Sub TestNeedsArg(ByVal x As Long)
End Sub
"""

    def test_discovers_and_reports_each_case(self, session):
        session.new_workbook()
        results = session.run_tests(self.SUITE)
        by_name = {r.name: r for r in results}
        # Helper excluded by prefix; arg-taking proc excluded by arity.
        assert sorted(by_name) == ["TestAddition", "TestErroring",
                                   "TestFailing"]
        assert by_name["TestAddition"].passed
        failing = by_name["TestFailing"]
        assert not failing.passed
        assert failing.is_assertion_failure
        assert "arithmetic is broken" in failing.result.error.description
        assert "Expected 5 but got 4" in failing.result.error.description
        erroring = by_name["TestErroring"]
        assert not erroring.passed
        assert not erroring.is_assertion_failure
        assert erroring.result.error.number == 91

    def test_hanging_test_recovers_and_suite_continues(self, session):
        """A hang costs one test, not the suite: the session recycles,
        reinjects the module, and the remaining tests really run."""
        suite = """
Public Sub TestHangs()
    Do
    Loop
End Sub

Public Sub TestStillRuns()
    PyVbaLog "made it"
End Sub
"""
        session.new_workbook()
        results = session.run_tests(suite, timeout=6.0)
        assert [r.name for r in results] == ["TestHangs", "TestStillRuns"]
        assert results[0].result.outcome == TIMEOUT
        assert results[1].result.outcome == PASSED
        assert results[1].result.output == ["made it"]


class TestRangesAndFiles:
    def test_range_roundtrip(self, session):
        session.new_workbook()
        session.write_range("Sheet1", "A1", [[1, 2], [3, 4]])
        data = session.read_range("Sheet1", "A1:B2")
        assert data == [[1.0, 2.0], [3.0, 4.0]]

    def test_large_write_after_macro_run(self, session):
        """Regression guard for the post-macro large-write wedge.

        A single Value2 assignment over ~6000+ cells hangs Excel once a
        macro has run in the workbook, so write_range chunks. Without the
        chunking this test hangs until the watchdog kills the session.
        """
        session.new_workbook()
        session.add_module("Warm", """
Public Function Ping() As Long
    Ping = 1
End Function
""")
        assert session.run_macro("Warm.Ping").outcome == PASSED
        block = [[float(r * 100 + c) for c in range(100)] for r in range(100)]
        started = time.monotonic()
        result = session.write_range("Sheet1", "A1", block, timeout=25.0)
        assert time.monotonic() - started < 15.0
        assert result["rows"] == 100 and result["columns"] == 100
        assert result["chunks"] > 1
        back = session.read_range("Sheet1", "A1:CV100")
        assert len(back) == 100 and len(back[0]) == 100
        assert back[0][0] == 0.0
        assert back[99][99] == 9999.0

    def test_save_open_run_roundtrip(self, session, tmp_path):
        target = tmp_path / "saved_harness.xlsm"
        session.new_workbook()
        session.add_module("SavedMod", """
Public Function Stamp() As String
    Stamp = "from saved workbook"
End Function
""")
        session.write_range("Sheet1", "A1", [["persisted"]])
        session.save_as(target)
        session.open_workbook(target, read_only=True)
        result = session.run_macro("SavedMod.Stamp")
        assert result.outcome == PASSED
        assert result.value == "from saved workbook"
        assert session.read_range("Sheet1", "A1") == [["persisted"]]


class TestCompileCheck:
    def test_valid_project_accepted_fast(self, session):
        session.new_workbook()
        session.add_module("GoodMod", """
Option Explicit
Public Function Fine() As Long
    Fine = 1
End Function
""")
        result = session.compile_project(watch_seconds=8.0)
        assert result.outcome == COMPILE_ACCEPTED
        # The disabled Compile control is a positive completion signal, so
        # a clean compile must not wait out the whole watch window.
        assert result.duration_s < 6.0

    def test_harness_calls_compile_with_support_module(self, session):
        session.new_workbook()
        session.add_module("UsesLog", """
Public Sub Main()
    PyVbaLog "hello"
    PyVbaAssertEqual 1, 1
End Sub
""")
        bare = session.compile_project(watch_seconds=8.0)
        assert bare.outcome == COMPILE_REJECTED  # PyVbaLog unresolved
        session.new_workbook()
        session.add_module("UsesLog", """
Public Sub Main()
    PyVbaLog "hello"
    PyVbaAssertEqual 1, 1
End Sub
""")
        with_support = session.compile_project(watch_seconds=8.0,
                                               include_harness_support=True)
        assert with_support.outcome == COMPILE_ACCEPTED

    def test_syntax_error_rejected_with_dialog_text(self, session):
        session.new_workbook()
        session.add_module("BadMod", """
Option Explicit
Public Sub Broken()
    Dim x As Long
    x =
End Sub
""")
        result = session.compile_project(watch_seconds=20.0)
        assert result.outcome == COMPILE_REJECTED
        assert result.dialog is not None
        joined = " ".join([result.dialog.title, result.dialog.message,
                           *result.dialog.texts]).lower()
        assert "compile error" in joined
        # Leave a clean workbook for any later test.
        session.new_workbook()


class TestRawRuns:
    def test_run_raw_value(self, session):
        session.new_workbook()
        session.add_module("RawMod", """
Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
""")
        assert session.run_raw("RawMod.Twice", 21) == 42

    def test_run_raw_missing_proc_raises(self, session):
        session.new_workbook()
        with pytest.raises(HarnessError):
            session.run_raw("NoSuchModule.NoSuchProc")


class TestOwnership:
    def test_sessions_never_share_an_excel_instance(self, session):
        """Regression guard for the attach bug.

        win32com's Dispatch("Excel.Application") calls GetActiveObject first
        and returns the ALREADY RUNNING Excel. Since the harness kills its
        instance on a hang, attaching would put a user's own workbooks in the
        blast radius. A second session must get its own process.
        """
        assert session.excel_pid > 0
        other = ExcelSession(HarnessConfig(exclusive=False,
                                           auto_recycle=False))
        try:
            assert other.excel_pid > 0
            assert other.excel_pid != session.excel_pid
        finally:
            other.close()
        # Killing the second session must not have disturbed the first.
        assert session.run_vba(BASIC_SOURCE, proc="Main").outcome == PASSED


class TestTeardownHygiene:
    def test_excel_process_dies_with_session(self):
        config = HarnessConfig(auto_recycle=False)
        with ExcelSession(config) as short:
            pid = short.excel_pid
            assert pid > 0
            assert is_process_alive(pid)
        # close() must guarantee this, not merely start it: Quit() returning
        # successfully does not prove the process ended.
        assert not is_process_alive(pid)
        assert short.oracle_issues == []

    def test_excel_dies_even_after_being_made_visible(self):
        """Excel made visible treats itself as user-launched and survives
        the disconnect of its last automation client. close() must still
        leave nothing behind."""
        config = HarnessConfig(exclusive=False, auto_recycle=False)
        session = ExcelSession(config)
        pid = session.excel_pid
        try:
            session.new_workbook()
            session.add_module("Vis", """
Public Function Fine() As Long
    Fine = 1
End Function
""")
            session.compile_project(watch_seconds=5.0)
        finally:
            session.close()
        assert not is_process_alive(pid)

    def test_job_object_kills_excel_when_worker_dies(self):
        """Killing the worker outright (no cleanup path runs) must still
        take Excel down: the kernel job's kill-on-close does it without any
        cooperation. This is the crash-safety guarantee."""
        config = HarnessConfig(exclusive=False, auto_recycle=False)
        session = ExcelSession(config)
        excel_pid = session.excel_pid
        worker_pid = session._proc.pid
        created = [e for e in session.events
                   if e.get("kind") == "excel-created"][0]
        try:
            assert created.get("job_kill_on_close") is True
            assert is_process_alive(excel_pid)
            subprocess.run(["taskkill.exe", "/PID", str(worker_pid), "/F"],
                           capture_output=True, check=False)
            deadline = time.monotonic() + 10.0
            while time.monotonic() < deadline and is_process_alive(excel_pid):
                time.sleep(0.2)
            assert not is_process_alive(excel_pid)
        finally:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                session.close()

    def test_save_trace_writes_jsonl(self, tmp_path):
        config = HarnessConfig(exclusive=False, auto_recycle=False)
        with ExcelSession(config) as session:
            session.new_workbook()
            trace_path = session.save_trace(tmp_path / "trace.jsonl")
        lines = trace_path.read_text(encoding="utf-8").splitlines()
        assert lines
        parsed = [json.loads(line) for line in lines]
        assert any(e.get("kind") == "excel-created" for e in parsed)

    def test_no_manifest_files_survive(self, tmp_path):
        config = HarnessConfig(exclusive=False, manifest_dir=tmp_path,
                               auto_recycle=True)
        session = ExcelSession(config)
        try:
            session.new_workbook()
            # Force a recycle so the manifest is reused, not orphaned.
            session.run_vba("Public Sub Spin()\n    Do\n    Loop\nEnd Sub",
                            proc="Spin", timeout=6.0)
            session.run_vba(BASIC_SOURCE, proc="Main")
        finally:
            session.close()
        assert list(tmp_path.glob("*.json")) == []
