"""Live coverage of the v0.4.0 feature set: stack traces, batching,
liveness timeouts, coverage, reset, screenshots, export/import, fuzzing.

Run with:  python -m pytest tests/live -m live -o addopts="" -v
"""
import time
from pathlib import Path

import pytest

from pyvbaharness import (
    PASSED,
    TIMEOUT,
    VBA_ERROR,
    ExcelSession,
    HarnessConfig,
)

pytestmark = pytest.mark.live


@pytest.fixture(scope="module")
def session():
    with ExcelSession(HarnessConfig(default_timeout_s=30.0,
                                    auto_recycle=True)) as live:
        yield live


class TestStackTraces:
    def test_nested_error_reports_every_frame(self, session):
        source = """
Public Sub Outer()
    Middle
End Sub

Public Sub Middle()
    Inner
End Sub

Public Sub Inner()
    Err.Raise 513, "Deep", "from the bottom"
End Sub
"""
        result = session.run_vba(source, proc="Outer")
        assert result.outcome == VBA_ERROR
        frames = result.error.stack
        assert [proc for proc, _line in frames] == [
            "PyVbaUserCode.Inner",
            "PyVbaUserCode.Middle",
            "PyVbaUserCode.Outer",
        ]
        # Deepest frame first: Err.Raise is line 11 of the source above
        # (the triple-quoted literal starts with a newline), and each outer
        # frame reports its own call site.
        assert frames[0][1] == 11
        assert frames[1][1] == 7
        assert frames[2][1] == 3
        assert result.error.line == 11
        assert "stack:" in str(result.error)

    def test_single_frame_error(self, session):
        result = session.run_vba("""
Public Sub Main()
    Err.Raise 5
End Sub
""", proc="Main")
        assert [p for p, _ in result.error.stack] == ["PyVbaUserCode.Main"]


class TestBatchExecution:
    SOURCE = """
Public Function AddTwo(ByVal a As Long, ByVal b As Long) As Long
    AddTwo = a + b
End Function

Public Function Shout(ByVal text As String) As String
    Shout = UCase$(text)
End Function

Public Sub Boom()
    Err.Raise 513, "Batch", "planned failure"
End Sub
"""

    def test_mixed_batch_results_in_order(self, session):
        session.new_workbook()
        session.add_module("BatchMod", self.SOURCE, line_numbers=True)
        results = session.run_batch([
            ("BatchMod.AddTwo", (2, 3)),
            ("BatchMod.Shout", ("hi",)),
            ("BatchMod.Boom", ()),
            ("BatchMod.AddTwo", (10, 20)),
        ])
        assert [r.outcome for r in results] == [
            PASSED, PASSED, VBA_ERROR, PASSED]
        assert results[0].value == 5
        assert results[1].value == "HI"
        assert results[2].error.number == 513
        # Error lines and stacks survive batching (Err.Raise is line 11).
        assert results[2].error.line == 11
        assert results[2].error.stack == [("BatchMod.Boom", 11)]
        assert results[3].value == 30

    def test_argument_types_round_trip_exactly(self, session):
        session.new_workbook()
        session.add_module("EchoMod", """
Public Function Echo(ByVal v As Variant) As Variant
    Echo = v
End Function
""")
        values = ["5", "3/4", "=SUM(A1)", "plain", 7, -2, 2.5, True, False,
                  None, ""]
        results = session.run_batch(
            [("EchoMod.Echo", (v,)) for v in values])
        assert all(r.outcome == PASSED for r in results)
        got = [r.value for r in results]
        # Text that Excel would coerce (numbers, dates, formulas) survives
        # as text; numerics stay numeric; None becomes Empty -> null.
        assert got[0] == "5"
        assert got[1] == "3/4"
        assert got[2] == "=SUM(A1)"
        assert got[3] == "plain"
        assert got[4] == 7
        assert got[5] == -2
        assert got[6] == 2.5
        assert got[7] is True and got[8] is False
        assert got[9] is None
        assert got[10] == ""

    def test_batch_is_much_faster_than_serial(self, session):
        session.new_workbook()
        session.add_module("PerfMod", """
Public Function Tiny(ByVal n As Long) As Long
    Tiny = n + 1
End Function
""")
        count = 500
        # Warm both paths: the first batch generates its dispatcher module,
        # which would otherwise be measured as batch cost.
        session.run_macro("PerfMod.Tiny", 0)
        session.run_batch([("PerfMod.Tiny", (0,))])

        started = time.monotonic()
        for n in range(count):
            session.run_macro("PerfMod.Tiny", n)
        serial_s = time.monotonic() - started

        started = time.monotonic()
        batched = session.run_batch(
            [("PerfMod.Tiny", (n,)) for n in range(count)], timeout=120)
        batch_s = time.monotonic() - started

        assert [r.value for r in batched] == list(range(1, count + 1))
        # Measured 4.7x at 200 calls and 9.9x at 3000 on Excel 365 x64; the
        # ratio grows with batch size as the single round trip amortizes.
        # Assert a conservative floor so the test is not machine-fragile.
        assert batch_s * 3 < serial_s, (
            f"batch {batch_s:.3f}s vs serial {serial_s:.3f}s")

    def test_unknown_target_rejected_before_com(self, session):
        with pytest.raises(ValueError):
            session.run_batch([("'Other.xlsm'!M.P", ())])
        assert session.run_batch([]) == []


class TestLivenessTimeout:
    def test_progress_keeps_a_long_run_alive(self, session):
        source = """
Public Sub Steady()
    Dim i As Long
    Dim started As Single
    For i = 1 To 6
        started = Timer
        Do While Timer - started < 0.7!
        Loop
        PyVbaProgress i / 6, "step " & i
    Next i
End Sub
"""
        seen = []
        result = session.run_vba(source, proc="Steady", idle_timeout=3.0,
                                 on_progress=lambda f, m: seen.append((f, m)))
        # Total runtime (~4.2s) exceeds the 3s idle window; only the
        # heartbeats keep it alive.
        assert result.outcome == PASSED
        assert result.duration_s > 3.5
        assert len(seen) >= 5
        assert seen[-1][1] == "step 6"

    def test_silence_kills_within_the_idle_window(self, session):
        source = """
Public Sub GoesQuiet()
    PyVbaProgress 0.1, "starting"
    Do
    Loop
End Sub
"""
        started = time.monotonic()
        result = session.run_vba(source, proc="GoesQuiet", idle_timeout=4.0)
        elapsed = time.monotonic() - started
        assert result.outcome == TIMEOUT
        assert "idle timeout" in result.message
        assert elapsed < 20.0


class TestCoverage:
    def test_line_coverage_marks_taken_branch(self, session):
        source = """
Public Sub Branch(ByVal flag As Boolean)
    Dim x As Long
    If flag Then
        x = 1
    Else
        x = 2
    End If
    x = x + 1
End Sub
"""
        session.new_workbook()
        session.add_module("CovMod", source, coverage=True)
        assert session.run_macro("CovMod.Branch", True).outcome == PASSED
        report = session.coverage_report()
        module = report.modules["covmod"]
        assert 5 in module.hit      # the taken branch
        assert 7 in module.missed   # the untaken branch
        assert 9 in module.hit      # after the If
        assert 0 < module.percent < 100

        assert session.run_macro("CovMod.Branch", False).outcome == PASSED
        after = session.coverage_report().modules["covmod"]
        assert 7 in after.hit       # hits accumulate across runs
        assert after.percent == 100.0

    def test_coverage_does_not_break_error_lines(self, session):
        source = """
Public Sub Boom()
    Dim x As Long
    x = 1
    Err.Raise 513, "Cov", "still mapped"
End Sub
"""
        session.new_workbook()
        session.add_module("CovErr", source, coverage=True)
        result = session.run_macro("CovErr.Boom")
        assert result.outcome == VBA_ERROR
        assert result.error.line == 5


class TestSheetReset:
    def test_reset_clears_cells_but_keeps_modules(self, session):
        session.new_workbook()
        session.add_module("KeptMod", """
Public Function Kept() As Long
    Kept = 99
End Function
""")
        session.write_range("Sheet1", "A1", [[1, 2], [3, 4]])
        started = time.monotonic()
        session.reset_sheets()
        reset_s = time.monotonic() - started
        assert session.read_range("Sheet1", "A1:B2") == [[None, None],
                                                         [None, None]]
        assert session.run_macro("KeptMod.Kept").value == 99
        assert reset_s < 1.0


class TestExportImport:
    def test_round_trip_through_files(self, session, tmp_path):
        session.new_workbook()
        session.add_module("RoundTrip", """
Public Function Answer() As Long
    Answer = 42
End Function
""")
        files = session.export_modules(tmp_path)
        # Document modules (ThisWorkbook, Sheet1) export too; find ours.
        matches = [f for f in files if Path(f).name == "RoundTrip.bas"]
        assert matches
        exported = Path(matches[0]).read_text(encoding="utf-8-sig",
                                              errors="replace")
        assert "Answer = 42" in exported

        session.new_workbook()
        imported = session.import_modules(tmp_path)
        assert "RoundTrip" in imported
        assert session.run_macro("RoundTrip.Answer").value == 42


class TestFailureArtifacts:
    def test_timeout_captures_a_screenshot(self, tmp_path):
        config = HarnessConfig(exclusive=False, auto_recycle=False,
                               artifact_dir=tmp_path)
        with ExcelSession(config) as short:
            result = short.run_vba("""
Public Sub Spin()
    Do
    Loop
End Sub
""", proc="Spin", timeout=5.0)
        assert result.outcome == TIMEOUT
        if result.screenshot:  # capture is best-effort by contract
            image = Path(result.screenshot)
            assert image.exists() and image.stat().st_size > 1000
            assert image.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


class TestExcelProvenance:
    def test_version_recorded_in_trace(self, session):
        created = [e for e in session.events
                   if e.get("kind") == "excel-created"][-1]
        assert created.get("excel_version")
        assert created.get("excel_build")


class TestPropertyBasedFuzzing:
    def test_finds_a_real_counterexample(self, session):
        from pyvbaharness.properties import check_vba_function

        # Overflows for large inputs: Integer maxes at 32767.
        source = """
Public Function Doubler(ByVal n As Integer) As Integer
    Doubler = n * 2
End Function
"""
        session.new_workbook()
        with pytest.raises(AssertionError) as caught:
            check_vba_function(session, source, "Doubler",
                               max_examples=200)
        assert "Overflow" in str(caught.value) or "6" in str(caught.value)

    def test_holds_for_a_correct_function(self, session):
        from pyvbaharness.properties import check_vba_function

        source = """
Public Function SafeAbs(ByVal n As Long) As Double
    SafeAbs = Abs(CDbl(n))
End Function
"""
        session.new_workbook()
        check_vba_function(
            session, source, "SafeAbs",
            check=lambda args, value: value == abs(args[0]),
            max_examples=40)
