"""Plugin collection and reporting, driven through pytester with the
scripted fake worker (no Excel).

The fake worker answers any run whose target contains "Hang" by never
replying and everything else by passing, so collection, selection, and the
infrastructure-vs-VBA failure distinction are all testable offline.
"""
import sys
from pathlib import Path

import pytest

pytest_plugins = ["pytester"]

FAKE_WORKER = str(Path(__file__).with_name("fake_worker.py"))
SRC_ROOT = str(Path(__file__).resolve().parents[2] / "src")

INI = f"""
[pytest]
vba_collect = true
vba_test_timeout = 5
vba_worker_argv = {sys.executable} {FAKE_WORKER}
"""

SUITE = """
Public Sub TestAlpha()
End Sub

Public Sub TestBeta()
End Sub

Public Sub HelperNotCollected()
End Sub

Public Sub TestNeedsArgs(ByVal x As Long)
End Sub
"""


def write_project(pytester, suite: str = SUITE,
                  filename: str = "test_suite.bas") -> None:
    pytester.makefile(".ini", pytest=INI.strip())
    (pytester.path / filename).write_text(suite, encoding="utf-8")
    pytester.makeconftest(
        f"import sys\nsys.path.insert(0, {SRC_ROOT!r})\n")


class TestCollection:
    def test_collects_zero_arg_test_procedures(self, pytester):
        write_project(pytester)
        result = pytester.runpytest("--collect-only", "-q")
        stdout = result.stdout.str()
        assert "TestAlpha" in stdout
        assert "TestBeta" in stdout
        assert "HelperNotCollected" not in stdout
        assert "TestNeedsArgs" not in stdout

    def test_non_test_bas_files_ignored(self, pytester):
        write_project(pytester, filename="helpers.bas")
        result = pytester.runpytest("--collect-only", "-q")
        assert "TestAlpha" not in result.stdout.str()

    def test_collection_can_be_disabled(self, pytester):
        write_project(pytester)
        pytester.makefile(".ini", pytest="[pytest]\nvba_collect = false\n")
        result = pytester.runpytest("--collect-only", "-q")
        assert "TestAlpha" not in result.stdout.str()

    def test_keyword_selection_works(self, pytester):
        write_project(pytester)
        result = pytester.runpytest("-q", "-k", "Alpha")
        result.assert_outcomes(passed=1)


class TestExecution:
    def test_passing_suite(self, pytester):
        write_project(pytester)
        result = pytester.runpytest("-q")
        result.assert_outcomes(passed=2)

    def test_marker_applied(self, pytester):
        write_project(pytester)
        result = pytester.runpytest("-q", "-m", "vba")
        result.assert_outcomes(passed=2)

    def test_hanging_test_reports_infrastructure_failure(self, pytester):
        write_project(pytester,
                      suite=("Public Sub TestHangs()\nEnd Sub\n"
                             "Public Sub TestAfter()\nEnd Sub\n"))
        result = pytester.runpytest("-q")
        # The hang fails as infrastructure; the session recycles and the
        # following test still runs and passes.
        result.assert_outcomes(passed=1, failed=1)
        assert "infrastructure failure" in result.stdout.str()
