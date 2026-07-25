import pytest

from pyvbaharness import codegen
from pyvbaharness.vbasig import (
    KIND_FUNCTION,
    KIND_SUB,
    ProcedureSignature,
)


class TestIdentifiers:
    def test_accepts_plain_identifiers(self):
        assert codegen.is_vba_identifier("Main")
        assert codegen.is_vba_identifier("a_1")
        assert codegen.is_vba_identifier("A" * 31)

    def test_rejects_bad_identifiers(self):
        for bad in ("", "1abc", "has space", "a-b", "a.b", "A" * 32, "a'b"):
            assert not codegen.is_vba_identifier(bad)


class TestRunTarget:
    def test_accepts_proc_and_module_proc(self):
        codegen.validate_run_target("Main")
        codegen.validate_run_target("Module1.DoThing")

    def test_rejects_cross_workbook_and_injection_forms(self):
        for bad in ("'Book1.xlsm'!Mod.Proc", "Book1.xlsm!Mod.Proc",
                    "a.b.c", "Mod.Proc()", "[x].y", "", "Mod. Proc"):
            with pytest.raises(ValueError):
                codegen.validate_run_target(bad)

    def test_rejects_harness_internal_modules(self):
        for reserved in codegen.HARNESS_MODULE_NAMES:
            with pytest.raises(ValueError):
                codegen.validate_run_target(f"{reserved}.PyVbaRun")
            with pytest.raises(ValueError):
                codegen.validate_module_name(reserved.lower())


class TestHeaderStrip:
    def test_strips_class_header(self):
        source = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1  'True\r\n"
            "END\r\n"
            'Attribute VB_Name = "Thing"\r\n'
            "Attribute VB_Exposed = False\r\n"
            "Option Explicit\r\n"
            "Public Sub A()\r\nEnd Sub\r\n"
        )
        body = codegen.strip_module_header(source)
        assert body.startswith("Option Explicit")
        assert "Attribute" not in body
        assert "VERSION" not in body

    def test_strips_bas_attribute_line(self):
        source = 'Attribute VB_Name = "Mod"\r\nSub A()\r\nEnd Sub\r\n'
        assert codegen.strip_module_header(source) == "Sub A()\r\nEnd Sub\r\n"

    def test_headerless_source_unchanged(self):
        source = "Option Explicit\r\nSub A()\r\nEnd Sub\r\n"
        assert codegen.strip_module_header(source) == source


class TestQualifiedRef:
    def test_escapes_quotes_in_workbook_name(self):
        ref = codegen.qualified_run_ref("Bob's Book.xlsm", "M", "P")
        assert ref == "'Bob''s Book.xlsm'!M.P"


class TestSupportModule:
    def test_generated_source_shape(self):
        source = codegen.support_module_source()
        assert "Public Sub PyVbaLog" in source
        assert "Public Function PyVbaOutputJson" in source
        assert "Public Function PyVbaJsonValue" in source
        # AddFromString safety: no Attribute header lines.
        assert "Attribute VB_" not in source
        # JSON literal survived Python quoting: VBA doubled quotes intact.
        assert '"{""outcome"":""vba-error"",""number"":"' in source
        # Locale-invariant numeric formatting only.
        assert "Str$(" in source
        assert "CStr(number)" not in source
        # Backslash escaping emits real single-backslash VBA literals.
        assert 'Replace$(text, "\\", "\\\\")' in source
        assert '"\\n"' in source

    def test_balanced_structure(self):
        source = codegen.support_module_source()
        # Seven functions; five subs (reset, log, fail, and the asserts).
        assert source.count("End Function") == 7
        assert source.count("End Sub") == 5
        assert source.count("Select Case") == source.count("End Select")
        assert source.count("If ") >= source.count("End If")

    def test_assert_helpers_present(self):
        source = codegen.support_module_source()
        assert "Public Sub PyVbaAssert(" in source
        assert "Public Sub PyVbaAssertEqual(" in source
        assert 'Err.Raise 517, "PyVbaHarness.Assert"' in source

    def test_erl_capture_plumbing(self):
        source = codegen.support_module_source()
        assert "Public Sub PyVbaFail(" in source
        assert "Public Function PyVbaLastErl(" in source
        assert "ByVal errorLine As Long" in source
        assert '",""line"":"' in source
        dispatcher = codegen.call_module_source(
            "M", "P", ProcedureSignature("P", "sub", 0, 0, False), 0)
        assert "errLine = PyVbaLastErl()" in dispatcher


class TestCallModule:
    def test_function_call_assigns_result(self):
        signature = ProcedureSignature("Add", KIND_FUNCTION, 2, 0, False)
        lines = codegen.call_expression("Mod", "Add", signature, 2)
        assert lines == [
            "    resultValue = Mod.Add((callArgs(0)), (callArgs(1)))"]

    def test_sub_call_uses_call_and_empty(self):
        signature = ProcedureSignature("Go", KIND_SUB, 1, 0, False)
        lines = codegen.call_expression("Mod", "Go", signature, 1)
        assert lines == ["    Call Mod.Go((callArgs(0)))",
                         "    resultValue = Empty"]

    def test_zero_args_omits_parentheses(self):
        function = ProcedureSignature("F", KIND_FUNCTION, 0, 0, False)
        sub = ProcedureSignature("S", KIND_SUB, 0, 0, False)
        assert codegen.call_expression("M", "F", function, 0) == [
            "    resultValue = M.F"]
        assert codegen.call_expression("M", "S", sub, 0)[0] == "    Call M.S"

    def test_module_source_wraps_call_in_error_handler(self):
        signature = ProcedureSignature("Boom", KIND_SUB, 0, 0, False)
        source = codegen.call_module_source("User", "Boom", signature, 0)
        assert "Public Function PyVbaRun" in source
        assert "On Error GoTo Caught" in source
        assert "Call User.Boom" in source
        # The dispatcher must not route through Application.Run: that would
        # break in-VBA error trapping.
        assert "Application.Run" not in source
        assert source.count("End Function") == 1
