from pyvbaharness.vbasig import (
    KIND_FUNCTION,
    KIND_PROPERTY_GET,
    KIND_SUB,
    find_procedure,
    list_procedures,
    logical_lines,
    parse_declaration,
    split_top_level,
    strip_comment,
)


class TestLexing:
    def test_strip_comment_outside_strings(self):
        assert strip_comment('x = 1 \' note').rstrip() == "x = 1"
        assert strip_comment('x = "a \' b"') == 'x = "a \' b"'

    def test_line_continuations_joined(self):
        source = ("Public Function F(ByVal a As Long, _\n"
                  "        ByVal b As Long) As Long\n")
        joined = logical_lines(source)
        assert len(joined) == 1
        assert "ByVal b As Long) As Long" in joined[0]

    def test_split_top_level_respects_nesting(self):
        assert split_top_level("a, b") == ["a", "b"]
        assert split_top_level("a As Variant = Array(1, 2), b") == [
            "a As Variant = Array(1, 2)", "b"]
        assert split_top_level('a As String = "x,y", b') == [
            'a As String = "x,y"', "b"]
        assert split_top_level("") == []


class TestDeclarations:
    def test_simple_sub(self):
        signature = parse_declaration("Public Sub Main()")
        assert signature.kind == KIND_SUB
        assert signature.name == "Main"
        assert signature.required == 0

    def test_sub_without_parentheses(self):
        signature = parse_declaration("Sub Main")
        assert signature is not None
        assert signature.kind == KIND_SUB
        assert signature.required == 0

    def test_function_with_params(self):
        signature = parse_declaration(
            "Private Function Add(ByVal a As Long, b) As Long")
        assert signature.kind == KIND_FUNCTION
        assert signature.required == 2
        assert signature.optional == 0

    def test_optional_and_paramarray(self):
        signature = parse_declaration(
            "Public Function F(a, Optional b As Long = 3, "
            "ParamArray rest() As Variant) As String")
        assert signature.required == 1
        assert signature.optional == 1
        assert signature.has_param_array

    def test_property_get_is_callable(self):
        signature = parse_declaration("Public Property Get Value() As Long")
        assert signature.kind == KIND_PROPERTY_GET
        assert signature.returns_value

    def test_property_let_and_set_rejected(self):
        assert parse_declaration("Public Property Let Value(v As Long)") is None
        assert parse_declaration("Public Property Set Obj(v As Object)") is None

    def test_declare_statement_parsed(self):
        signature = parse_declaration(
            'Public Declare PtrSafe Function Sleep Lib "kernel32" '
            '(ByVal ms As Long)')
        assert signature is not None
        assert signature.name == "Sleep"

    def test_non_declarations(self):
        for line in ("Dim x As Long", "End Sub", "' Sub Commented()",
                     "x = 1"):
            assert parse_declaration(line) is None

    def test_static_and_friend_modifiers(self):
        assert parse_declaration("Friend Static Sub S()").name == "S"


class TestArity:
    def test_exact_arity(self):
        signature = parse_declaration("Sub S(a, b)")
        assert not signature.accepts(1)
        assert signature.accepts(2)
        assert not signature.accepts(3)
        assert signature.arity_text() == "2"

    def test_optional_range(self):
        signature = parse_declaration("Sub S(a, Optional b, Optional c)")
        assert signature.accepts(1)
        assert signature.accepts(3)
        assert not signature.accepts(4)
        assert signature.arity_text() == "1 to 3"

    def test_param_array_unbounded(self):
        signature = parse_declaration("Sub S(a, ParamArray rest())")
        assert signature.accepts(1)
        assert signature.accepts(50)
        assert not signature.accepts(0)
        assert signature.arity_text() == "1 or more"


class TestModuleScan:
    SOURCE = """
Option Explicit

' Sub NotReal()
Public Sub First()
End Sub

Private Function Second(ByVal a As Long) As Long
    Second = a
End Function
"""

    def test_find_by_name_case_insensitive(self):
        signature = find_procedure(self.SOURCE, "second")
        assert signature is not None
        assert signature.kind == KIND_FUNCTION
        assert signature.required == 1

    def test_commented_declaration_ignored(self):
        assert find_procedure(self.SOURCE, "NotReal") is None

    def test_missing_name(self):
        assert find_procedure(self.SOURCE, "Absent") is None

    def test_list_all(self):
        names = [s.name for s in list_procedures(self.SOURCE)]
        assert names == ["First", "Second"]
