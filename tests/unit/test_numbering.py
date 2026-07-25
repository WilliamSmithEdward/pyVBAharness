from pyvbaharness.numbering import add_line_numbers, instrument_error_lines


def lines(source: str) -> list[str]:
    return add_line_numbers(source).splitlines()


class TestBasicNumbering:
    def test_numbers_executable_body_lines_with_physical_index(self):
        source = ("Public Sub Main()\n"      # 1
                  "    x = 1\n"              # 2
                  "    y = 2\n"              # 3
                  "End Sub\n")               # 4
        out = lines(source)
        assert out[0] == "Public Sub Main()"
        assert out[1] == "2 x = 1"
        assert out[2] == "3 y = 2"
        assert out[3] == "End Sub"

    def test_module_level_lines_untouched(self):
        source = ("Option Explicit\n"
                  "Private mState As Long\n"
                  "Public Sub A()\n"
                  "    mState = 1\n"
                  "End Sub\n")
        out = lines(source)
        assert out[0] == "Option Explicit"
        assert out[1] == "Private mState As Long"
        assert out[3] == "4 mState = 1"

    def test_multiple_procedures(self):
        source = ("Sub A()\n"
                  "    x = 1\n"
                  "End Sub\n"
                  "Function B() As Long\n"
                  "    B = 2\n"
                  "End Function\n")
        out = lines(source)
        assert out[1] == "2 x = 1"
        assert out[4] == "5 B = 2"

    def test_property_get_body_numbered(self):
        source = ("Public Property Get Value() As Long\n"
                  "    Value = 9\n"
                  "End Property\n")
        assert lines(source)[1] == "2 Value = 9"


class TestSkippedLines:
    def test_declarations_comments_blanks(self):
        source = ("Sub A()\n"
                  "    Dim x As Long\n"
                  "    Const c = 1\n"
                  "    Static s As Long\n"
                  "\n"
                  "    ' comment\n"
                  "    Rem old comment\n"
                  "    x = c\n"
                  "End Sub\n")
        out = lines(source)
        assert out[1].lstrip().startswith("Dim")
        assert out[2].lstrip().startswith("Const")
        assert out[3].lstrip().startswith("Static")
        assert out[7] == "8 x = c"

    def test_structure_keywords_not_numbered(self):
        source = ("Sub A()\n"
                  "    Select Case x\n"
                  "        Case 1\n"
                  "            y = 1\n"
                  "        Case Else\n"
                  "            y = 2\n"
                  "    End Select\n"
                  "    If y = 1 Then\n"
                  "        z = 1\n"
                  "    ElseIf y = 2 Then\n"
                  "        z = 2\n"
                  "    Else\n"
                  "        z = 3\n"
                  "    End If\n"
                  "    For i = 1 To 3\n"
                  "        t = i\n"
                  "    Next i\n"
                  "    Do\n"
                  "        u = 1\n"
                  "    Loop While False\n"
                  "End Sub\n")
    # Case/Else/ElseIf/End/Next/Loop must stay unnumbered; the statements
    # inside them must be numbered.
        out = lines(source)
        assert out[1] == "2 Select Case x"
        assert out[2].lstrip() == "Case 1"
        assert out[3] == "4 y = 1"
        assert out[4].lstrip() == "Case Else"
        assert out[6].lstrip() == "End Select"
        assert out[7] == "8 If y = 1 Then"
        assert out[9].lstrip().startswith("ElseIf")
        assert out[11].lstrip() == "Else"
        assert out[13].lstrip() == "End If"
        assert out[14] == "15 For i = 1 To 3"
        assert out[16].lstrip().startswith("Next")
        assert out[17] == "18 Do"
        assert out[19].lstrip().startswith("Loop")

    def test_continuation_lines_not_numbered(self):
        source = ("Sub A()\n"
                  "    x = 1 + _\n"
                  "        2\n"
                  "    y = 3\n"
                  "End Sub\n")
        out = lines(source)
        assert out[1] == "2 x = 1 + _"
        assert out[2].lstrip() == "2"
        assert out[3] == "4 y = 3"

    def test_labels_and_preprocessor_not_numbered(self):
        source = ("Sub A()\n"
                  "    On Error GoTo Handler\n"
                  "    #If Win64 Then\n"
                  "    x = 1\n"
                  "    #End If\n"
                  "    Exit Sub\n"
                  "Handler:\n"
                  "    x = 2\n"
                  "End Sub\n")
        out = lines(source)
        assert out[1] == "2 On Error GoTo Handler"
        assert out[2].lstrip().startswith("#If")
        assert out[3] == "4 x = 1"
        assert out[5] == "6 Exit Sub"
        assert out[6] == "Handler:"
        assert out[7] == "8 x = 2"

    def test_single_line_procedure_untouched(self):
        source = "Sub A(): x = 1: End Sub\n"
        assert add_line_numbers(source).splitlines()[0] == source.rstrip()


class TestSafety:
    def test_existing_numbers_disable_numbering(self):
        source = ("Sub A()\n"
                  "10 x = 1\n"
                  "    y = 2\n"
                  "End Sub\n")
        assert add_line_numbers(source) == source

    def test_idempotent_via_existing_number_guard(self):
        source = ("Sub A()\n"
                  "    x = 1\n"
                  "End Sub\n")
        once = add_line_numbers(source)
        assert add_line_numbers(once) == once


class TestInstrumentation:
    def test_handler_injected_per_procedure(self):
        source = ("Public Sub Main()\n"
                  "    x = 1\n"
                  "End Sub\n"
                  "Public Function F() As Long\n"
                  "    F = 2\n"
                  "End Function\n")
        out = instrument_error_lines(source, module_name="Mod").splitlines()
        assert out[0] == "Public Sub Main()"
        assert out[1] == "    On Error GoTo PyVbaErl__"
        assert out[2] == "2 x = 1"
        assert out[3] == "    Exit Sub"
        assert out[4] == "PyVbaErl__:"
        assert out[5].lstrip() == ('PyVbaFail "Mod.Main", Erl, Err.Number, '
                                   "Err.Source, Err.Description")
        assert out[6] == "End Sub"
        # Second procedure gets its own handler with the right Exit kind
        # and its own frame name.
        assert "    Exit Function" in out
        assert out.count("PyVbaErl__:") == 2
        assert any('PyVbaFail "Mod.F"' in line for line in out)

    def test_frame_name_without_module(self):
        source = "Sub Solo()\n    x = 1\nEnd Sub\n"
        out = instrument_error_lines(source)
        assert 'PyVbaFail "Solo", Erl' in out

    def test_property_gets_exit_property(self):
        source = ("Public Property Get V() As Long\n"
                  "    V = 1\n"
                  "End Property\n")
        out = instrument_error_lines(source).splitlines()
        assert "    Exit Property" in out
        assert "PyVbaErl__:" in out

    def test_handler_lines_carry_no_numbers(self):
        source = ("Sub A()\n"
                  "    x = 1\n"
                  "End Sub\n")
        for line in instrument_error_lines(source).splitlines():
            if "PyVbaFail" in line or "On Error GoTo PyVbaErl__" in line:
                assert not line.strip()[0].isdigit()

    def test_single_line_procedure_untouched(self):
        source = "Sub A(): x = 1: End Sub\n"
        out = instrument_error_lines(source)
        assert "PyVbaErl__" not in out

    def test_existing_label_disables_instrumentation(self):
        source = ("Sub A()\n"
                  "    GoTo PyVbaErl__\n"
                  "PyVbaErl__:\n"
                  "End Sub\n")
        assert instrument_error_lines(source) == source

    def test_existing_numbers_disable_instrumentation(self):
        source = ("Sub A()\n"
                  "10 x = 1\n"
                  "End Sub\n")
        assert instrument_error_lines(source) == source

    def test_continued_declaration_handler_after_signature(self):
        source = ("Public Sub Main(ByVal a As Long, _\n"
                  "        ByVal b As Long)\n"
                  "    x = a + b\n"
                  "End Sub\n")
        out = instrument_error_lines(source).splitlines()
        assert out[0].endswith("_")
        assert out[1].lstrip().startswith("ByVal b")
        assert out[2] == "    On Error GoTo PyVbaErl__"
        assert out[3] == "3 x = a + b"

    def test_user_on_error_left_in_place(self):
        source = ("Sub A()\n"
                  "    On Error Resume Next\n"
                  "    x = 1 / 0\n"
                  "End Sub\n")
        out = instrument_error_lines(source)
        assert "On Error Resume Next" in out
        assert "On Error GoTo PyVbaErl__" in out


class TestCoverageInstrumentation:
    def test_hit_rides_the_numbered_line(self):
        from pyvbaharness.numbering import instrument_module

        source = ("Sub A()\n"          # 1
                  "    Dim x As Long\n"  # 2 (declaration: no hit)
                  "    x = 1\n"          # 3
                  "    x = 2\n"          # 4
                  "End Sub\n")
        result = instrument_module(source, module_name="M", coverage_id=7)
        out = result.source.splitlines()
        assert result.coverable_lines == [3, 4]
        assert "3 PyVbaCovHit 7, 3: x = 1" in out
        assert "4 PyVbaCovHit 7, 4: x = 2" in out
        # Line count is preserved apart from the three handler lines and
        # the On Error line, so Erl still maps to original lines.
        assert len(out) == len(source.splitlines()) + 4

    def test_block_openers_are_numbered_but_never_prefixed(self):
        from pyvbaharness.numbering import instrument_module

        source = ("Sub A()\n"
                  "    If flag Then\n"
                  "        x = 1\n"
                  "    End If\n"
                  "    For i = 1 To 3\n"
                  "        y = i\n"
                  "    Next i\n"
                  "    Do\n"
                  "        z = 1\n"
                  "    Loop While False\n"
                  "    With obj\n"
                  "        .A = 1\n"
                  "    End With\n"
                  "    Select Case x\n"
                  "    End Select\n"
                  "End Sub\n")
        result = instrument_module(source, module_name="M", coverage_id=1)
        out = result.source
        # A hit prefixed onto a block opener would turn it into a
        # single-line form and break compilation.
        for opener in ("If flag Then", "For i = 1 To 3", "Do",
                       "With obj", "Select Case x"):
            assert f"PyVbaCovHit 1, {opener}" not in out
        assert "2 If flag Then" in out.splitlines()
        # Bodies are still instrumented, and openers are not counted as
        # coverable because their execution cannot be observed. Line 12
        # (".A = 1" inside With) starts with a dot, which the conservative
        # numberer leaves alone.
        assert result.coverable_lines == [3, 6, 9]

    def test_without_coverage_no_hits(self):
        from pyvbaharness.numbering import instrument_module

        result = instrument_module("Sub A()\n    x = 1\nEnd Sub\n",
                                   module_name="M")
        assert "PyVbaCovHit" not in result.source
        assert result.coverable_lines == [2]

    def test_unsafe_source_reports_no_coverable_lines(self):
        from pyvbaharness.numbering import instrument_module

        source = "Sub A()\n10 x = 1\nEnd Sub\n"
        result = instrument_module(source, module_name="M", coverage_id=1)
        assert result.source == source
        assert result.coverable_lines == []
