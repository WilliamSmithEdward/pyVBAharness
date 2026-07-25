"""VBA source generation and validation for the injected harness modules.

Two modules are injected:

- ``PyVbaHarnessRunner`` (static): output collection and JSON serialization.
  Added once per workbook.
- ``PyVbaHarnessCall`` (generated per target): ``PyVbaRun`` wraps one direct
  call to the target procedure in ``On Error`` and returns a JSON string.

The direct call is the accuracy core. An error raised inside a procedure
invoked through ``Application.Run`` does not unwind into the calling VBA
procedure's error handler (the call crosses a COM boundary), so Excel shows
its runtime-error dialog instead of the harness capturing the error. A direct
call unwinds normally and is trapped. Verified live on 2026-07-25: the same
``Err.Raise`` produced a dialog through ``Application.Run`` and
``caught:513|TestSource|custom failure`` through a direct call.

Because the call is direct, its shape must match the callee (Sub versus
Function, exact argument count), which is why vbasig parses the declaration
first and the dispatcher is regenerated when the target or arity changes.

Other contracts:

- Sources injected via ``CodeModule.AddFromString`` must not contain
  ``Attribute`` header lines; ``strip_module_header`` removes them.
- ``Str$`` is used for numeric text because it is locale-invariant; ``CStr``
  would emit the user's decimal separator and break the JSON.
- Arguments are passed wrapped in parentheses to force ByVal evaluation, so
  a Variant from the ParamArray can satisfy a typed ByRef parameter.
"""
from __future__ import annotations

import re

from .vbasig import ProcedureSignature

SUPPORT_MODULE_NAME = "PyVbaHarnessRunner"
CALL_MODULE_NAME = "PyVbaHarnessCall"
RUNNER_ENTRY = "PyVbaRun"
MAX_RUN_ARGS = 10

HARNESS_MODULE_NAMES = (SUPPORT_MODULE_NAME, CALL_MODULE_NAME)

# Plain VBA identifier: letter first, then letters/digits/underscore, max 31.
_VBA_IDENTIFIER_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")

_ATTRIBUTE_LINE_RE = re.compile(r"^Attribute\s+VB_", re.IGNORECASE)
_VERSION_LINE_RE = re.compile(r"^VERSION\s+\d", re.IGNORECASE)


def is_vba_identifier(name: str) -> bool:
    return bool(_VBA_IDENTIFIER_RE.match(name))


def validate_run_target(target: str) -> None:
    """Accept ``Proc`` or ``Module.Proc``; reject anything else.

    Rejecting workbook-qualified and bracketed forms keeps every run inside
    the harness-owned workbook (authority boundary, checked before COM).
    """
    parts = target.split(".")
    if len(parts) not in (1, 2) or not all(is_vba_identifier(p) for p in parts):
        raise ValueError(
            f"Run target {target!r} is not a plain 'Proc' or 'Module.Proc' "
            "VBA name inside the harness workbook."
        )
    if parts[0].lower() in {n.lower() for n in HARNESS_MODULE_NAMES}:
        raise ValueError(
            f"Run target {target!r} names a harness-internal module.")


def validate_module_name(name: str) -> None:
    if not is_vba_identifier(name):
        raise ValueError(f"Module name {name!r} is not a valid VBA identifier.")
    if name.lower() in {n.lower() for n in HARNESS_MODULE_NAMES}:
        raise ValueError(
            f"Module name {name!r} is reserved by the harness.")


def strip_module_header(source: str) -> str:
    """Remove the VERSION/BEGIN/END block and Attribute VB_* header lines.

    Exported .bas/.cls files carry these headers; ``AddFromString`` treats
    them as code and the project stops compiling.
    """
    lines = source.splitlines(keepends=True)
    i = 0
    if lines and _VERSION_LINE_RE.match(lines[i]):
        i += 1
        if i < len(lines) and lines[i].strip().upper() == "BEGIN":
            i += 1
            while i < len(lines) and lines[i].strip().upper() != "END":
                i += 1
            if i < len(lines):
                i += 1
    while i < len(lines) and _ATTRIBUTE_LINE_RE.match(lines[i]):
        i += 1
    return "".join(lines[i:]).lstrip("\r\n")


def qualified_run_ref(workbook_name: str, module: str, proc: str) -> str:
    """Workbook-qualified macro reference with quote escaping."""
    escaped = workbook_name.replace("'", "''")
    return f"'{escaped}'!{module}.{proc}"


def call_expression(module: str, proc: str, signature: ProcedureSignature,
                    arg_count: int) -> list[str]:
    """VBA statement lines that invoke the target and set ``resultValue``."""
    args = ", ".join(f"(callArgs({i}))" for i in range(arg_count))
    call_args = f"({args})" if args else ""
    qualified = f"{module}.{proc}"
    if signature.returns_value:
        return [f"    resultValue = {qualified}{call_args}"]
    return [f"    Call {qualified}{call_args}", "    resultValue = Empty"]


def call_module_source(module: str, proc: str, signature: ProcedureSignature,
                       arg_count: int) -> str:
    """Generated per-target dispatcher module."""
    body = "\n".join(call_expression(module, proc, signature, arg_count))
    return f'''Option Explicit

Public Function {RUNNER_ENTRY}(ParamArray callArgs() As Variant) As String
    Dim resultValue As Variant
    PyVbaResetOutput
    On Error GoTo Caught
{body}
    On Error GoTo 0
    {RUNNER_ENTRY} = "{{""outcome"":""passed"",""value"":" & _
        PyVbaJsonValue(resultValue) & ",""output"":" & PyVbaOutputJson() & "}}"
    Exit Function
Caught:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errLine As Long
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errLine = PyVbaLastErl()
    On Error GoTo 0
    {RUNNER_ENTRY} = PyVbaFailureJson(errNumber, errSource, _
        errDescription, errLine)
End Function
'''


def support_module_source() -> str:
    """Static support module: output collection and JSON serialization.

    ``PyVbaLog`` is public so user VBA can call it unqualified; the collected
    lines travel back inside the JSON result.
    """
    return '''Option Explicit

Private mOutput As Collection
Private mErlNoted As Boolean
Private mErl As Long

Public Sub PyVbaResetOutput()
    Set mOutput = New Collection
    mErlNoted = False
    mErl = 0
End Sub

Public Sub PyVbaFail(ByVal atLine As Long, ByVal number As Long, _
        ByVal source As String, ByVal description As String)
    If Not mErlNoted Then
        mErlNoted = True
        mErl = atLine
    End If
    If number = 0 Then Exit Sub
    Err.Raise number, source, description
End Sub

Public Function PyVbaLastErl() As Long
    PyVbaLastErl = mErl
End Function

Public Sub PyVbaLog(ByVal message As Variant)
    If mOutput Is Nothing Then Set mOutput = New Collection
    On Error Resume Next
    mOutput.Add CStr(message)
End Sub

Public Function PyVbaOutputJson() As String
    Dim parts As String
    Dim item As Variant
    If mOutput Is Nothing Then
        PyVbaOutputJson = "[]"
        Exit Function
    End If
    For Each item In mOutput
        If Len(parts) > 0 Then parts = parts & ","
        parts = parts & """" & PyVbaJsonEscape(CStr(item)) & """"
    Next item
    PyVbaOutputJson = "[" & parts & "]"
End Function

Public Function PyVbaFailureJson(ByVal number As Long, _
        ByVal source As String, ByVal description As String, _
        ByVal errorLine As Long) As String
    PyVbaFailureJson = "{""outcome"":""vba-error"",""number"":" & _
        Trim$(Str$(number)) & ",""source"":""" & PyVbaJsonEscape(source) & _
        """,""description"":""" & PyVbaJsonEscape(description) & _
        """,""line"":" & Trim$(Str$(errorLine)) & _
        ",""output"":" & PyVbaOutputJson() & "}"
End Function

Public Sub PyVbaAssert(ByVal condition As Boolean, _
        Optional ByVal message As String = "")
    If condition Then Exit Sub
    If Len(message) = 0 Then message = "Assertion failed"
    Err.Raise 517, "PyVbaHarness.Assert", message
End Sub

Public Sub PyVbaAssertEqual(ByVal expected As Variant, _
        ByVal actual As Variant, Optional ByVal message As String = "")
    Dim same As Boolean
    On Error Resume Next
    same = (expected = actual)
    If Err.Number <> 0 Then same = False
    On Error GoTo 0
    If same Then Exit Sub
    Dim detail As String
    detail = "Expected " & PyVbaJsonValue(expected) & _
        " but got " & PyVbaJsonValue(actual)
    If Len(message) > 0 Then detail = message & ": " & detail
    Err.Raise 517, "PyVbaHarness.Assert", detail
End Sub

Public Function PyVbaJsonValue(ByVal value As Variant) As String
    On Error GoTo Unserializable
    Select Case VarType(value)
        Case vbEmpty, vbNull
            PyVbaJsonValue = "null"
        Case vbBoolean
            If value Then
                PyVbaJsonValue = "true"
            Else
                PyVbaJsonValue = "false"
            End If
        Case vbByte, vbInteger, vbLong, vbLongLong, vbSingle, vbDouble, _
                vbCurrency, vbDecimal
            PyVbaJsonValue = Trim$(Str$(value))
        Case vbDate
            PyVbaJsonValue = """" & Format$(value, "yyyy-mm-dd") & "T" & _
                Format$(value, "hh:nn:ss") & """"
        Case vbString
            PyVbaJsonValue = """" & PyVbaJsonEscape(CStr(value)) & """"
        Case Else
            If IsArray(value) Then
                PyVbaJsonValue = PyVbaArrayJson(value)
            ElseIf IsObject(value) Then
                PyVbaJsonValue = """<object:" & _
                    PyVbaJsonEscape(TypeName(value)) & ">"""
            Else
                PyVbaJsonValue = """<" & _
                    PyVbaJsonEscape(TypeName(value)) & ">"""
            End If
    End Select
    Exit Function
Unserializable:
    PyVbaJsonValue = """<unserializable>"""
End Function

Private Function PyVbaArrayJson(ByVal value As Variant) As String
    Dim dimensions As Long
    Dim parts As String
    Dim rowParts As String
    Dim r As Long
    Dim c As Long
    dimensions = PyVbaArrayDimensions(value)
    If dimensions = 1 Then
        For r = LBound(value, 1) To UBound(value, 1)
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & PyVbaJsonValue(value(r))
        Next r
        PyVbaArrayJson = "[" & parts & "]"
    ElseIf dimensions = 2 Then
        For r = LBound(value, 1) To UBound(value, 1)
            rowParts = ""
            For c = LBound(value, 2) To UBound(value, 2)
                If Len(rowParts) > 0 Then rowParts = rowParts & ","
                rowParts = rowParts & PyVbaJsonValue(value(r, c))
            Next c
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & "[" & rowParts & "]"
        Next r
        PyVbaArrayJson = "[" & parts & "]"
    Else
        PyVbaArrayJson = """<array:" & Trim$(Str$(dimensions)) & "d>"""
    End If
End Function

Private Function PyVbaArrayDimensions(ByVal value As Variant) As Long
    Dim probe As Long
    Dim count As Long
    On Error GoTo Done
    For count = 1 To 60
        probe = LBound(value, count)
    Next count
Done:
    PyVbaArrayDimensions = count - 1
End Function

Public Function PyVbaJsonEscape(ByVal text As String) As String
    Dim result As String
    Dim i As Long
    Dim code As Long
    result = Replace$(text, "\\", "\\\\")
    result = Replace$(result, """", "\\""")
    result = Replace$(result, vbCrLf, "\\n")
    result = Replace$(result, vbCr, "\\n")
    result = Replace$(result, vbLf, "\\n")
    result = Replace$(result, vbTab, "\\t")
    For i = 1 To Len(result)
        code = AscW(Mid$(result, i, 1))
        If code >= 0 And code < 32 Then
            result = Left$(result, i - 1) & " " & Mid$(result, i + 1)
        End If
    Next i
    PyVbaJsonEscape = result
End Function
'''
