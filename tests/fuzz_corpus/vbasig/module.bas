Attribute VB_Name = "Module1"
Option Explicit

Public Sub TestAddition()
    Debug.Print 1 + 1 ' a comment
End Sub

Private Function Add(ByVal a As Long, Optional b As Long = 1) As Long
    Add = a + b
End Function

Public Function Total(ParamArray values() As Variant) As Double
End Function
