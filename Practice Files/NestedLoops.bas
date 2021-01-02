Attribute VB_Name = "Module22"
Option Explicit
Sub loopy()
Dim i As Integer, j As Integer
For i = 1 To 2
    For j = 7 To 5 Step -2
        MsgBox "i = " & i & ", j = " & j
    Next j
    MsgBox "i = " & i & ", j = " & j
Next i
MsgBox "i = " & i & ", j = " & j
End Sub
