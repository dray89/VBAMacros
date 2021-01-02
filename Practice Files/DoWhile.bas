Attribute VB_Name = "Module20"
Option Explicit

Sub loopy()
Dim x As Double
x = 10
Do While (x > 9)
    If (x < 9.6) Then
        Exit Do
    End If
    MsgBox "Inside loop, x = " & x
    x = x - 0.2
Loop
MsgBox "Outside Loop, x = " & x
End Sub
