Attribute VB_Name = "Module14"
Option Explicit

Sub main()
Dim x As Double, y As Double
x = 40: y = 50
If (x > 10) Then
    MsgBox "Inside IF structure"
    y = 300
End If
MsgBox "y = " & y
End Sub
