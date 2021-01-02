Attribute VB_Name = "Module27"
Option Explicit

Sub main()
Dim b(3) As Double
ReDim f(2) As Double
f(0) = 5: f(1) = 3: f(2) = -1
MsgBox f(0) & f(2)
ReDim Preserve f(4) As Double
f(3) = 11: f(4) = 13
MsgBox f(0) & f(4)
ReDim Preserve f(3) As Double
MsgBox f(0) & f(3)
End Sub
