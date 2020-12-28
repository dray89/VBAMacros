Attribute VB_Name = "Module10"
Option Explicit

Sub main()

Dim a As Double, b As Double, c As Double, d As Double
a = 2: b = 3: c = -1:
d = funky(a, b, c) + 10

MsgBox d & " " & funky(c, a, b)

End Sub

Function funky(a, b, c) As Double
MsgBox "In funky: " & a & " " & b & " " & c
funky = a * b + c

End Function
