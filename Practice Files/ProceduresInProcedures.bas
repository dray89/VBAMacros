Attribute VB_Name = "Module12"
Option Explicit

Sub main()
Dim a As Double, b As Double
Dim C As String
a = 1: b = 2: C = "Hi"
MsgBox a & "+" & b & "=" & addy(a, b)
Call jumble(a, (b), C)
MsgBox a & " " & b & " " & C
End Sub

Sub jumble(x, y, z)
Dim w As Double
Call calcw(x, y, w)
y = 10: x = w + addy(x, y)
z = z & "Bye"
End Sub

Function addy(y, z) As Double
addy = y + z
End Function

Sub calcw(n1, n2, ans)
ans = n1 * n2
End Sub
