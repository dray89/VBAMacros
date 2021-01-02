Attribute VB_Name = "Module24"
Option Explicit
Option Base 1
Sub main()
Const N As Integer = 12
Dim a(N) As Integer, i As Integer

For i = 1 To N
    a(i) = i ^ 2
Next i
MsgBox a(1) & " " & a(2) & " " & a(N)
End Sub
