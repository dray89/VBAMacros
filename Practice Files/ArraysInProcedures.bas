Attribute VB_Name = "Module26"
Option Explicit
Option Base 1
Sub main()
Const nrow As Integer = 2, ncol As Integer = 2
Dim i As Integer, j As Integer
Dim A(nrow, ncol) As Double, B(nrow, ncol) As Double
Dim C(nrow, ncol) As Double

A(1, 1) = 2: A(1, 2) = 3: A(2, 1) = 10: A(2, 2) = 20
B(1, 1) = 3: B(1, 2) = 1: B(2, 1) = -2: B(2, 2) = -100
Call addarrays((A), (B), (nrow), (ncol), C)
MsgBox C(1, 1) & " " & C(2, 2)

End Sub
