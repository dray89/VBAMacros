Attribute VB_Name = "Module9"
Option Explicit

Sub main()
Dim a As Double, b As Double, x As Double, y As Double
a = 1: b = 2: x = 5: y = 6

Sheets("Sheet1").Select
Range("A10").Select
ActiveCell.Value = a
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = b
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = x
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = y

Call addy(a, b)
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = a


Call addy(x, y)
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = x

End Sub

Sub addy(m, n)
Dim answer As Double
answer = m + n
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = answer

m = 900
End Sub
