Attribute VB_Name = "Module8"
Option Explicit

Sub math()

Dim num1 As Double, num2 As Double
Dim diff As Double, prod As Double

Sheets("Sheet1").Select
Range("B2").Select: num1 = ActiveCell.Value
ActiveCell.Offset(1, 0).Select: num2 = ActiveCell.Value

MsgBox num1 & " " & num2 & " " & diff & " " & prod

Call calcdp(num1, num2, diff, prod)

MsgBox num1 & " " & num2 & " " & diff & " " & prod
ActiveCell.Offset(2, 0).Select: ActiveCell.Value = diff
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = prod


End Sub

Sub calcdp(num1, num2, diff, prod)

diff = num1 - num2
prod = num1 * num2

End Sub
