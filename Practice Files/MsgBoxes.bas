Attribute VB_Name = "Module3"
Sub inout()

Sheets("Sheet2").Select

Range("B2").Select
num1 = ActiveCell.Value

Range("B3").Select
num2 = ActiveCell.Value

ans = num1 + num2

MsgBox "VBA is fun"
MsgBox "num1 = " & num1
MsgBox num1 & "+" & num2 & "=" & ans

End Sub
