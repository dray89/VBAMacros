Attribute VB_Name = "Module2"
Sub mathfun()

Sheets("Sheet1").Select

Range("B3").Select
a = ActiveCell.Value

Range("B4").Select
b = ActiveCell.Value

Range("B5").Select
c = ActiveCell.Value

summy = a + b
divvy = b / c

Range("A7").Select
ActiveCell.Value = summy

Range("A8").Select
ActiveCell.Value = divvy

End Sub
