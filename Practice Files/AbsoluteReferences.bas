Attribute VB_Name = "Module5"
Sub absoluteref()

Sheets("Sheet1").Select
Range("A1").Select
ActiveCell.Value = "VBA is fun"

Range("A2").Select
ActiveCell.Value = "Seriously."

Range("B3").Select
ActiveCell.Value = 100

Range("C2").Select
ActiveCell.Value = 2.134

End Sub
