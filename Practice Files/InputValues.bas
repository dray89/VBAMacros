Attribute VB_Name = "Module1"
Sub myfirstmacro()

Sheets("Sheet1").Select
Range("A3").Select
ActiveCell.Value = 5

Sheets("Sheet3").Select
Range("B2").Select
ActiveCell.Value = 62

End Sub
