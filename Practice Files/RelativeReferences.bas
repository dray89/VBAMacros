Attribute VB_Name = "Module6"
Sub relref()

ActiveCell.Value = "VBA is fun"

ActiveCell.Offset(1, 0).Select
ActiveCell.Value = "Seriously"

ActiveCell.Offset(1, 1).Select
ActiveCell.Value = 100

ActiveCell.Offset(-1, 1).Select
ActiveCell.Value = 2.134

End Sub
