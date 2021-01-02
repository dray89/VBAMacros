Attribute VB_Name = "Module25"
Option Explicit
Option Base 1

Sub main()
Const nrow As Integer = 2, ncol As Integer = 2
Dim c(nrow, ncol) As Integer
Dim i As Integer, j As Integer
For i = 1 To nrow
    For j = 1 To ncol
        c(i, j) = i * 2 + j ^ 2
    Next j
Next i
Sheets("Sheet1").Select: Range("A2").Select
For i = 1 To nrow
    For j = 1 To ncol
        ActiveCell.Value = c(i, j)
        ActiveCell.Offset(0, 1).Select
    Next j
    ActiveCell.Offset(1, -ncol).Select
Next i
     
End Sub
