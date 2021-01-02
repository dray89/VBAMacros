Attribute VB_Name = "Module23"
Option Explicit
Sub meanscore()
Dim i As Integer, firstrow As Integer, lastrow As Integer
Dim numscores As Integer, mean As Double

Sheets("Sheet1").Select
Range("B2").Select
firstrow = ActiveCell.Row
Selection.End(x1Down).Select
lastrow = ActiveCell.Row
numscores = lastrow - firstrow + 1

Range("B2").Select
mean = 0
For i = 1 To numscores
    mean = mean + ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
Next i
mean = mean / numscores
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = mean

End Sub
