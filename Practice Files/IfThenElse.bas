Attribute VB_Name = "Module15"
Option Explicit

Sub main()
Dim costpizza As Double, x As Double
costpizza = 11
If (costpizza > 10) Then
    MsgBox "The pizza is too expensive"
    x = 700
Else
    MsgBox "Buy the pizza"
    x = 22
End If
MsgBox "x is " & x

End Sub
