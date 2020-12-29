Attribute VB_Name = "Module16"
Option Explicit

Sub eat()

Dim funds As Double
funds = 26

MsgBox "Before If structure"
If (funds >= 25) Then
    MsgBox "Eat at the fancy restaurant"
ElseIf (funds >= 10 And funds < 25) Then
    MsgBox "Eat fast food"
ElseIf (funds >= 0 And funds < 10) Then
    MsgBox "Eat instant noodles"
ElseIf (funds > 1) Then
    MsgBox "Never runs because prior test condition is true"
Else
    MsgBox "funds must be an invalid number if not one of the above"
End If
MsgBox "After If structure"
End Sub
