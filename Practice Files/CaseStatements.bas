Attribute VB_Name = "Module17"
Option Explicit
Sub read()
Dim chapter As Integer, message As String
chapter = 8
Select Case chapter
    Case 0
        message = "You didn't even try"
    Case 1, 2, 3
        message = "You could do better"
    Case 4 To 6
        message = "Not Bad"
    Case Is > 6
        message = "Very good"
    Case Else
        message = "Invalid number"
End Select
MsgBox message
End Sub
