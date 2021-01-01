Attribute VB_Name = "Module18"
Sub school()
Dim SAT As Integer, admitscore As Double
Dim yrsFrench As Integer, chess As String
SAT = inputbox("Enter SAT score, 0 - 2400")
yrsFrench = Val(inputbox("Number of years in French class?"))
chess = inputbox("Do you play chess? (y o n)")
admitscore = SAT / 20
'If play Chess,     +10 pts for each year in French class
'                   -15 pts if never taken French
' If not play chess, +5 pts for each year in French class
'                   - 20 pts if never taken French
If (chess = "y") Then
    If (yrsFrench > 0) Then
        admitscore = admitscore + 10 * yrsFrench
    Else
        admitscore = admitscore - 15
    End If
Else
    If (yrsFrench > 0) Then
        admitscore = admitscore + 5 * yrsFrench
    Else
        admitscore = admitscore - 20
    End If
    MsgBox "Hello"
End If
MsgBox "Your admission score is " & admitscore & "."

End Sub
