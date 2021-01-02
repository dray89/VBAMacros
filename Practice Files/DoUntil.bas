Attribute VB_Name = "Module21"
Option Explicit

Sub loopy()
Dim i As Integer, n As Integer, sum As Integer
Dim avg As Double
sum = 0: n = 3: i = 1
'Do While (i<=n)
Do Until (i > n)
    sum = sum + i
    MsgBox "i = " & i & " Sum = " & sum
    i = i + 1
Loop
avg = sum / n
MsgBox "i=" & i & " sum = " & sum & " average = " & avg
End Sub
