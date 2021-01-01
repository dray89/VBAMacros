Attribute VB_Name = "Module19"
Option Explicit
Sub loopy()
Dim i As Integer, sum As Integer, avg As Integer
Dim n As Integer
n = 4
sum = 0
For i = n To 1 Step -1
    sum = sum + i
    MsgBox "i =" & i & " sum = " & sum
Next i
avg = sum / n
MsgBox "i = " & i & " sum = " & sum & " n = " & n

End Sub
