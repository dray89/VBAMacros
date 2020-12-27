Attribute VB_Name = "Module4"
Sub input2()

favcolor = inputbox("What is your favorite color?")
MsgBox "Your favorite color is " & favcolor

num1 = Val(inputbox("Enter number 1"))
num2 = Val(inputbox("Enter number 2"))

ans = num1 + num2

MsgBox "num1 = " & num1
MsgBox num1 & "+" & num2 & "=" & ans


End Sub

