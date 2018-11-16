Dim YourName 
Dim Greeting

YourName = InputBox("Hello!  What is your name?")

If YourName = "" Then
    Greeting = "OK.  You don't want to tell me your name."
Else
    Greeting = "Hello, " & YourName & ", great to meet you."
End If

MsgBox Greeting
