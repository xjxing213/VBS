Option Explicit

GreetUser

Sub GreetUser
    MsgBox "Hello, " & GetUserName & _
        ". Welcome to our script."
End Sub

Function GetUserName
    GetUserName = InputBox("Please enter your name.")
End Function
