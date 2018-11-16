Option Explicit

GreetUser GetUserName

Sub GreetUser(strUserName)
    MsgBox "Hello, " & strUserName & _
        ". Welcome to our script."
End Sub

Function GetUserName
    GetUserName = InputBox("Please enter your name.")
End Function
