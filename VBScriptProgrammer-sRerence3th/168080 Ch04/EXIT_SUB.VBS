Option Explicit

GreetUser InputBox("Please enter your name.")

Sub GreetUser(strUserName)
    If IsNumeric(strUserName) or IsDate(strUserName) Then
        MsgBox "That is not a legal name."
        Exit Sub
    End If

    MsgBox "Hello, " & strUserName & _
        ". Welcome to our script."
End Sub
