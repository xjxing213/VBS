Option Explicit

GreetUser InputBox("Please enter your name.")

Sub GreetUser(strUserName)
    If IsNumeric(strUserName) or IsDate(strUserName) Then
        MsgBox "That is not a legal name."
    Else
        MsgBox "Hello, " & strUserName & _
            ". Welcome to our script."
    End If
End Sub
