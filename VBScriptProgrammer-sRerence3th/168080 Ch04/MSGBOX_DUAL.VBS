Option Explicit

Dim lngResponse
Dim strUserName

lngResponse = MsgBox("Would you like a greeting?", vbYesNo)

If lngResponse = vbYes Then
    strUserName = GetUserName
    GreetUser strUserName
End If

Sub GreetUser(strUserName)
    MsgBox "Hello, " & strUserName & _
        ". Welcome to our script."
End Sub

Function GetUserName
    GetUserName = InputBox("Please enter your name.")
End Function
