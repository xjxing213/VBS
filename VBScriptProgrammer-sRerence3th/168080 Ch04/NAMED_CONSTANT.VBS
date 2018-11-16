Option Explicit

Const GREETING = "Hello there, "

Dim strUserName

strUserName = InputBox("Please enter your name.")
If Trim(strUserName) <> "" Then
  MsgBox GREETING & strUserName & "."
End If
