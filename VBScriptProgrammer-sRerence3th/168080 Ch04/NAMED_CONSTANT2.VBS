Option Explicit

Const RESPONSE_YES = "YES"
Const RESPONSE_NO = "NO"

Dim strResponse

strResponse = InputBox("Is today a Tuesday? Please answer Yes or No.")
strResponse = UCase(strResponse)
If strResponse = RESPONSE_YES Then
  MsgBox "I love Tuesdays."
ElseIf strResponse = RESPONSE_NO Then
  MsgBox "I will gladly pay you Tuesday for a hamburger today."
Else
  MsgBox "Invalid response."
End If
