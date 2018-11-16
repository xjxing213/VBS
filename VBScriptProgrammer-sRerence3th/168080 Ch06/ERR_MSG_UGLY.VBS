Option Explicit

Dim x, y

x = InputBox("Please enter a number to divide into 100.")
If x <> 0 Then
    y = 100 / x
    MsgBox "100 divided by " & x & " is " & y & "."
Else
    Err.Raise vbObjectError + 15000, _
        "ERR_MSG_UGLY.VBS", _
        "Hey, stupid, you can't enter a zero. It will " & _
        "cause a divide by zero error."
End If
