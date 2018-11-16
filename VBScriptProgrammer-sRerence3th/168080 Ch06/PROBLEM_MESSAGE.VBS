Option Explicit

Dim x, y

x = InputBox("Please enter a number to divide into 100.")
If x <> 0 Then
    y = 100 / x
    MsgBox "100 divided by " & x & " is " & y & "."
Else
    MsgBox "Please enter a number other than zero."
End If