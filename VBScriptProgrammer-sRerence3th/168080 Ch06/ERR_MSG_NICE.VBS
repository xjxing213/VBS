Option Explicit

Dim x, y

x = InputBox("Please enter a number to divide into 100.")

On Error Resume Next
y = DivideNumbers(100, x)
If Err.Number = (vbObjectError + 15000) Then
    On Error GoTo 0
    MsgBox "Please enter a number other than zero."
Else
    On Error GoTo 0
    MsgBox "100 divided by " & x & " is " & y & "."
End If

Private Function DivideNumbers(dblNumber, dblDivideBy)

    If dblDivideBy = 0 Then
        Err.Raise vbObjectError + 15000, _
            "ERR_MSG_NICE.DivideNumbers()", _
            "Division by zero not allowed."
    Else
        DivideNumbers = dblNumber / dblDivideBy
    End If

End Function