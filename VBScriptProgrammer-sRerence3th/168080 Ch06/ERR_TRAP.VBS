Option Explicit

Dim x

On Error Resume Next
x = GetValueFromDatabase()
If Err.Number = 0 Then
    MsgBox "The value of x is: " & x
Else
    DisplayError Err.Number, Err.Source, _
        Err.Description
End If
On Error GoTo 0

Private Function GetValueFromDatabase()
    
    'Deliberately create an error for
    'demonstration purposes.
    Err.Raise vbObjectError + 1000, _
        "GetValueFromDatabase()", _
        "A database error occurred."

End Function

Private Sub DisplayError(lngNumber, strSource, strDescription)
    MsgBox "An error occurred. Please write down " & _
        "the error information displayed below " & _
        "and contact your system administrator:" & _
        vbNewLine & vbNewLine & _
        "Error Description: " & strDescription & vbNewLine & _
        "Error Number: " & lngNumber & vbNewLine & _
        "Error Source: " & strSource, _
        vbExclamation
End Sub