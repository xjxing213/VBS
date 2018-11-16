Dim datBirth

datBirth = InputBox("Please enter the date on which you were born.")

If IsDate(datBirth) Then
    datBirth = CDate(datBirth)
    MsgBox "You were born on day " & Day(datBirth) & _
        " of month " & Month(datBirth) & " in the year " & _
        Year(datBirth) & "."
Else
    MsgBox "Sorry, but you did not enter a valid date."
End If
