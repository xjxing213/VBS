Dim lngAge

lngAge = InputBox("Please enter your age in years.")

If IsNumeric(lngAge) Then
    lngAge = CLng(lngAge)
    lngAge = lngAge + 50
    MsgBox "In 50 years, you will be " & CStr(lngAge) & " years old."
Else
    MsgBox "Sorry, but you did not enter a valid number."
End If
