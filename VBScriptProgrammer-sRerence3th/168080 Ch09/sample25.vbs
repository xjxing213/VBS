Dim re, s
Set re = New RegExp
re.IgnoreCase = True
re.Pattern = "<(.*)>.*<\/\1>"
s = "<p>This is a paragraph</p>"
If re.Test(s) Then
    MsgBox "HTML element found."
Else
    MsgBox "No HTML element found."
End If