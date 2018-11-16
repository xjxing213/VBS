Dim re, s, objMatch, colMatches
Set re = New RegExp
re.Pattern = "\([0-9]{3}\)[0-9]{3}-[0-9]{4}"
re.Global = True
re.IgnoreCase = True
s = InputBox("Enter your phone number in the following Format (XXX)XXX-XXXX:")
If re.Test(s) Then
    MsgBox "Thank you!"
Else
    MsgBox "Sorry but that number is not in a valid format."
End If
