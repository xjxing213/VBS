Dim re, s
Set re = New RegExp
re.IgnoreCase = True
re.Pattern = "http://(\w+[\w-]*\w+\.)*\w+"
s = "Some long string with http://www.wrox.com buried in it."
If re.Test(s) Then
    MsgBox "Found a URL."
Else
    MsgBox "No URL found."
End If
