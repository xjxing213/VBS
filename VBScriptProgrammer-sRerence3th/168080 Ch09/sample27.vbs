Dim re, s
Set re = New RegExp
re.Global = True	
re.Pattern = "^.*<!--.*-->.*$"
s = " <title>A Title</title>  <!-- a title tag -->"
If re.Test(s) Then
    MsgBox "HTML comment tags found."
Else
    MsgBox "No HTML comment tags found."
End If
