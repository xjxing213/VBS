Dim re, s, sc
Set re = New RegExp
s = InputBox("Type a string for the code to search")
re.Pattern = InputBox("Type in a pattern to find")
sc = InputBox("Type in a string to replace the pattern")
MsgBox re.Replace(s, sc)


