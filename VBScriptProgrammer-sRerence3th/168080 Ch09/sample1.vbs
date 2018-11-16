Dim re, s
Set re = New RegExp
re.Pattern = "France"
s = "The rain in France falls mainly on the plains."
MsgBox re.Replace(s, "Spain")


