Dim re, s
Set re = New RegExp
re.Pattern = "\bin"
re.Global = True
re.IgnoreCase = True
s = "The rain In Spain falls mainly on the plains."
MsgBox re.Replace(s, "in the country of")


