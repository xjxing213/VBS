Dim re, s
Set re = New RegExp
re.Pattern = "in"
re.Global = True
s = "The rain in Spain falls mainly on the plains."
MsgBox re.Replace(s, "in the country of")


