Dim re, s
Set re = New RegExp
re.Pattern = "[2-9]"
s = "Spain received 3 millimeters of rain last week."
MsgBox re.Replace(s, "many")


