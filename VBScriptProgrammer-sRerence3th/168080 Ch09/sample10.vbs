Dim re, s
Set re = New RegExp
re.Pattern = "\d"
s = "Spain received 100 millimeters of rain in the last 2 weeks."
MsgBox re.Replace(s, "a whopping number of")


