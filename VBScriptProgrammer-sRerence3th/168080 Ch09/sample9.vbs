Dim re, s
Set re = New RegExp
re.Pattern = "\d{3}"
s = "Spain received 100 millimeters of rain in the last 2 weeks."
MsgBox re.Replace(s, "a whopping number of")


