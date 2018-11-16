Dim re, s
Set re = New RegExp
re.Pattern = "\d"
s = "a b c d e f 1 g 2 h ... 10 z"
MsgBox re.Replace(s, "a number")


