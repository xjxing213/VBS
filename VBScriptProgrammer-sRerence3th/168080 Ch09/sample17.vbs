Dim re, s
Set re = New RegExp
re.Pattern = "(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)"
s = "VBScript is not very cool."
MsgBox re.Replace(s, "$1 $2 $4 $5")
