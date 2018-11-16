Dim re, s
Set re = New RegExp
re.Pattern = "(\w+):\/\/([^/:]+)(:\d*)?([^# ]*)"
re.Global = True
re.IgnoreCase = True
s = "http://www.wrox.com:80/misc-pages/support.shtml"
MsgBox re.Replace(s, "$1")
MsgBox re.Replace(s, "$2")
MsgBox re.Replace(s, "$3")
MsgBox re.Replace(s, "$4")
