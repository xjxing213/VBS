Dim re, s
Set re = New RegExp
re.Pattern = "http://(\w+[\w-]*\w+\.)*\w+"
s = "http://www.kingsley-hughes.com is a valid web address.  And so is "
s = s & vbCrLf & "http://www.wrox.com.  And " 
s = s & vbCrLf & "http://www.pc.ibm.com - even with 4 levels." 
MsgBox re.Replace(s, "** TOP SECRET! **") 
