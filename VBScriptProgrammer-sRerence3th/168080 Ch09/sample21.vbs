Dim re, objMatch, colMatches, sMsg
Set re = New RegExp
re.Global = True	
re.Pattern = "http://(\w+[\w-]*\w+\.)*\w+" 
s = "http://www.kingsley-hughes.com is a valid web address. And so is "
s = s & vbCrLf & "http://www.wrox.com.  As is " 
s = s & vbCrLf & "http://www.wiley.com." 
Set colMatches = re.Execute(s)	
MsgBox colMatches.item(0)
MsgBox colMatches.item(1)
MsgBox colMatches.item(2)