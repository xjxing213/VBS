Dim re, objMatch, colMatches, sMsg
Set re = New RegExp
re.Global = True	
re.Pattern = "http://(\w+[\w-]*\w+\.)*\w+" 
s = "http://www.kingsley-hughes.com is a valid web address. And so is "
s = s & vbCrLf & "http://www.wrox.com.  As is " 
s = s & vbCrLf & "http://www.wiley.com." 
Set colMatches = re.Execute(s)	
sMsg = ""
For Each objMatch in colMatches		
    sMsg = sMsg & "Match of " & objMatch.Value 
    sMsg = sMsg & ", found at position " & objMatch.FirstIndex & " of the string. "
    sMsg = sMsg & "The length matched is "
    sMsg = sMsg & objMatch.Length & "." & vbCrLf
Next
MsgBox sMsg
