Dim re, s, colMatches, objMatch, sMsg
Set re = New RegExp
re.Global = True	
re.Pattern = "^[ \t]*$"
s = " "
Set colMatches = re.Execute(s)
sMsg = ""
For Each objMatch in colMatches
    sMsg = sMsg & "Blank line found at position " & objMatch.FirstIndex & " of the string."
Next
MsgBox sMsg
