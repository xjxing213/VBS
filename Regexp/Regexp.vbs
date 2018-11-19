'Dim re,s,sc

'Set re = new RegExp
're.pattern = "France"
'sc = "The rain in France falls mainly on the plains."
'MsgBox re.Replace(sc,"Spain")


Dim re,s
Set re = new RegExp
re.pattern = "\bin"
sc = "The rain in Spain falls mainly on the plains."
MsgBox re.Replace(sc,"in the country of")