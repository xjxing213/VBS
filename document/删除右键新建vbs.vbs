Set wshshell=CreateObject("wscript.shell") 
filetype=".vbs" 
wshshell.RegDelete "HKCR\"&filetype&"\shellnew\" 
MsgBox "command removed!"