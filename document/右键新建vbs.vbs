filetype=".vbs" 
Set wshshell=CreateObject("wscript.shell") 
prg=readreg("HKCR\"&filetype&"\") 
prgname=readreg("HKCR\"&prg&"\") 
ask="Vbscript �ļ�" 
title="�����µ�vbs�ű�" 
prgname=InputBox(ask,title,prgname) 
wshshell.RegWrite "HKCR\"&prg&"\",prgname 
wshshell.RegWrite "HKCR\"&filetype&"\shellnew\nullfile","" 

Function readreg(key) 
On Error Resume Next 
readreg=wshshell.RegRead(key) 
If Err.Number>0 Then 
Error="error:ע����ֵ"""&key_&"""�����ҵ�!" 
MsgBox Error,vbCritical 
WScript.Quit 
End If 
End Function 