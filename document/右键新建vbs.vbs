filetype=".vbs" 
Set wshshell=CreateObject("wscript.shell") 
prg=readreg("HKCR\"&filetype&"\") 
prgname=readreg("HKCR\"&prg&"\") 
ask="Vbscript 文件" 
title="创建新的vbs脚本" 
prgname=InputBox(ask,title,prgname) 
wshshell.RegWrite "HKCR\"&prg&"\",prgname 
wshshell.RegWrite "HKCR\"&filetype&"\shellnew\nullfile","" 

Function readreg(key) 
On Error Resume Next 
readreg=wshshell.RegRead(key) 
If Err.Number>0 Then 
Error="error:注册表键值"""&key_&"""不能找到!" 
MsgBox Error,vbCritical 
WScript.Quit 
End If 
End Function 