Dim objFSO
Dim boolExists

Set objFSO = _
   WScript.CreateObject("Scripting.FileSystemObject")

boolExists = objFSO.FileExists("C:\autoexec.bat")
MsgBox boolExists
