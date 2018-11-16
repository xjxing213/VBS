Option Explicit

Dim objFSO
Dim objAutoExecFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objAutoExecFile = objFSO.GetFile("C:\autoexec.bat")

MsgBox "The autoexec.bat was last changed on: " & _
    objAutoExecFile.DateLastModified
