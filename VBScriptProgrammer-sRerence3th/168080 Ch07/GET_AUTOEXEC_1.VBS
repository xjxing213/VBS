Option Explicit

Dim objFSO
Dim objCDrive
Dim objRootFolder
Dim objFileLoop
Dim objAutoExecFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objCDrive = objFSO.GetDrive("C")
Set objRootFolder = objCDrive.RootFolder

For Each objFileLoop in objRootFolder.Files
    If UCase(objFileLoop.Name) = "AUTOEXEC.BAT" Then
        Set objAutoExecFile = objFileLoop
        Exit For
    End If
Next

If IsObject(objAutoExecFile) Then
    MsgBox "The autoexec.bat was last changed on: " & _
        objAutoExecFile.DateLastModified
Else
    MsgBox "Could not find autoexec.bat."
End If
