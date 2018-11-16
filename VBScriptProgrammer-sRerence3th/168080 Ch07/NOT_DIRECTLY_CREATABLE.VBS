Option Explicit

Dim objFSO
Dim objFolder

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\")

If objFolder.IsRootFolder Then
    MsgBox "We have opened the root folder."
End If