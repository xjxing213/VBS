Dim objFSO
Dim boolExists

Set objFSO = _
    WScript.CreateObject("Scripting.FileSystemObject")

If IsObject(objFSO) Then
    boolExists = objFSO.FileExists("C:\autoexec.bat")
    MsgBox boolExists
End If
