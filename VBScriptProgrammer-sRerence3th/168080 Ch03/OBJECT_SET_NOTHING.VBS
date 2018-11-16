Dim objFSO
Dim boolExists

Set objFSO = _
   WScript.CreateObject("Scripting.FileSystemObject")

If IsObject(objFSO) Then
    boolExists = objFSO.FileExists("C:\autoexec.bat")
    MsgBox boolExists
    Set objFSO = Nothing
    If objFSO Is Nothing Then
        MsgBox "The object has been destroyed, which " & _
            "frees up the resources it was using."
    End If
End If
