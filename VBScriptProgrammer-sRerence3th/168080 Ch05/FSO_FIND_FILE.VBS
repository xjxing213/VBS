Option Explicit

Dim objFSO
Dim objRootFolder
Dim objFileLoop
Dim boolFoundIt

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objRootFolder = objFSO.GetFolder("C:\")
Set objFSO = Nothing

boolFoundIt = False
For Each objFileLoop In objRootFolder.Files
    If UCase(objFileLoop.Name) = "AUTOEXEC.BAT" Then
        boolFoundIt = True
        Exit For
    End If
Next
Set objFileLoop = Nothing
Set objRootFolder = Nothing

If boolFoundIt Then
    MsgBox "We found your AUTOEXEC.BAT file in " & _
        "the C:\ directory."
Else
    MsgBox "We could not find AUTOEXEC.BAT in " & _
        "the C:\ directory."
End If
