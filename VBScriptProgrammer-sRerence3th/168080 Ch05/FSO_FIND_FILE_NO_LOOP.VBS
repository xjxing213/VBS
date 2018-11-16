Option Explicit

Dim objFSO

Set objFSO = _
    WScript.CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists("C:\AUTOEXEC.BAT") Then
    MsgBox "We found your AUTOEXEC.BAT file in the " & _
        "C:\ directory."
Else
    MsgBox "We could not find AUTOEXEC.BAT in " & _
        "the C:\ directory."
End If

Set objFSO = Nothing
