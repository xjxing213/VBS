Option Explicit

Dim FSO

Set FSO = CreateObject("Scripting.FileSystemObject")

'If C:\TestVBScriptFolder does not exist, you can
'run FSO_CREATE_FOLDER.VBS to create it. The script
'FSO_CLEANUP.VBS will delete any of the test folders
'and files created by these scripts.
If FSO.FolderExists("C:\TestVBScriptFolder") Then
    FSO.CopyFolder "C:\TestVBScriptFolder", _
        "C:\Program Files\", True
End If
