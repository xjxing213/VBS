Option Explicit

Dim FSO
Dim objFolder

'You can run the script FSO.DELETE_FOLDER.VBS to
'delete the folder that this script creates.

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder("C:\")

If Not FSO.FolderExists("C:\TestVBScriptFolder") Then
    objFolder.SubFolders.Add "TestVBScriptFolder"
    MsgBox "The C:\TestVBScriptFolder folder was " & _
        "created successfully."
End If
