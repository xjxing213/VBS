Option Explicit

Dim FSO

'You can run the script FSO.DELETE_FOLDER.VBS to
'delete the folder that this script creates.

Set FSO = CreateObject("Scripting.FileSystemObject")

FSO.CreateFolder("C:\TestVBScriptFolder")

MsgBox "The C:\TestVBScriptFolder folder was " & _
    "created successfully."
