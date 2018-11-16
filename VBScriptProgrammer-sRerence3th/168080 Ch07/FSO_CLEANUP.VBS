Option Explicit

'This script deletes any test/temporary folders and files
'created by the Chapter 7 scripts.

Dim FSO

Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FolderExists("C:\TestVBScriptFolder") Then
    FSO.DeleteFolder("C:\TestVBScriptFolder")
End If

If FSO.FolderExists("C:\Program Files\TestVBScriptFolder") Then
    FSO.DeleteFolder("C:\Program Files\TestVBScriptFolder")
End If

If FSO.FileExists("TEST_INPUT_FILE_COPY.TXT") Then
    FSO.DeleteFile("TEST_INPUT_FILE_COPY.TXT")
End If

If FSO.FileExists("CREATE_FILE_TEST.TXT") Then
    FSO.DeleteFile("CREATE_FILE_TEST.TXT")
End If

MsgBox "Test folders and files created by these scripts have been deleted."
