Option Explicit

Dim FSO

Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists("TEST_INPUT_FILE.TXT") Then
    FSO.CopyFile "TEST_INPUT_FILE.TXT", _
        "TEST_INPUT_FILE_COPY.TXT", True
End If
