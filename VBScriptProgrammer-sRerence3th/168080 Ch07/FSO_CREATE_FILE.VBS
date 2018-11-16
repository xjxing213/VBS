Option Explicit

Dim FSO
Dim objStream

Const TristateFalse = 0
Const FILE_NAME = "CREATE_FILE_TEST.TXT"

Set FSO = CreateObject("Scripting.FileSystemObject")

Set objStream = FSO.CreateTextFile(FILE_NAME, _
    True, TristateFalse)

With objStream
    .WriteLine "Test Line 1"
    .WriteLine "Test Line 2"
    .WriteLine "Test Line 3"
    .Close
End With

MsgBox "Successfully created " & FILE_NAME & "."
