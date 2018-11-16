Option Explicit

Dim FSO
Dim objStream

Const ForAppending = 8
Const TristateFalse = 0
Const FILE_NAME = "CREATE_FILE_TEST.TXT"

Set FSO = CreateObject("Scripting.FileSystemObject")

If Not FSO.FileExists(FILE_NAME) Then
    MsgBox "Could not find " & FILE_NAME & ".  " & _
        "Please run FSO_CREATE_FILE.VBS first."
Else
    Set objStream = FSO.OpenTextFile(FILE_NAME, _
        ForAppending, False, TristateFalse)
    
    With objStream
        .WriteLine "Appended Line 1"
        .WriteLine "Appended Line 2"
        .WriteLine "Appended Line 3"
        .Close
    End With
    
    MsgBox "Successfully appended to " & FILE_NAME & "."
End If