Option Explicit

Dim objFSO
Dim objStream
Dim strText

Set objFSO = _
    WScript.CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile("testfile.txt")
Set objFSO = Nothing

strText = ""
Do While Not objStream.AtEndOfStream
    strText = strText & objStream.ReadLine & vbNewLine
Loop
Set objStream = Nothing

If strText <> "" Then
    MsgBox strText
Else
    MsgBox "The file is empty."
End If
