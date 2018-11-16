Option Explicit

Dim objFSO
Dim objFolder

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Program Files")

With objFolder
    MsgBox "Here are some properties from the Folder object:" & _
        vbNewLine & vbNewLine & _
        "Path: " & .Path & vbNewLine & _
        "DateCreated: " & .DateCreated & vbNewLine & _
        "DateLastAccessed: " & .DateLastAccessed & vbNewLine & _
        "DateLastModified: " & .DateLastModified
End With