Option Explicit

Dim strFileName
Dim strFullPath

strFullPath = "C:\MyStuff\Documents\Personal\resume.doc"
strFileName = Right(strFullPath, _
    Len(strFullPath) - InStrRev(strFullPath,"\"))

MsgBox "The filename is: " & strFileName
