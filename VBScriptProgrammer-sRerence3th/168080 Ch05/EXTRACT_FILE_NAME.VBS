Option Explicit

Dim lngIndex
Dim lngStrLen
Dim strFullPath
Dim strFileName

'This code will extract the filename from a path

strFullPath = "C:\Windows\Temp\Test\myfile.txt"
lngStrLen = Len(strFullPath)

For lngIndex = lngStrLen To 1 Step -1
    If Mid(strFullPath, lngIndex, 1) = "\" Then
        strFileName = Right(strFullPath, _
            lngStrLen - lngIndex)
        Exit For
    End If
Next

MsgBox "The filename is: " & strFileName
