Option Explicit

Dim lngA
Dim lngB

lngA = 1
lngB = 1

ByRefByValExample lngA, lngB

MsgBox "lngA = " & lngA & vbNewLine & _
    "lngB = " & lngB

Sub ByRefByValExample(ByRef lngFirst, ByVal lngSecond)
    lngFirst = lngFirst + 1
    lngSecond = lngSecond + 1
End Sub