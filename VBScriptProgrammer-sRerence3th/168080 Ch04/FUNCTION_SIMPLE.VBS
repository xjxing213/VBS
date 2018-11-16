Option Explicit

Dim lngFirst
Dim lngSecond

lngFirst = 10
lngSecond = 20

MsgBox "The sum is: " & AddNumbers(lngFirst, lngSecond)

Function AddNumbers(lngFirstNumber, lngSecondNumber)

    AddNumbers = lngFirstNumber + lngSecondNumber

End Function