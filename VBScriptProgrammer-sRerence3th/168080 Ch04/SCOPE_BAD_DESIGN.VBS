Option Explicit

Private datToday

datToday = Date
AddOneDay
MsgBox "Tommorrow's date will be " & datToday & "."

Sub AddOneDay()
    datToday = DateAdd("d", 1, datToday)
End Sub
