Option Explicit

Private datToday

datToday = Date
MsgBox "Tommorrow's date will be " & AddOneDay(datToday) & "."

Function AddOneDay(datAny)

    Dim datResult
    
    datResult = DateAdd("d", 1, datAny)
    AddOneDay = datResult

End Function
