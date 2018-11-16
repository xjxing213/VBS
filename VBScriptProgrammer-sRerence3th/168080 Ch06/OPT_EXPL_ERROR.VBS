Option Explicit

Dim datToday
Dim lngMonth

datToday = Date()
lngMonth = DatePart("m", datToday)
If lngMonth = 11 Then
    SpecialNovemberProcessing(datToday)
End If

Private Sub SpecialNovemberProcessing(datAny)
    lngDay = DatePart("d", datAny)
    MsgBox "Today is day " & lngDay & " of November."
End Sub