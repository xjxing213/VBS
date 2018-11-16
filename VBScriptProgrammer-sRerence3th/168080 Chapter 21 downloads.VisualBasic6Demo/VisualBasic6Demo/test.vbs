Option Explicit

Dim datToday
Dim lngMonth

datToday = Date()
lngMonth = DatePart("m", datToday)
If lngMonth = 11 Then
    SpecialNovemberProcessing(datToday)
End If

Private Sub SpecialNovemberProcessing(datAny)
    datDay = DatePart("d", datAny)
    MsgBox "Today is day " & datDay & " of November."
End Sub