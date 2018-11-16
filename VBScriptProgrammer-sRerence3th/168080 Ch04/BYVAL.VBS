Option Explicit

Dim strWord

strWord = "aligator"
strWord = AppendSuffix(strWord)
MsgBox strWord

Function AppendSuffix(ByVal strAny)
    AppendSuffix = strAny & "XXX"
End Function