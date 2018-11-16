Option Explicit

Dim strWord

strWord = "aligator"
AppendSuffix strWord
MsgBox strWord

Sub AppendSuffix(ByRef strAny)
    strAny = strAny & "XXX"
End Sub