Option Explicit

Dim strSentence

strSentence = "The "
strSentence = strSentence & GetNoun & " "
strSentence = strSentence & GetVerb
strSentence = strSentence & GetPeriod
DisplayMessage strSentence

Function GetNoun
    GetNoun = Trim(InputBox("Please enter a noun (person, place, or thing)."))
End Function

Function GetVerb
    GetVerb = Trim(InputBox("Please enter a past tense verb."))
End Function

Function GetPeriod
    GetPeriod = "."
End Function

Sub DisplayMessage(strAny)
    MsgBox strAny
End Sub