Option Explicit

Dim strSentence
Dim strVerb
Dim strNoun

'Start the sentence
strSentence = "The "

'Get a noun from the user
strNoun = InputBox("Please enter a noun (person, place, or thing).")

'Add the noun to the sentence
strSentence = strSentence & Trim(strNoun) & " "

'Get a verb from the user
strVerb = InputBox("Please enter a past tense verb.")

'Add the verb to the sentence
strSentence = strSentence & Trim(strVerb)

'Finish the sentence
strSentence = strSentence & "."

'Display the sentence
MsgBox strSentence
