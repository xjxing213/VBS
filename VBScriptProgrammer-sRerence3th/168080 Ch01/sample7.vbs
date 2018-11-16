Dim Counter
Dim WordLength
Dim InputWord
Dim WordBuilder

InputWord = InputBox ("Type in a word of phrase to use")

WordLength = Len(InputWord)

For Counter = 1 to WordLength
    MsgBox Mid(InputWord, Counter, 1)
    WordBuilder = WordBuilder & Mid(InputWord, Counter, 1)
Next

MsgBox WordBuilder & " contains " & WordLength & " characters."
