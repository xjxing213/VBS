Dim intTest

intTest = CInt(100)
MsgBox "TypeName after initialization to 100: " & _
    TypeName(intTest)

intTest = intTest + 1000000
MsgBox "TypeName after adding 1,000,000: " & _
    TypeName(intTest)

intTest = intTest + 10000000000
MsgBox "TypeName after adding another 10,000,000,000: " & _
    TypeName(intTest)
