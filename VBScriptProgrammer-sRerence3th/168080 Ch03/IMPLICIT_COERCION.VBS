Dim lngTest

lngTest = CLng(100)
MsgBox "TypeName after initialization: " & TypeName(lngTest)

lngTest = lngTest + 1000
MsgBox "TypeName after adding 1000: " & TypeName(lngTest)

lngTest = lngTest * 50
MsgBox "TypeName after multiplying by 50: " & _
    TypeName(lngTest)

lngTest = "Hello"
MsgBox "TypeName after assigning value of 'Hello': " & _
    TypeName(lngTest)
