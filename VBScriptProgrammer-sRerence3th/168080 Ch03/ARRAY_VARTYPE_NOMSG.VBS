Dim strTest(1)
Dim lngSubType

strTest(0) = CLng(12)
strTest(1) = "Hello"

lngSubType = VarType(strTest) - vbArray

If lngSubType = vbVariant Then
  MsgBox "The Subtype is Variant."
Else
  MsgBox "The subtype is: " & lngSubType
End If
