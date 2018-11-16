Dim astrPhoneList()
Dim strMsg
Dim lngIndex
Dim lngUBound

'Add the first row
ReDim Preserve astrPhoneList(2,0)
astrPhoneList(0, 0) = "Williams"
astrPhoneList(1, 0) = "Tony"
astrPhoneList(2, 0) = "404-555-6328"

'Add the second row
ReDim Preserve astrPhoneList(2,1)
astrPhoneList(0, 1) = "Carter"
astrPhoneList(1, 1) = "Ron"
astrPhoneList(2, 1) = "305-555-2514"

'Add the third row
ReDim Preserve astrPhoneList(2,2)
astrPhoneList(0, 2) = "Davis"
astrPhoneList(1, 2) = "Miles"
astrPhoneList(2, 2) = "212-555-5314"

'Add the fourth row
ReDim Preserve astrPhoneList(2,3)
astrPhoneList(0, 3) = "Hancock"
astrPhoneList(1, 3) = "Herbie"
astrPhoneList(2, 3) = "616-555-6943"

'Add the fifth row
ReDim Preserve astrPhoneList(2,4)
astrPhoneList(0, 4) = "Shorter"
astrPhoneList(1, 4) = "Wayne"
astrPhoneList(2, 4) = "853-555-0060"

'Loop through the array and display its contents
lngUBound = UBound(astrPhoneList, 2)

strMsg = "The phone list is:" & vbNewLine & vbNewLine

For lngIndex = 0 to lngUBound
   strMsg = strMsg & astrPhoneList(0, lngIndex) & ", "
   strMsg = strMsg & astrPhoneList(1, lngIndex) & " - "
   strMsg = strMsg & astrPhoneList(2, lngIndex) & vbNewLine
Next

MsgBox strMsg
