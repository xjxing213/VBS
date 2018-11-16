Option Explicit

Const LAST = 0
Const FIRST = 1
Const PHONE = 2

Dim dicPhoneList

Set dicPhoneList = CreateObject("Scripting.Dictionary")
FillPhoneList
SearchPhoneList

Sub FillPhoneList

    Dim strItemAdd(2,0)
    Dim strKey
    
    'Populate the list, using phone number as the key.
    'Add values to temp array, then add temp array to dictionary.
    strItemAdd(LAST, 0) = "Williams"
    strItemAdd(FIRST, 0) = "Tony"
    strItemAdd(PHONE, 0) = "404-985-6328"
    strKey = strItemAdd(PHONE, 0)
    dicPhoneList.Add strKey, strItemAdd
    
    strItemAdd(LAST, 0) = "Carter"
    strItemAdd(FIRST, 0) = "Ron"
    strItemAdd(PHONE, 0) = "305-781-2514"
    strKey = strItemAdd(PHONE, 0)
    dicPhoneList.Add strKey, strItemAdd
    
    strItemAdd(LAST, 0) = "Davis"
    strItemAdd(FIRST, 0) = "Miles"
    strItemAdd(PHONE, 0) = "212-963-5314"
    strKey = strItemAdd(PHONE, 0)
    dicPhoneList.Add strKey, strItemAdd
    
    strItemAdd(LAST, 0) = "Hancock"
    strItemAdd(FIRST, 0) = "Herbie"
    strItemAdd(PHONE, 0) = "616-752-6943"
    strKey = strItemAdd(PHONE, 0)
    dicPhoneList.Add strKey, strItemAdd
    
    strItemAdd(LAST, 0) = "Shorter"
    strItemAdd(FIRST, 0) = "Wayne"
    strItemAdd(PHONE, 0) = "853-238-0060"
    strKey = strItemAdd(PHONE, 0)
    dicPhoneList.Add strKey, strItemAdd

End Sub

Sub SearchPhoneList

    Dim strPhone
    Dim strItemRead
    
    strPhone = InputBox("Please enter a phone number " & _
        "(XXX-XXX-XXXX) with which to search the list.")
    
    If dicPhoneList.Exists(strPhone) Then
        strItemRead = dicPhoneList.Item(strPhone)
        MsgBox "We found that entry:" & vbNewLine & _
            vbNewLine & _
            "Last: " & strItemRead(LAST,0) & vbNewLine & _
            "First: " & strItemRead(FIRST,0) & vbNewLine & _
            "Phone: " & strItemRead(PHONE,0)
    Else
        MsgBox "That number was not found in the " & _
            "phone list."
    End If

End Sub