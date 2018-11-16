Option Explicit

Dim objList

FillPhoneList
On Error Resume Next
objList.DisplayEntry(GetNumberFromUser)

If Err.Number <> 0 Then
    If Err.Number = vbObjectError + 32002 Then
        MsgBox "That phone number is not in the list.", _
            vbInformation
    Else
        DisplayError Err.Number, Err.Source, _
            Err.Description
    End If
End If

Public Sub FillPhoneList
    Dim objNewEntry
    
    Set objList = New PhoneList
    
    Set objNewEntry = New ListEntry
    With objNewEntry
        .LastName = "Williams"
        .FirstName = "Tony"
        .PhoneNumber = "404-555-6328"
    End With
    objList.AddEntry objNewEntry
    Set objNewEntry = Nothing
    
    Set objNewEntry = New ListEntry
    With objNewEntry
        .LastName = "Carter"
        .FirstName = "Ron"
        .PhoneNumber = "305-555-2514"
    End With
    objList.AddEntry objNewEntry
    Set objNewEntry = Nothing

    Set objNewEntry = New ListEntry
    With objNewEntry
        .LastName = "Davis"
        .FirstName = "Miles"
        .PhoneNumber = "212-555-5314"
    End With
    objList.AddEntry objNewEntry
    Set objNewEntry = Nothing

    Set objNewEntry = New ListEntry
    With objNewEntry
        .LastName = "Hancock"
        .FirstName = "Herbie"
        .PhoneNumber = "616-555-6943"
    End With
    objList.AddEntry objNewEntry
    Set objNewEntry = Nothing

    Set objNewEntry = New ListEntry
    With objNewEntry
        .LastName = "Shorter"
        .FirstName = "Wayne"
        .PhoneNumber = "853-555-0060"
    End With
    objList.AddEntry objNewEntry
    Set objNewEntry = Nothing

End Sub

Public Function GetNumberFromUser
    GetNumberFromUser = InputBox("Please enter " & _
        "a phone number (XXX-XXX-XXXX) with " & _
        "which to search the list.")
End Function

Class PhoneList
    
    Private objDict

    Private Sub Class_Initialize
        Set objDict = CreateObject("Scripting.Dictionary")
    End Sub
    Private Sub Class_Terminate
        Set objDict = Nothing
    End Sub

    Public Property Get ListCount
        ListCount = objDict.Count
    End Property

    Public Function EntryExists(strPhoneNumber)
        EntryExists = _
            objDict.Exists(strPhoneNumber)
    End Function

    Public Sub AddEntry(objListEntry)
        If TypeName(objListEntry) <> "ListEntry" Then
            Err.Raise vbObjectError + 32000, "PhoneList", _
                "Only ListEntry objects can be stored " & _
                "in a PhoneList class."
        End If
        'We use the PhoneNumber property as the key.
        If Trim("" & objListEntry.PhoneNumber) = "" Then
            Err.Raise vbObjectError + 32001, "PhoneList", _
                "A ListEntry object must have a phone number " & _
                "to be added to the phone list."
        End If

        objDict.Add objListEntry.PhoneNumber, objListEntry
    End Sub

    Public Sub DisplayEntry(strPhoneNumber)
        Dim objEntry
        
        If objDict.Exists(strPhoneNumber) Then
            Set objEntry = objDict(strPhoneNumber)
            objEntry.DisplayEntry
        Else
            Err.Raise vbObjectError + 32002, "PhoneList", _
                "The phone number '" & strPhoneNumber & _
                    "' is not in the list."
        End If
    End Sub

End Class

Class ListEntry

    Private mstrLast
    Private mstrFirst
    Private mstrPhone
    
    Public Property Let LastName(strLastName)
        If IsNumeric(strLastName) or IsDate(strLastName) Then
            Err.Raise vbObjectError + 32003, "ListEntry", _
                "The LastName property may not " & _
                "be a number or date."
        End If 
    
        mstrLast = strLastName
    End Property
    Public Property Get LastName
        LastName = mstrLast
    End Property

    Public Property Let FirstName(strFirstName)
        If IsNumeric(strFirstName) or IsDate(strFirstName) Then
            Err.Raise vbObjectError + 32004, "ListEntry", _
                "The FirstName property may not " & _
                "be a number or date."
        End If 

        mstrFirst = strFirstName
    End Property
    Public Property Get FirstName
        FirstName = mstrFirst
    End Property

    Public Property Let PhoneNumber(strPhoneNumber)
        mstrPhone = strPhoneNumber
    End Property
    Public Property Get PhoneNumber
        PhoneNumber = mstrPhone
    End Property

    Public Sub DisplayEntry
        MsgBox "Phone list entry:" & vbNewLine & _
            vbNewLine & _
            "Last: " & mstrLast & vbNewLine & _
            "First: " & mstrFirst & vbNewLine & _
            "Phone: " & mstrPhone
    End Sub

End Class

Private Sub DisplayError(lngNumber, strSource, strDescription)
    MsgBox "An error occurred. Please write down " & _
        "the error information displayed below " & _
        "and contact your system administrator:" & _
        vbNewLine & vbNewLine & _
        "Error Description: " & strDescription & vbNewLine & _
        "Error Number: " & lngNumber & vbNewLine & _
        "Error Source: " & strSource, _
        vbExclamation
End Sub