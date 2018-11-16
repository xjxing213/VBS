Option Explicit

Const ORDER_ID_TO_FIND = "345"
Dim strCustID

strCustID = ""
strCustID = GetCustIDForOrder(ORDER_ID_TO_FIND)

If strCustID <> "" Then
    MsgBox "The CustomerID for Order " & _
        ORDER_ID_TO_FIND & " is: " & strCustID
Else
    MsgBox "We could not find OrderID " & _
        ORDER_ID_TO_FIND & "."
End If

Function GetCustIDForOrder(strOrderIDSeek)

    Const TristateFalse = 0
    Const ForReading = 1
    Const ORDER_FIELD = "OrderID="
    Const CUST_FIELD = "CustID="
    Const FILE_NAME = "TEST_INPUT_FILE.TXT"

    Dim FSO
    Dim objStream
    Dim strLine
    Dim lngFirstTab
    Dim lngSecondTab
    Dim strOrderID
    Dim strCustID
    
    strCustID = ""
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FileExists(FILE_NAME) Then
        Set objStream = FSO.OpenTextFile(FILE_NAME, _
            ForReading, False, TristateFalse)
    Else
        MsgBox "Could not find " & FILE_NAME & "."
        GetCustIDForOrder = ""
        Exit Function
    End If
    
    Do While Not objStream.AtEndOfStream
        strLine = objStream.ReadLine
        lngFirstTab = InStr(Len(ORDER_FIELD), strLine, _
            vbTab, vbBinaryCompare)
        strOrderID = Mid(strLine, Len(ORDER_FIELD) + 1, _
            lngFirstTab - Len(ORDER_FIELD) - 1)
        If strOrderID = strOrderIDSeek Then
            lngSecondTab = InStr(lngFirstTab + 1, strLine, _
                vbTab, vbBinaryCompare)
            strCustID = Mid(strLine, lngFirstTab + _
                Len(CUST_FIELD) + 1, _
                lngSecondTab - (lngFirstTab + _
                Len(CUST_FIELD)))
            Exit Do
        End If
    Loop

    objStream.Close
    GetCustIDForOrder = strCustID
    
End Function
