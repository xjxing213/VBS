Attribute VB_Name = "Registry"
Option Explicit

' $Revision: 2 $
' $Author: Davids $
' $Date: 4/15/98 4:55p $

' registry information
Private Const APP_NAME          As String = "Wrox Press"
Private Const APP_SECTION       As String = "Wrox Car Co"
Private Const NOT_FOUND         As String = "<Not Found>"

Public Sub RegistrySave(ByVal sKey As String, ByVal sValue As String)
'
' Purpose:      Save a value to the registry
' Arguments:    sKey        Key to save
'               sValue      Value of key
' Returns:      none

    SetKeyValue sKey, sValue
    
End Sub

Public Function RegistryRestore(ByVal sKey As String, ByVal sDefault As String) As String
'
' Purpose:      Restore a value from the registry
' Arguments:    sKey        Key to restore
'               sDefault    Default value if key doesn't exist
' Returns:      none

    Dim sValue      As String

    sValue = GetKeyValue(sKey)

    If sValue = "" Then
        ' not found, so return default
        sValue = sDefault
    End If

    RegistryRestore = sValue

End Function

