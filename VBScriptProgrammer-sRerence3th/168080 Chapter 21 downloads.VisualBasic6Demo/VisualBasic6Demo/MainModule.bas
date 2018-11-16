Attribute VB_Name = "MainModule"
Option Explicit

Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

Sub Main()
    frmConnection.Show
End Sub

Function InitScriptControl(frmForm As Form) As ScriptControl
    Dim objSC As ScriptControl
    Dim fileName As String, intFnum As Integer
    Dim objShare As New CShared
    
    ' create a new instance of the control
    Set objSC = New ScriptControl
    objSC.Language = "VBScript"
    objSC.AllowUI = True
    Set objShare.Form = frmForm
    objSC.AddObject "share", objShare, True
   
    ' load the code into the script control
    fileName = App.Path & "\regeditor.scp"
    intFnum = FreeFile
    Open fileName For Input As #intFnum

    objSC.AddCode Input$(LOF(intFnum), intFnum)
    Close #intFnum
    
    ' return to the caller
    Set InitScriptControl = objSC
    
End Function

Function SetKeyValue(strSubPath As String, strKeyName As String, strValue As String) As Boolean
   Dim lngResult As Integer
   Dim lngThisKey As Long
   lngResult = RegCreateKey(HKEY_LOCAL_MACHINE, strSubPath & strKeyName, lngThisKey)
   If lngResult = ERROR_SUCCESS Then
      lngResult = RegSetValue(lngThisKey, vbNullString, REG_SZ, strValue, Len(strValue))
   Else
      Err.Raise vbObjectError + 101, "WCCCars: SetKeyValue", "Cannot set Registry value"
   End If
   lngResult = RegCloseKey(lngThisKey)
End Function

Public Sub RegistrySave(ByVal strSubPath As String, ByVal sKey As String, ByVal sValue As String)
'
' Purpose:      Save a value to the registry
' Arguments:    sKey        Key to save
'               sValue      Value of key
' Returns:      none

    SetKeyValue strSubPath, sKey, sValue
    
End Sub

