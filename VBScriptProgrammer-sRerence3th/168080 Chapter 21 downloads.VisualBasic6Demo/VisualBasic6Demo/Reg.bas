Attribute VB_Name = "RegHandler"
Option Explicit

' Note: This module introduced as SaveSetting and GetSetting write to the
' HKEY_CURRENT_USER path, which only applies if a user is currently logged
' in.  The components are intended to run on a server, which may not have
' a current user, so we've modified the registry position to HKEY_LOCAL_MACHINE

Const OUR_SUBKEY_PATH = "SOFTWARE\Wrox Press\Wrox Car Co\"
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

Function GetKeyValue(strKeyName As String) As String
   Dim lngResult As Integer
   Dim lngThisKey As Long
   Dim lngLength As Long
   Dim strValue As String
   strValue = String(1024, "0")
   lngLength = 1023
   lngResult = RegOpenKey(HKEY_LOCAL_MACHINE, OUR_SUBKEY_PATH & strKeyName, lngThisKey)
   If lngResult = ERROR_SUCCESS Then lngResult = RegQueryValue(lngThisKey, vbNullString, strValue, lngLength)
   If lngResult = ERROR_SUCCESS Then
      GetKeyValue = Left(strValue, lngLength - 1)
   Else
      GetKeyValue = ""
   End If
   lngResult = RegCloseKey(lngThisKey)
End Function

Function SetKeyValue(strKeyName As String, strValue As String) As Boolean
   Dim lngResult As Integer
   Dim lngThisKey As Long
   lngResult = RegCreateKey(HKEY_LOCAL_MACHINE, OUR_SUBKEY_PATH & strKeyName, lngThisKey)
   If lngResult = ERROR_SUCCESS Then
      lngResult = RegSetValue(lngThisKey, vbNullString, REG_SZ, strValue, Len(strValue))
   Else
      Err.Raise vbObjectError + 101, "WCCCars: SetKeyValue", "Cannot set Registry value"
   End If
   lngResult = RegCloseKey(lngThisKey)
End Function

