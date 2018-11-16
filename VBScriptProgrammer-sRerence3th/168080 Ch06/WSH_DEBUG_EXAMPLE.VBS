Option Explicit

'This script uses the scripting runtime FileSystemObject to copy a file
'called WSH_DEBUG_TEST_FILE.TXT that is assumed to be in the same directory
'as this script.

Const FILE_NAME = "WSH_DEBUG_TEST_FILE.TXT"
Const COPY_SUFFIX = "_COPY"
Const OVERWRITE_FILE = True

'***** Main Code

Dim objFSO
Dim strExtension
Dim blnFileExists
Dim strNewFileName
Dim strScriptPath

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = GetScriptPath()
blnFileExists = VerifyFile(strScriptPath, FILE_NAME)

If blnFileExists Then
    strExtension = GetExtension(FILE_NAME)
    strNewFileName = MakeNewFileName(FILE_NAME, _
        strExtension, COPY_SUFFIX)
    CopyFile strScriptPath & FILE_NAME, _
        strScriptPath & strNewFileName, _
        OVERWRITE_FILE
Else
    On Error GoTo 0
    Err.Raise vbObjectError + 10000, _
        "WSH_DEBUG_EXAMPLE.VBS", _
        "Expected file " & FILE_NAME & " not found."
End If

'***** Supporting procedures and functions

Private Sub CopyFile(strFileName, strNewFileName, blnOverwrite)
    objFSO.CopyFile strFileName, strNewFileName, blnOverwrite
End Sub

Private Function GetExtension(strFileName)
    GetExtension = objFSO.GetExtensionName(strFileName)
End Function

Private Function GetScriptPath
    Dim strPath

    strPath = objFSO.GetAbsolutePathName(WScript.ScriptFullName)
    strPath = Left(strPath, _
        Len(strPath) - Len(objFSO.GetFileName(strPath)))
    GetScriptPath = strPath
End Function

Private Function VerifyFile(strPath, strFileName)
    VerifyFile = objFSO.FileExists(strPath & strFileName)
End Function

Private Function MakeNewFileName(strFileName, strExtension, strSuffix)
    MakeNewFileName = Left(strFileName, Len(strFileName) _
        - (1 + Len(strExtension))) & strSuffix & _
        "." & strExtension
End Function