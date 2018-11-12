Set objDialog = CreateObject("UserAccounts.CommonDialog") 
objDialog.Filter = "VBScript Scripts|*.vbs|All Files|*.*" 
objDialog.Flags = &H0200 
objDialog.FilterIndex = 1 
objDialog.InitialDir = "C:/Scripts" 
intResult = objDialog.ShowOpen 
If intResult = 0 Then 
 Wscript.Quit 
Else 
 arrFiles = Split(objDialog.FileName, " ") 
 For i = 1 to Ubound(arrFiles) 
 strFile = arrFiles(0) & arrFiles(i) 
 Wscript.Echo strFile 
 Next 
End If