
On Error Resume Next
Const CONVERSION_FACTOR=1048576
Const WARNING_THRESHOLD =1000
If Wscript.Arguments.Count =0 Then
	Wscript.Echo "Usage: test.vbs serverl [server2] [server3] ..."
	Wscript.Quit
End If
For Each Computer In Wscript.Arguments
	Set objWMIService =GetObject("winmgmts://" & Computer)
	If Err.Number <> 0 Then
		Wscript.Echo Computer & " " & Err.Description
		Err.Clear
	Else
		Set colLogicalDisk = objWMIService.InstancesOf("Win32_Logicaldisk")
		For Each objLogicalDisk In collogicaldisk
			FreeMegaBytes = objLogicalDisk.Freespace / CONVERSION_FACTOR
			Wscript.Echo Computer & " " &objLogicalDisk.DeviceID & _
				" ʣ����̿ռ�Ϊ��M����" & FreeMegaBytes
			If FreeMegaBytes< WARNING_THRESHOLD Then
				Wscript.Echo Computer & " " &objLogicalDisk.DeviceID & _
				"is low on disk space"
			End If
		Next
	End If
Next


'C:\Users\Administrator\Desktop\22-VBS>cscript err������.vbs
'Microsoft (R) Windows Script Host Version 5.8
'��Ȩ����(C) Microsoft Corporation 1996-2001����������Ȩ����

'Usage: test.vbs serverl [server2] [server3] ...

'C:\Users\Administrator\Desktop\22-VBS>cscript err������.vbs "PC-20180"
'Microsoft (R) Windows Script Host Version 5.8
'��Ȩ����(C) Microsoft Corporation 1996-2001����������Ȩ����

'PC-20180 Զ�̷����������ڻ򲻿���

'C:\Users\Administrator\Desktop\22-VBS>cscript err������.vbs "PC-20180403XXXX"
'Microsoft (R) Windows Script Host Version 5.8
'��Ȩ����(C) Microsoft Corporation 1996-2001����������Ȩ����

'PC-20180403XXXX C: ʣ����̿ռ�Ϊ��M����38210.5
'PC-20180403XXXX D: ʣ����̿ռ�Ϊ��M����100612.05078125
'PC-20180403XXXX E: ʣ����̿ռ�Ϊ��M����108320.75390625
'PC-20180403XXXX F: ʣ����̿ռ�Ϊ��M����1144.8125

'C:\Users\Administrator\Desktop\22-VBS>cscript err������.vbs "PC-20180403XXXX"
'Microsoft (R) Windows Script Host Version 5.8
'��Ȩ����(C) Microsoft Corporation 1996-2001����������Ȩ����

'PC-20180403XXXX C: ʣ����̿ռ�Ϊ��M����38207.734375
'PC-20180403XXXX D: ʣ����̿ռ�Ϊ��M����100612.05078125
'PC-20180403XXXX E: ʣ����̿ռ�Ϊ��M����108320.75390625
'PC-20180403XXXX F: ʣ����̿ռ�Ϊ��M����826.171875
'PC-20180403XXXX F:is low on disk space
