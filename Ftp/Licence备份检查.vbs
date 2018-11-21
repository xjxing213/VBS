
'20160426-DG-BSC10-cbsclicense.dat

Dim ask,fdate,tDate,password
ask = MsgBox("检查当天（是）|其他日期（否）",vbYesNo)
If ask = vbyes Then
	tDate = Replace(Replace(Date,"/",""),"-","") '系统不一样，年月日间隔符（/ -）也不一样
Else
	tDate = InputBox("输入日期格式：如20181121")
End if

password=InputBox("输入密码：")
If Len(password) <> 0 then
	Dim WshShell,FileSystem,path,a,ips,arr
	Set WshShell = WScript.CreateObject("WScript.Shell") 
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	'WshShell.SpecialFolders(4) '获取桌面路径，备用
	'---------------------主用备份----------------------------
	path="D:\license主用备份" & tDate
	download = path & "\download.txt"
	If Not FileSystem.FolderExists(path) Then 
		FileSystem.createfolder(path) '创建目录
	End if
	
	ips = "10.254.65.195,10.254.65.199,10.254.89.195,10.254.89.211,10.254.89.203,10.254.89.219,10.254.67.195,10.254.65.211,10.254.65.215,10.254.65.219,10.254.67.211,10.254.67.179,10.254.65.179,10.254.65.183,10.254.67.183,10.254.65.187,10.254.67.215,10.254.67.219,10.219.0.7,10.254.94.195,10.254.95.195,10.254.95.211,10.254.94.211,10.254.94.203,10.254.95.203,10.254.94.179,10.254.94.182,10.254.95.179,10.254.95.182,10.254.94.163,10.254.94.166,10.254.95.163,10.254.95.166,10.254.94.185,10.254.94.147,10.254.95.185,10.254.95.147,10.254.95.221,10.254.68.163,10.254.68.166,10.254.68.169,10.254.68.172,10.254.69.163,10.254.69.166,10.254.68.175,10.254.68.178,10.254.69.169,10.254.69.172,10.254.69.175,10.254.70.211,10.254.70.195,10.254.70.203,10.254.70.217,10.254.70.171,10.254.70.220,10.254.70.174,10.254.70.163,10.254.70.167,10.254.74.195,10.254.74.211,10.254.74.163,10.254.74.166,10.254.74.169,10.254.74.205,10.254.74.222,10.254.71.163,10.254.71.166,10.254.71.169,10.254.71.172,10.254.71.175,10.254.71.197,10.254.78.118,10.254.78.179,10.254.78.182,10.254.76.118,10.254.76.179,10.254.84.195,10.254.84.203,10.254.75.251"
	'ips = "10.254.65.195"
	arr = Split(ips,",")
	For Each a In arr
		Set OutPutFile = FileSystem.OpenTextFile(download,2,True) '2表示直接覆盖原文件
		OutPutFile.WriteLine "open " & a 
		OutPutFile.WriteLine "user backup " & password
		OutPutFile.WriteLine "lcd " & path
		OutPutFile.WriteLine "binary"
		OutPutFile.WriteLine "prompt"
		OutPutFile.WriteLine "mget " & tDate & "*.dat" 
		OutPutFile.WriteLine "bye" 
		OutPutFile.Close
		Wshshell.run "ftp -n -s:" & download , 1 , True '逐个关闭窗口神了，加了True都可以下载下来了。
	Next
	
	Dim sfile
	on error resume next
	set sfile=FileSystem.getfile(download)
	sfile.attributes=0
	sfile.Delete
	'打开结果目录
	Wshshell.Run path
	
	'---------------------备用备份----------------------------
	path="D:\license备用备份" & tDate
	download = path & "\download.txt"
	If Not FileSystem.FolderExists(path) Then 
		FileSystem.createfolder(path) '创建目录
	End if
	
	ips = "10.254.65.194,10.254.65.198,10.254.89.194,10.254.89.210,10.254.89.202,10.254.89.218,10.254.67.194,10.254.65.210,10.254.65.214,10.254.65.218,10.254.67.210,10.254.67.178,10.254.65.178,10.254.65.182,10.254.67.182,10.254.65.186,10.254.67.214,10.254.67.218,10.254.89.92,10.219.0.6,10.254.94.194,10.254.95.194,10.254.95.210,10.254.94.210,10.254.94.202,10.254.95.202,10.254.94.178,10.254.94.181,10.254.95.178,10.254.95.181,10.254.94.162,10.254.94.165,10.254.95.162,10.254.95.165,10.254.94.184,10.254.94.146,10.254.95.184,10.254.95.146,10.254.95.220,10.254.68.161,10.254.68.164,10.254.68.167,10.254.68.170,10.254.69.161,10.254.69.164,10.254.68.173,10.254.68.176,10.254.69.167,10.254.69.170,10.254.69.174,10.254.70.210,10.254.70.194,10.254.70.202,10.254.70.216,10.254.70.170,10.254.70.219,10.254.70.173,10.254.70.162,10.254.70.166,10.254.74.194,10.254.74.210,10.254.74.162,10.254.74.165,10.254.74.168,10.254.74.204,10.254.74.221,10.254.71.162,10.254.71.165,10.254.71.168,10.254.71.171,10.254.71.174,10.254.71.196,10.254.78.120,10.254.78.178,10.254.78.181,10.254.76.120,10.254.76.178,10.254.84.194,10.254.84.202,10.254.75.250"
	'ips = "10.254.65.195"
	arr = Split(ips,",")
	For Each a In arr
		Set OutPutFile = FileSystem.OpenTextFile(download,2,True) '2表示直接覆盖原文件
		OutPutFile.WriteLine "open " & a 
		OutPutFile.WriteLine "user backup " & password
		OutPutFile.WriteLine "lcd " & path
		OutPutFile.WriteLine "binary"
		OutPutFile.WriteLine "prompt"
		OutPutFile.WriteLine "mget " & tDate & "*.dat" 
		OutPutFile.WriteLine "bye" 
		OutPutFile.Close
		Wshshell.run "ftp -n -s:" & download , 1 , True '逐个关闭窗口神了，加了True都可以下载下来了。
	Next
	
	on error resume next
	set sfile=FileSystem.getfile(download)
	sfile.attributes=0
	sfile.Delete
	'打开结果目录
	Wshshell.Run path
Else
	MsgBox "密码输入为空"
End if