
Dim tDate,TodayorNot,password,city,ips,citys,i,j,a,WshShell,FileSystem,path
'True False
'���ã��Է��Ժ�����������ڵı����ļ�
TodayorNot = True
If TodayorNot Then
	tDate = Replace(Replace(Date,"/",""),"-","") 'ϵͳ��һ���������ռ������/ -��Ҳ��һ��
Else
	tDate = InputBox("�������ڸ�ʽ����20181121")
End if

password=InputBox("�������룺")
city=InputBox("�����������ĸ�����������gz")

If Len(password) <> 0 Or Len(city) <> 0 then
	Set WshShell = WScript.CreateObject("WScript.Shell") 
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	ips="10.254.65.193,10.254.65.197,10.254.89.193,10.254.89.209,10.254.89.201,10.254.89.217,10.254.67.193,10.254.65.209,10.254.65.213,10.254.65.217,10.254.67.209,10.254.67.177,10.254.65.177,10.254.65.181,10.254.67.181,10.254.65.185,10.254.67.213,10.254.67.217,10.219.0.5,10.254.94.193,10.254.95.193,10.254.95.209,10.254.94.209,10.254.94.201,10.254.95.201,10.254.94.177,10.254.94.180,10.254.95.177,10.254.95.180,10.254.94.161,10.254.94.164,10.254.95.161,10.254.95.164,10.254.94.183,10.254.94.145,10.254.95.183,10.254.95.145,10.254.95.219,10.254.68.162,10.254.68.165,10.254.68.168,10.254.68.171,10.254.69.162,10.254.69.165,10.254.68.174,10.254.68.177,10.254.69.168,10.254.69.171,10.254.69.173,10.254.70.209,10.254.70.193,10.254.70.201,10.254.70.215,10.254.70.169,10.254.70.218,10.254.70.172,10.254.70.161,10.254.70.165,10.254.74.193,10.254.74.209,10.254.74.161,10.254.74.164,10.254.74.167,10.254.74.203,10.254.74.220,10.254.71.161,10.254.71.164,10.254.71.167,10.254.71.170,10.254.71.173,10.254.71.195,10.254.78.119,10.254.78.177,10.254.78.180,10.254.76.119,10.254.76.177,10.254.84.193,10.254.84.201,10.254.75.249|10.254.65.194,10.254.65.198,10.254.89.194,10.254.89.210,10.254.89.202,10.254.89.218,10.254.67.194,10.254.65.210,10.254.65.214,10.254.65.218,10.254.67.210,10.254.67.178,10.254.65.178,10.254.65.182,10.254.67.182,10.254.65.186,10.254.67.214,10.254.67.218,10.219.0.6,10.254.94.194,10.254.95.194,10.254.95.210,10.254.94.210,10.254.94.202,10.254.95.202,10.254.94.178,10.254.94.181,10.254.95.178,10.254.95.181,10.254.94.162,10.254.94.165,10.254.95.162,10.254.95.165,10.254.94.184,10.254.94.146,10.254.95.184,10.254.95.146,10.254.95.220,10.254.68.161,10.254.68.164,10.254.68.167,10.254.68.170,10.254.69.161,10.254.69.164,10.254.68.173,10.254.68.176,10.254.69.167,10.254.69.170,10.254.69.174,10.254.70.210,10.254.70.194,10.254.70.202,10.254.70.216,10.254.70.170,10.254.70.219,10.254.70.173,10.254.70.162,10.254.70.166,10.254.74.194,10.254.74.210,10.254.74.162,10.254.74.165,10.254.74.168,10.254.74.204,10.254.74.221,10.254.71.162,10.254.71.165,10.254.71.168,10.254.71.171,10.254.71.174,10.254.71.196,10.254.78.120,10.254.78.178,10.254.78.181,10.254.76.120,10.254.76.178,10.254.84.194,10.254.84.202,10.254.75.250"
	citys = "gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,dg,dg,dg,dg,dg,dg,dg,dg,dg,dg,dg,fs,fs,fs,fs,fs,fs,fs,fs,fs,zs,zs,zs,zs,zs,zs,zs,hz,hz,hz,hz,hz,hz,zj,zj,zj,mm,mm,yj,yj,zh,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,gz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,sz,dg,dg,dg,dg,dg,dg,dg,dg,dg,dg,dg,fs,fs,fs,fs,fs,fs,fs,fs,fs,zs,zs,zs,zs,zs,zs,zs,hz,hz,hz,hz,hz,hz,zj,zj,zj,mm,mm,yj,yj,zh"
	
	Dim objDict,ip,backupPath
	Set objDict = WSH.CreateObject("Scripting.Dictionary")
	ipsArr = Split(ips,"|")
	backupPath=Array("���ñ���","���ñ���")

	j = 0
	citysArr = Split(citys,",")
	For i = 0 To 1
		For Each a In Split(ipsArr(i),",")
			objDict.Add a, citysArr(j)
			j = j + 1
		Next		
	Next
	
	path="D:\license���ݼ��"
	If Not FileSystem.FolderExists(path) Then 
		FileSystem.createfolder(path)
	End if
	
	j = 0
	For i = 0 To 1
		path="D:\license���ݼ��\" & backupPath(i) & tDate
		download = "D:\download.txt"
		If Not FileSystem.FolderExists(path) Then 
			FileSystem.createfolder(path) '����Ŀ¼
		End if
		For Each a In Split(ipsArr(i),",")
			If objDict(a)=city Then	
				Set OutPutFile = FileSystem.OpenTextFile(download,2,True) '2��ʾֱ�Ӹ���ԭ�ļ�
				OutPutFile.WriteLine "open " & a 
				OutPutFile.WriteLine "user backup " & password
				OutPutFile.WriteLine "lcd " & path
				OutPutFile.WriteLine "binary"
				OutPutFile.WriteLine "prompt"
				OutPutFile.WriteLine "mget " & tDate & "*.dat" 
				OutPutFile.WriteLine "bye" 
				OutPutFile.Close
				Wshshell.run "ftp -n -s:" & download , 1 , True '����رմ������ˣ�����True���������������ˡ�
			End if
			j = j + 1
		Next
		Wshshell.Run path			
	Next
	
	Dim sfile
	on error resume next
	set sfile=FileSystem.getfile(download)
	sfile.attributes=0
	sfile.Delete
Else
	MsgBox "����Ϊ��"
End if