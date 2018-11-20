20160426-DG-BSC10-cbsclicense.dat
Dim WshShell, curDir, wShell, file,desktopDate,a,ips,arr
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
tDate = Replace(Replace(Date,"/",""),"-","") '系统不一样，年月日间隔符（/ -）也不一样
desktopDate=WshShell.SpecialFolders(4)&"\license备份文件" & tDate
downloadPath = desktopDate & "\download.txt"
'判断目录是否存在
If Not FileSystem.FolderExists(desktopDate) Then 
	FileSystem.createfolder(desktopDate) '创建目录
End if

ips = "10.254.69.171"
arr = Split(ips,",")
For Each a In arr
	Set OutPutFile = FileSystem.OpenTextFile(downloadPath,2,True) '2表示直接覆盖原文件
	OutPutFile.WriteLine "open " & a 
	OutPutFile.WriteLine "user "user" "password""
	OutPutFile.WriteLine "lcd " & desktopDate
	OutPutFile.WriteLine "binary"
	OutPutFile.WriteLine "prompt"
	OutPutFile.WriteLine "mget " & tDate & "*.dat" 
	OutPutFile.WriteLine "bye" 
	OutPutFile.Close
	Wshshell.run "ftp -n -s:" & downloadPath
Next
'Wshshell.run "ftp -n -s: ""C:\Documents and Settings\Administrator\桌面\license备份文件20181120\download.txt"""
Dim sfile
on error resume next
set sfile=FileSystem.getfile(downloadPath)
sfile.attributes=0
sfile.Delete
