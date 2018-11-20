Dim fso,ws,desktopDir
set fso  = createobject("scripting.filesystemobject")
Set ws=CreateObject("WScript.Shell")
desktopDir=ws.SpecialFolders(4)&"\license备份文件" & Replace(Date,"/","")
fso.createfolder(desktopDir)
