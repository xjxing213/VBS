dim s
yn = msgbox("Yes2.7|No3.7",vbYesNo,"��ѡ��")
if yn =vbYes then
	pip = "py -2.7 -m pip install "
else
	pip = "py -3.7 -m pip install "
End if

whl = inputbox("��������������","��������","")
If whl <> "" then	
	pipWhl = "cmd.exe /k " & pip & whl
	rem pipWhl = "cmd.exe /k echo hello"
	rem msgbox pipWhl
	dim WSHshellA
	set WSHshellA = wscript.createobject("wscript.shell")
	WSHshellA.run pipWhl
End if