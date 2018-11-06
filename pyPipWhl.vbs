dim s
yn = msgbox("Yes-py2.7 | No-py3.7",vbYesNo,"请选择")
if yn =vbYes then
	pip = "py -2.7 -m pip install "
else
	pip = "py -3.7 -m pip install "
End if

whl = inputbox("请输入您的数据","输入数据","")
If whl <> "" then	
	pipWhl = "cmd.exe /k " & pip & whl
	rem pipWhl = "cmd.exe /k echo hello"
	rem msgbox pipWhl
	dim WSHshellA
	set WSHshellA = wscript.createobject("wscript.shell")
	WSHshellA.run pipWhl
End if
