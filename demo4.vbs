dim arges,i,sum,s1,s2

set arges = Wscript.Arguments
s1 = arges.Item(0)
s2 = arges.Item(1)
for i = s1 to s2
	sum = sum + i
next

Wscript.Echo "2位之间求和为：" & sum

'cscript 本文件路径 1,100
'C:\Users\admin>cscript C:\Users\admin\Desktop\VBS\demo4.vbs 1 100
'Microsoft (R) Windows Script Host Version 5.8
'版权所有(C) Microsoft Corporation 1996-2001。保留所有权利。

'2位之间求和为：5050

' *****/nologo
' C:\Users\admin>cscript /nologo C:\Users\admin\Desktop\VBS\demo4.vbs 1 100
' 2位之间求和为：5050

' *****/e:vbscript：文本也可以当作脚本执行，用vbscript引擎执行
' C:\Users\admin>cscript /nologo /e:vbscript C:\Users\admin\Desktop\VBS\demo4.txt 1 100
' 2位之间求和为：5050






