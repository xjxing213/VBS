dim arges,i,sum,s1,s2

set arges = Wscript.Arguments
s1 = arges.Item(0)
s2 = arges.Item(1)
for i = s1 to s2
	sum = sum + i
next

Wscript.Echo "2λ֮�����Ϊ��" & sum

'cscript ���ļ�·�� 1,100
'C:\Users\admin>cscript C:\Users\admin\Desktop\VBS\demo4.vbs 1 100
'Microsoft (R) Windows Script Host Version 5.8
'��Ȩ����(C) Microsoft Corporation 1996-2001����������Ȩ����

'2λ֮�����Ϊ��5050

' *****/nologo
' C:\Users\admin>cscript /nologo C:\Users\admin\Desktop\VBS\demo4.vbs 1 100
' 2λ֮�����Ϊ��5050

' *****/e:vbscript���ı�Ҳ���Ե����ű�ִ�У���vbscript����ִ��
' C:\Users\admin>cscript /nologo /e:vbscript C:\Users\admin\Desktop\VBS\demo4.txt 1 100
' 2λ֮�����Ϊ��5050






