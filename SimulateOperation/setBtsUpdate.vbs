Dim str,arr,a,wsh

set wsh = createobject("Wscript.Shell")
Set mouse=New SetMouse
Call main()
'Call getPosxy()

Sub getPosxy()
mouse.getpos x,y ''�����굱ǰλ������
WScript.Echo x & "," & y
End Sub

Sub main()
Dim lastBts
str=Inputbox("��������������")
arr = Split(str, ",")
lastBts=20
For Each a In arr
	'click bts
	Call cz(746,139)
	'Ctrl+A
'	wsh.sendkeys "^(A)" 'ȫѡ��Ч��Ϊɶ��
'	WScript.Sleep 300
	'lastBts������һ����վ�ĳ��ȣ��������ɾ������
	wsh.sendkeys "{DEL " & lastBts &"}" 
	lastBts=Len(a)
	WScript.Sleep 300
	'input btsname,'д����Ϣ�����а�
	wsh.Run "mshta vbscript:ClipboardData.SetData("&chr(34)&"text"&chr(34)&"," &Chr(34)& a &Chr(34)& ")(close)",0,True
	'Ctrl+V
	wsh.sendkeys "^(V)"
	'��ѯ
	Call cz(1239,183)
	WScript.Sleep 300
	'��ѡ
	Call cz(301,236)
	'����
	Call cz(1268,181)
	'ȷ��
	Call cz(955,817)
	'�ɹ�ȷ��
	Call cz(796,681)
Next
End sub

Sub cz(x, y)	
	mouse.move x,y '������ƶ�������
	WScript.Sleep 100
	mouse.clik "LEFT" '���
	WScript.Sleep 400
	'Case "LEFT"
	'Case "RIGHT"
	'Case "MIDDLE"
	'Case "DBCLICK"
End Sub

Class SetMouse
	Private S
	Private xls, wbk, module1
	Private reg_key, xls_code, x, y
	
	Private Sub Class_Initialize()
		Set xls = CreateObject("Excel.Application") 
		Set S = CreateObject("wscript.Shell")
		'vbs ��ȫ����excel
		reg_key = "HKEY_CURRENT_USER\Software\Microsoft\Office\$\Excel\Security\AccessVBOM"
		reg_key = Replace(reg_key, "$", xls.Version)
		S.RegWrite reg_key, 1, "REG_DWORD"
		'model ����
		xls_code = _
		"Private Type POINTAPI : X As Long : Y As Long : End Type" & vbCrLf & _
		"Private Declare Function SetCursorPos Lib ""user32"" (ByVal x As Long, ByVal y As Long) As Long" & vbCrLf & _
		"Private Declare Function GetCursorPos Lib ""user32"" (lpPoint As POINTAPI) As Long" & vbCrLf & _
		"Private Declare Sub mouse_event Lib ""user32"" Alias ""mouse_event"" " _
		& "(ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)" & vbCrLf & _
		"Public Function getx() As Long" & vbCrLf & _
		"Dim pt As POINTAPI : GetCursorPos pt : getx = pt.X" & vbCrLf & _
		"End Function" & vbCrLf & _
		"Public Function gety() As Long" & vbCrLf & _
		"Dim pt As POINTAPI: GetCursorPos pt : gety = pt.Y" & vbCrLf & _
		"End Function"
		Set wbk = xls.Workbooks.Add 
		Set module1 = wbk.VBProject.VBComponents.Add(1)
		module1.CodeModule.AddFromString xls_code 
	End Sub
	
	'�ر�
	Private Sub Class_Terminate
		xls.DisplayAlerts = False
		wbk.Close
		xls.Quit
	End Sub
	
	'�ɵ��ù���
	Public Sub getpos( x, y) 
		x = xls.Run("getx") 
		y = xls.Run("gety") 
	End Sub
	
	Public Sub move(x,y)
		xls.Run "SetCursorPos", x, y
	End Sub
	
	Public Sub clik(keydown)
		Select Case UCase(keydown)
			Case "LEFT"
			xls.Run "mouse_event", &H2 + &H4, 0, 0, 0, 0
			Case "RIGHT"
			xls.Run "mouse_event", &H8 + &H10, 0, 0, 0, 0
			Case "MIDDLE"
			xls.Run "mouse_event", &H20 + &H40, 0, 0, 0, 0
			Case "DBCLICK"
			xls.Run "mouse_event", &H2 + &H4, 0, 0, 0, 0
			xls.Run "mouse_event", &H2 + &H4, 0, 0, 0, 0
		End Select
	End Sub
	
End Class
