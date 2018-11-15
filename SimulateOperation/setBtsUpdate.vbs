Dim str,parr,arr,a,wsh,p1,p2,p，mg
'输入基站名称|搜索|勾选升级或取消升级|确定|确定
p1=Array(Array(746,139),Array(1239,183),Array(301,236),Array(1270,186),Array(955,817),Array(796,681)) '升级
p2=Array(Array(746,139),Array(1147,185),Array(301,236),Array(1270,186),Array(906,159),Array(796,679)) '取消升级
p=p2
set wsh = createobject("Wscript.Shell")
Set mouse=New SetMouse
'True False
mg=True
If mg Then
	Call main()
Else
	Call getPosxy()
End If

Sub getPosxy()
mouse.getpos x,y ''获得鼠标当前位置坐标
WScript.Echo x & "," & y
End Sub

Sub main()
Dim lastBts
str=Inputbox("请输入您的数据")
arr = Split(str, ",")
lastBts=20
For Each a In arr
	'click bts
	Call cz(p(0)(0),p(0)(1)) '1
	'Ctrl+A
'	wsh.sendkeys "^(A)" '全选无效，为啥？
'	WScript.Sleep 300
	'lastBts根据上一个基站的长度，合理调整删除次数
	wsh.sendkeys "{DEL " & lastBts &"}" 
	lastBts=Len(a)
	WScript.Sleep 300
	'input btsname,'写入信息到剪切板
	wsh.Run "mshta vbscript:ClipboardData.SetData("&chr(34)&"text"&chr(34)&"," &Chr(34)& a &Chr(34)& ")(close)",0,True
	'Ctrl+V
	wsh.sendkeys "^(V)"
	'查询
	Call cz(p(1)(0),p(1)(1))'2
	WScript.Sleep 300
	'勾选
	Call cz(p(2)(0),p(2)(1))'3
	'升级
	Call cz(p(3)(0),p(3)(1))'4
	'确定
	Call cz(p(4)(0),p(4)(1))'5
	'成功确定
	Call cz(p(5)(0),p(5)(1))'6
Next
'	把最后那个站清除掉
	Call cz(p(0)(0),p(0)(1)) '1
	wsh.sendkeys "{DEL " & lastBts &"}" 

End Sub

Sub cz(x, y)	
	mouse.move x,y '把鼠标移动到坐标
	WScript.Sleep 100
	mouse.clik "LEFT" '左击
	WScript.Sleep 500
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
		'vbs 完全控制excel
		reg_key = "HKEY_CURRENT_USER\Software\Microsoft\Office\$\Excel\Security\AccessVBOM"
		reg_key = Replace(reg_key, "$", xls.Version)
		S.RegWrite reg_key, 1, "REG_DWORD"
		'model 代码
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
	
	'关闭
	Private Sub Class_Terminate
		xls.DisplayAlerts = False
		wbk.Close
		xls.Quit
	End Sub
	
	'可调用过程
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
