
	'����һ���ļ����µ������ļ�
	Dim fso, ts, s,ret,arr,thispath
	Set fso = CreateObject("Scripting.FileSystemObject")  
	thispath = fso.GetFolder(".").Path '��ǰ�ű�Ŀ¼
	Set oFolder = fso.GetFolder(thispath)  
	Set oFiles = oFolder.Files  
	For Each oFile In oFiles
		If Right(oFile.Path,3)="dat" Then
			ReadFiles(oFile.Path)
		End If
	Next 	
	Set oFolder = Nothing  
	Set oSubFolders = Nothing  
	Set oFso = Nothing  


Function RegExpTest(patrn, strng)
	Dim regEx, Match, Matches   ' ����������
	Set regEx = New RegExp   ' ����������ʽ��
	regEx.Pattern = patrn   ' ����ģʽ��
	regEx.IgnoreCase = True   ' �����Ƿ������ַ���Сд��
	regEx.Global = True   ' ����ȫ�ֿ����ԡ�
	Set Matches = regEx.Execute(strng)   ' ִ��������
	RegExpTest=Matches(0)
'	For Each Match In Matches   ' ����ƥ�伯�ϡ�
'		RetStr = Match.Value
'	Next
'	RegExpTest = RetStr
End Function

Sub ReadFiles(fpath)
	Dim fso, ts, s,ret,arr,bscpath,f,thispath
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(fpath)
	s = ts.ReadAll
	ret = RegExpTest("(210305)[0-9A-Z]{14}",s)
	ts.Close
	Dim a, d, i,esn1,esn2,bsc   ' ����һЩ������
	Set d = CreateObject("Scripting.Dictionary")
	esn1=Array()
	esn2=Array()
	bsc=Array("����_BSC01","����_BSC02","����_BSC03","����_BSC04","����_BSC05","����_BSC06","����_BSC07","����_BSC08","����_BSC09","����_BSC10","����_BSC11","����_BSC12","����_BSC13","����_BSC14","����_BSC15","����_BSC16","����_BSC17","����_BSC18","����_BSC20","����_BSC01","����_BSC02","����_BSC03","����_BSC04","����_BSC05","����_BSC06","����_BSC07","����_BSC08","����_BSC09","����_BSC10","����_BSC11","����_BSC12","����_BSC13","����_BSC14","����_BSC15","����_BSC16","����_BSC17","����_BSC18","����_BSC19","��ݸ_BSC01","��ݸ_BSC02","��ݸ_BSC03","��ݸ_BSC04","��ݸ_BSC05","��ݸ_BSC06","��ݸ_BSC07","��ݸ_BSC08","��ݸ_BSC09","��ݸ_BSC10","��ݸ_BSC11","��ɽ_BSC01","��ɽ_BSC02","��ɽ_BSC03","��ɽ_BSC04","��ɽ_BSC05","��ɽ_BSC06","��ɽ_BSC07","��ɽ_BSC08","��ɽ_BSC09","��ɽ_BSC01","��ɽ_BSC02","��ɽ_BSC03","��ɽ_BSC04","��ɽ_BSC05","��ɽ_BSC06","��ɽ_BSC07","����_BSC01","����_BSC02","����_BSC03","����_BSC04","����_BSC05","����_BSC06","տ��_BSC01","տ��_BSC02","տ��_BSC03","ï��_BSC01","ï��_BSC02","����_BSC01","����_BSC02","�麣_BSC01")
	i=0
	For i=0 To UBound(esn1)
		d.Add esn1(i), bsc(i)
		d.Add esn2(i), bsc(i)
	Next
	thispath = fso.GetFolder(".").Path '��ǰ�ű�Ŀ¼
	bscpath=d(ret)
	bscpath=thispath & "\" & bscpath '��ǰ�ű�Ŀ¼\����_BSC
    fso.CreateFolder(bscpath)'������ǰ�ű�Ŀ¼\����_BSC
	fso.MoveFile fpath, bscpath & "\cbsclicense.dat"'�ƶ���������license
	d.RemoveAll
End Sub

