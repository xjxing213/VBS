
	'遍历一个文件夹下的所有文件
	Dim fso, ts, s,ret,arr,thispath
	Set fso = CreateObject("Scripting.FileSystemObject")  
	thispath = fso.GetFolder(".").Path '当前脚本目录
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
	Dim regEx, Match, Matches   ' 建立变量。
	Set regEx = New RegExp   ' 建立正则表达式。
	regEx.Pattern = patrn   ' 设置模式。
	regEx.IgnoreCase = True   ' 设置是否区分字符大小写。
	regEx.Global = True   ' 设置全局可用性。
	Set Matches = regEx.Execute(strng)   ' 执行搜索。
	RegExpTest=Matches(0)
'	For Each Match In Matches   ' 遍历匹配集合。
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
	Dim a, d, i,esn1,esn2,bsc   ' 创建一些变量。
	Set d = CreateObject("Scripting.Dictionary")
	esn1=Array()
	esn2=Array()
	bsc=Array("广州_BSC01","广州_BSC02","广州_BSC03","广州_BSC04","广州_BSC05","广州_BSC06","广州_BSC07","广州_BSC08","广州_BSC09","广州_BSC10","广州_BSC11","广州_BSC12","广州_BSC13","广州_BSC14","广州_BSC15","广州_BSC16","广州_BSC17","广州_BSC18","广州_BSC20","深圳_BSC01","深圳_BSC02","深圳_BSC03","深圳_BSC04","深圳_BSC05","深圳_BSC06","深圳_BSC07","深圳_BSC08","深圳_BSC09","深圳_BSC10","深圳_BSC11","深圳_BSC12","深圳_BSC13","深圳_BSC14","深圳_BSC15","深圳_BSC16","深圳_BSC17","深圳_BSC18","深圳_BSC19","东莞_BSC01","东莞_BSC02","东莞_BSC03","东莞_BSC04","东莞_BSC05","东莞_BSC06","东莞_BSC07","东莞_BSC08","东莞_BSC09","东莞_BSC10","东莞_BSC11","佛山_BSC01","佛山_BSC02","佛山_BSC03","佛山_BSC04","佛山_BSC05","佛山_BSC06","佛山_BSC07","佛山_BSC08","佛山_BSC09","中山_BSC01","中山_BSC02","中山_BSC03","中山_BSC04","中山_BSC05","中山_BSC06","中山_BSC07","惠州_BSC01","惠州_BSC02","惠州_BSC03","惠州_BSC04","惠州_BSC05","惠州_BSC06","湛江_BSC01","湛江_BSC02","湛江_BSC03","茂名_BSC01","茂名_BSC02","阳江_BSC01","阳江_BSC02","珠海_BSC01")
	i=0
	For i=0 To UBound(esn1)
		d.Add esn1(i), bsc(i)
		d.Add esn2(i), bsc(i)
	Next
	thispath = fso.GetFolder(".").Path '当前脚本目录
	bscpath=d(ret)
	bscpath=thispath & "\" & bscpath '当前脚本目录\地市_BSC
    fso.CreateFolder(bscpath)'创建当前脚本目录\地市_BSC
	fso.MoveFile fpath, bscpath & "\cbsclicense.dat"'移动并重命名license
	d.RemoveAll
End Sub

