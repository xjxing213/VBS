'Dim re,s,sc

'Set re = new RegExp
're.pattern = "France"
'sc = "The rain in France falls mainly on the plains."
'MsgBox re.Replace(sc,"Spain")


'Dim re,s
'Set re = new RegExp
're.pattern = "\bin"
'sc = "The rain in Spain falls mainly on the plains."
'MsgBox re.Replace(sc,"in the country of")

'Dim re,s
'Set re = new RegExp
're.pattern = "in"
're.Global = True
'sc = "The rain in Spain falls mainly on the plains."
'MsgBox re.Replace(sc,"in the country of")
'#RegExp 对象
'------------------------------------
're属性：Global，IgnoreCase，Pattern
're方法：Execute，Replace，Test
'Global：匹配所有
'IgnoreCase：是否忽略大小写
'Pattern：需要搜索的正则字符串表达式
'------------------------------------
'Dim re,s
'Set re = new RegExp
're.pattern = "in"
're.Global = True
're.IgnoreCase = True	
'sc = "The rain In Spain falls mainly on the plains."
'WScript.Echo re.Replace(sc,"in the country of")

'#Result:
'#The rain the country of in the country of Spain the country of falls main the country ofly on the plain the country ofs.

'锚定和缩短模式
'Dim re, s
'Set re = New RegExp
're.Pattern = "\d{3}"
's = "Spain received 100 millimeters of rain in the last 2 weeks."
'MsgBox re.Replace(s, "a whopping number of")


'Dim re, s
'Set re = New RegExp
're.Pattern = "http://(\w+[\w-]*\w+\.)*\w+"
's = "http://www.kingsley-hughes.com is a valid web address. And so is "
's = s & vbCrLf & "http://www.wrox.com. And "
's = s & vbCrLf & "http://www.pc.ibm.com-even with 4 levels."
'MsgBox re.Replace(s, "** TOP SECRET! **")

'匹配字母
'Dim re, s
'Set re = New RegExp
're.Pattern = "(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)"
's = "VBScript is not very cool."
'MsgBox re.Replace(s, "$1 $2 $4 $5")
'#VBScript is very cool.

'Test的方法
'Dim re, s
'Set re = New RegExp
're.IgnoreCase = True
're.Pattern = "http://(\w+[\w-]*\w+\.)*\w+"
's = "Some long string with http://www.wrox.com buried in it."
'If re.Test(s) Then
'MsgBox "Found a URL."
'Else
'MsgBox "No URL found."
'End If

'Matches 集合
'Matches 集合含有正则表达式的 Match 对象。
'只有使用 RegExp 对象的 Execute 方法才能创建这个集合。要记住 Matches 集合的属性
'跟独立的 Match 对象的属性一样，都是只读的。
'关于 RegExp 对象详见“RegExp 对象”一节。
'当一个正则表达式被执行时，会产生 0 个或多个 Match 对象。每个 Match 对象提供以
'下三个内容：
'● 正则表达式所找到的字符串
'● 字符串的长度
'● 指向找到该匹配的位置的索引
'要记得将 Global 属性设为 True，否则您的 Matches 集合中最多也只会有一个成员。这
'种方法很简单，但是很难调试！
'Dim re, objMatch, colMatches, sMsg
'Set re = New RegExp
're.Global = True 
're.Pattern = "http://(\w+[\w-]*\w+\.)*\w+" 
's = "http://www.kingsley-hughes.com is a valid web address. And so is "
's = s & vbCrLf & "http://www.wrox.com. As is "
's = s & vbCrLf & "http://www.wiley.com."
'Set colMatches = re.Execute(s)
'sMsg = ""
'For Each objMatch in colMatches
'sMsg = sMsg & "Match of " & objMatch.Value
'sMsg = sMsg & ", found at position " & objMatch.FirstIndex & " of the string."
'sMsg = sMsg & "The length matched is "
'sMsg = sMsg & objMatch.Length & "." & vbCrLf
'Next
'MsgBox sMsg

'Matches 的属性
'Matches 是一个简单的集合，只有两个属性：
'1. Count 返回集合中的元素数量。
'Dim re, objMatch, colMatches, sMsg
'Set re = New RegExp
're.Global = True
're.Pattern = "http://(\w+[\w-]*\w+\.)*\w+"
's = "http://www.kingsley-hughes.com is a valid web address. And so is "
's = s & vbCrLf & "http://www.wrox.com. As is "
's = s & vbCrLf & "http://www.wiley.com."
'Set colMatches = re.Execute(s)
'MsgBox colMatches.Count

'2. Item 根据指定的键返回元素。
'Dim re, objMatch, colMatches, sMsg
'Set re = New RegExp
're.Global = True
're.Pattern = "http://(\w+[\w-]*\w+\.)*\w+"
's = "http://www.kingsley-hughes.com is a valid web address. And so is "
's = s & vbCrLf & "http://www.wrox.com. As is " 
's = s & vbCrLf _
'& "http://www.wiley.com."
'Set colMatches = re.Execute(s)
'MsgBox colMatches.item(0)
'MsgBox colMatches.item(1)
'MsgBox colMatches.item(2)

'分解 URI
'这个例子用来将统一资源定位器(Universal Resource Indicator， URI)分解成多个部分。
'比如下面这个 URI：
'www.wrox.com:80/misc-pages/support.shtml
'可以编写一个脚本将其分解成协议(比如 ftp、 http 等)、域名地址、页面和路径。可以用
'下面的模式实现这个功能：
'"(\w+):\ / \ /([^ / :]+)(:\d*)?( [^ # ]*)"
'下面的代码完成了这个工作：
'Dim re, s 
'Set re = New RegExp
're.Pattern = "(\w+):\/\/([^/:]+)(:\d*)?([^#]*)"
're.Global = True
're.IgnoreCase = True
's = "http://www.wrox.com:80/misc-pages/support.shtml"
'MsgBox re.Replace(s, "$1")
'MsgBox re.Replace(s, "$2")
'MsgBox re.Replace(s, "$3")
'MsgBox re.Replace(s, "$4")

'检查 HTML 元素
'检查 HTML 元素很简单，需要的只是一个正确的模式。下面这个就能检查元素开始和结
'束标签。
'"<(.*)>.*<\/\1>"
'如何编写脚本取决于您想要实现什么功能。这个简单的脚本只是一个演示。可以改进这
'段代码专门用于检查某种元素，或是实现基本的错误检查。
'Dim re, s
'Set re = New RegExp
're.IgnoreCase = True
're.Pattern = "<(.*)>.*<\/\1>"
's = "<a>This is a paragraph</a>"
'If re.Test(s) Then
'MsgBox "HTML element found."
'Else
'MsgBox "No HTML element found."
'End If

'匹配空白
'有时您可能真的需要匹配空白，也就是空行或是只有空白(空格和制表符)的行。下面的
'模式可以满足这个需求。
'"^[ \t]*$"
'说明如下：
'^—— 匹配每一行的开头。
'[ \t]*—— 匹配 0 个或多个空格或制表符(\t)。
'$—— 匹配行结尾。
'Dim re, s, colMatches, objMatch, sMsg
'Set re = New RegExp
're.Global = True 
're.Pattern = "^[ \t]*$"
's = " "
'Set colMatches = re.Execute(s)
'sMsg = ""
'For Each objMatch in colMatches
'	sMsg = sMsg & "Blank line found at position " & _
'	objMatch.FirstIndex & " of the string."
'Next
'MsgBox sMsg

'#匹配 HTML 注释标签
'当您学习稍后的第 15 章，“Windows 脚本宿主”时，您将学会如何使用 VBScript 和 Wind
'ows 脚本宿主操作文件系统，这样您就能读取和修改文件。正则表达式的一个应用就是查找
'HTML 文件中的注释标签。可以在将其发布到网络之前将其中的注释清除。
'这个脚本可以检查 HTML 注释标签。
'Dim re, s
'Set re = New RegExp
're.Global = True
're.Pattern = "^.*<!--.*-->.*$"
's = " <title>A Title</title> <!-- a title tag -->"
'If re.Test(s) Then
'MsgBox "HTML comment tags found."
'Else
'MsgBox "No HTML comment tags found."
'End If
'对该模式稍作修改，并使用 Replace 方法就能将脚本中的注释清除。
Dim re, s
Set re = New RegExp
re.Global = True
re.Pattern = "(^.*)(<!--.*-->)(.*$)"
s = " <title>A Title</title> <!-- a title tag -->"
If re.Test(s) Then
MsgBox "HTML comment tags found."
Else
MsgBox "No HTML comment tags found."
End If
MsgBox re.Replace(s, "$1" & "$3") '1，3替换了所有
