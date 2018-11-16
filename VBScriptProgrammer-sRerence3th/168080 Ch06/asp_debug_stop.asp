<%@ Language=VBScript %>
<% Option Explicit %>
<%

Dim strResult

Call Main()

Sub Main
    Dim x, y, z

    Stop
    x = 5
    y = 100
    z = y / x
    strResult = z
End Sub

%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>VBScript Programmer's Reference - ASP Debugger Example</title>
</head>
<body>
<b>Result of Server-Side Main() Script:</b> <%=strResult%>
</body>
</html>