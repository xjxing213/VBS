<html>

<head>
<title>Sample ASP Page 1</title>
</head>

<body>
<P>The time is now <%=Time()%></P>
<%
  Dim iHour

  iHour = Hour(Time())

  If (iHour >= 0 And iHour < 12 ) Then
%>
Good Morning!
<%
  ElseIf (iHour > 11 And iHour < 5 ) Then
%>
Good Afternoon!
<%
  Else
%>
Good Evening!
<%
End If
%>

</body>
</html>

