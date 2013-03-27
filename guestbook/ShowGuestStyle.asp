<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ShowGStyle
ShowGStyle = request("ShowGStyle")
If ShowGStyle = "" Or Not IsNumeric(ShowGStyle) Then
    ShowGStyle = 1
Else
    ShowGStyle = Int(ShowGStyle)
End If
response.cookies("ShowGStyle") = ShowGStyle
response.redirect request.servervariables("http_referer")
%>
