<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ShowFileStyle
ShowFileStyle = request("ShowFileStyle")
If ShowFileStyle = "" Or Not IsNumeric(ShowFileStyle) Then
    ShowFileStyle = 1
Else
    ShowFileStyle = Int(ShowFileStyle)
End If
Response.cookies("ShowFileStyle") = ShowFileStyle
Response.redirect request.servervariables("http_referer")
%>
