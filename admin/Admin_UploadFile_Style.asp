<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
