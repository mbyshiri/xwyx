<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
