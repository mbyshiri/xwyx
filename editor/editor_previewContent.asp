<%@language=vbscript codepage=936 %>
<%
option explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<html>
<head>
<title>�Զ����ǩԤ��</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='/Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>
</head>
<body>
<%=replace(Request("Content"),"{$ID}","&")%>
</body>
</html>