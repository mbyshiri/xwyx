<%@language=vbscript codepage=936 %>
<%
option explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<html>
<head>
<title>自定义标签预览</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='/Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>
</head>
<body>
<%=replace(Request("Content"),"{$ID}","&")%>
</body>
</html>