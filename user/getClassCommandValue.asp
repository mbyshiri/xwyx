<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private"
Response.CacheControl = "no-cache"
Response.Charset = "GB2312"
Response.ContentType = "text/html"

Dim strSql, ClassID
ClassID = PE_CLng(Trim(Request("ClassID")))
strSql = "Select CommandClassPoint From PE_Class Where ClassID=" & ClassID & ""
Response.Write PE_CLng(Conn.Execute(strSql)(0))
Call CloseConn
%>
