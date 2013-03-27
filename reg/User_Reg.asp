<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

strHtml = GetTemplate(0, 18, 0)
Call ReplaceCommonLabel

Dim strPath
strPath = "您现在的位置：&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;服务条款和声明"

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> 服务条款和声明")
strHtml = Replace(strHtml, "{$ShowPath}", strPath)

strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
Response.Write strHtml
Call CloseConn
%>
