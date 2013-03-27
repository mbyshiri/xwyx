<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2008 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 0
PageTitle = "版权申明"

strHTML = GetTemplate(ChannelID, 7, 0)

Call ReplaceCommonLabel

strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHTML = Replace(strHTML, "{$ShowPath}", strNavPath)

strHTML = Replace(strHTML, "{$MenuJS}", GetMenuJS("", False))
strHTML = Replace(strHTML, "{$Skin_CSS}", GetSkin_CSS(0))
Response.Write strHTML
%>
