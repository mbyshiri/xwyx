<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

strHtml = GetTemplate(0, 18, 0)
Call ReplaceCommonLabel

Dim strPath
strPath = "�����ڵ�λ�ã�&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;�������������"

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> �������������")
strHtml = Replace(strHtml, "{$ShowPath}", strPath)

strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
Response.Write strHtml
Call CloseConn
%>
