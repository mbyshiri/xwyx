<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<!--#include file="Include/PowerEasy.Class.asp"-->
<!--#include file="Include/PowerEasy.Special.asp"-->
<!--#include file="Include/PowerEasy.Article.asp"-->
<!--#include file="Include/PowerEasy.Soft.asp"-->
<!--#include file="Include/PowerEasy.Photo.asp"-->
<!--#include file="Include/PowerEasy.Product.asp"-->
<!--#include file="Include/PowerEasy.SiteSpecial.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

ChannelID = 0

MaxPerPage = 20
SkinID = DefaultSkinID
strFileName = "SpecialList.asp"
PageTitle = "ȫվר���б�"
strPageTitle = SiteTitle & " >> " & PageTitle
strHtml = GetTemplate(0, 29, TemplateID)
Call GetHtml_SpecialList

Response.Write strHtml
Call CloseConn
%>
