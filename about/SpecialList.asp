<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

SkinID = DefaultSkinID
strFileName = ChannelUrl_ASPFile & "/SpecialList.asp"
PageTitle = "ר���б�"
strHtml = GetTemplate(ChannelID, 22, TemplateID)
Call PE_Content.GetHtml_SpecialList
Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
