<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

SkinID = DefaultSkinID
strFileName = ChannelUrl_ASPFile & "/SpecialList.asp"
PageTitle = "专题列表"
strHtml = GetTemplate(ChannelID, 22, TemplateID)
Call PE_Content.GetHtml_SpecialList
Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
