<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If CurrentPage > 1 Or PE_Cache.CacheIsEmpty(ChannelID & "_HTML_New") Then
    MaxPerPage = MaxPerPage_New
    SkinID = DefaultSkinID
    PageTitle = "����" & ChannelShortName
    strFileName = "ShowNew.asp"
    strHtml = GetTemplate(ChannelID, 6, 0)
    Call PE_Content.GetHtml_List
    If CurrentPage = 1 Then PE_Cache.SetValue ChannelID & "_HTML_New", strHtml
Else
    strHtml = PE_Cache.GetValue(ChannelID & "_HTML_New")
End If
Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
