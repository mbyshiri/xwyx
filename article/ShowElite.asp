<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If CurrentPage > 1 Or PE_Cache.CacheIsEmpty(ChannelID & "_HTML_Elite") Then
    MaxPerPage = MaxPerPage_Elite
    SkinID = DefaultSkinID
    PageTitle = "�Ƽ�" & ChannelShortName
    strFileName = "ShowElite.asp"
    strHtml = GetTemplate(ChannelID, 7, 0)
    Call PE_Content.GetHtml_List
    If CurrentPage = 1 Then PE_Cache.SetValue ChannelID & "_HTML_Elite", strHtml
Else
    strHtml = PE_Cache.GetValue(ChannelID & "_HTML_Elite")
End If
Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
