<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim Today
Today = Date
'每日更新数据
If PE_Cache.CacheIsEmpty(ChannelID & "_Photo_ShowHot_date") Then
    PE_Cache.SetValue ChannelID & "_Photo_ShowHot_date", Today
End If
If CurrentPage > 1 Or PE_Cache.CacheIsEmpty(ChannelID & "_HTML_Hot") Or PE_Cache.GetValue(ChannelID & "_Photo_ShowHot_date") <> CStr(Today) Then
    MaxPerPage = MaxPerPage_Hot
    SkinID = DefaultSkinID
    PageTitle = "热点" & ChannelShortName
    strFileName = "ShowHot.asp"
    strHtml = GetTemplate(ChannelID, 8, 0)
    Call PE_Content.GetHtml_List
    If CurrentPage = 1 Then PE_Cache.SetValue ChannelID & "_HTML_Hot", strHtml
    If CurrentPage = 1 Then PE_Cache.SetValue ChannelID & "_Photo_ShowHot_date", Today
Else
    strHtml = PE_Cache.GetValue(ChannelID & "_HTML_Hot")
End If
Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
