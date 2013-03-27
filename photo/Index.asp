<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

If CurrentPage > 1 Or PE_Cache.CacheIsEmpty(ChannelID & "_HTML_Index") Then
    MaxPerPage = MaxPerPage_Index
    SkinID = DefaultSkinID
    PageTitle = "首页"
    strPageTitle = strPageTitle & " >> " & PageTitle
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
    strFileName = ChannelUrl_ASPFile & "/Index.asp"
    Call PE_Content.GetHTML_Index
    If CurrentPage = 1 Then PE_Cache.SetValue ChannelID & "_HTML_Index", strHtml
Else
    strHtml = PE_Cache.GetValue(ChannelID & "_HTML_Index")
End If
Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
