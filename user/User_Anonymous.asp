<!--#include file="User_Anonymous_Code.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
Dim arrstrHtml
ChannelID = 0
SkinID = DefaultSkinID
PageTitle = "游客投稿"
strPageTitle = SiteTitle & " >> " & PageTitle
strHtml = GetTemplate(0, 103, 0)
If strHtml = XmlText("BaseText", "TemplateErr", "找不到模板") Then
    Response.Write strHtml
    Response.End	        
End If	
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
Call ReplaceCommonLabel
strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strPageTitle)
If Instr(strHtml,"{$MainContent}") = 0 Then 
    Response.Write "会员中心通用模板里面必须包含模板内容'{$MainContent}'，请参考默认模板。"
    Response.End		
End If
arrstrHtml = Split(strHtml,"{$MainContent}")
Response.Write arrstrHtml(0)
call Execute
Response.Write arrstrHtml(1)
Call CloseConn
%>