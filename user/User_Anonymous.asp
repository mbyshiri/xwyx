<!--#include file="User_Anonymous_Code.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************
Dim arrstrHtml
ChannelID = 0
SkinID = DefaultSkinID
PageTitle = "�ο�Ͷ��"
strPageTitle = SiteTitle & " >> " & PageTitle
strHtml = GetTemplate(0, 103, 0)
If strHtml = XmlText("BaseText", "TemplateErr", "�Ҳ���ģ��") Then
    Response.Write strHtml
    Response.End	        
End If	
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
Call ReplaceCommonLabel
strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strPageTitle)
If Instr(strHtml,"{$MainContent}") = 0 Then 
    Response.Write "��Ա����ͨ��ģ������������ģ������'{$MainContent}'����ο�Ĭ��ģ�塣"
    Response.End		
End If
arrstrHtml = Split(strHtml,"{$MainContent}")
Response.Write arrstrHtml(0)
call Execute
Response.Write arrstrHtml(1)
Call CloseConn
%>