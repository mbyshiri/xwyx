<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim SkinID
action = Trim(Request("Action"))
ComeUrl = Request.ServerVariables("HTTP_REFERER")
SkinID = Trim(Request("SkinID"))

If action = "SetSkin" Then
    If SkinID = "" Then
        SkinID = 0
    Else
        SkinID = CLng(SkinID)
    End If
    Response.Cookies("asp163")("SkinID") = SkinID
End If
Response.Redirect ComeUrl
%>
