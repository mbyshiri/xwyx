<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../Include/PowerEasy.UserXml.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************
Dim CookieDate
Response.ContentType = "text/xml; charset=gb2312"
CookieDate = 0
If CheckUserLogined() = False Then
    ErrMsg = ""
    Call ShowUserErr
Else
    Call GetUser(UserName)
    Call ShowUserXml(False)
End If

Call CloseConn
%>
