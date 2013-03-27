<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../Include/PowerEasy.UserXml.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
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
