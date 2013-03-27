<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5_New.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "SMS_MessageReceive"   '其他权限    

Dim MD5String
dim PE_MD5
set PE_MD5 = new Md5_Class
MD5String = UCase(Trim(PE_MD5.MD5(SMSUserName & SMSKey)))
set PE_MD5 = nothing

Response.write "<Meta http-equiv='Refresh' Content='0; Url=http://sms.powereasy.net/MessageGate/MessageReceive.aspx?UserName=" & SMSUserName & "&MD5String=" & MD5String & "'>"
%>
