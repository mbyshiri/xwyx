<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5_New.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "SMS_MessageReceive"   '����Ȩ��    

Dim MD5String
dim PE_MD5
set PE_MD5 = new Md5_Class
MD5String = UCase(Trim(PE_MD5.MD5(SMSUserName & SMSKey)))
set PE_MD5 = nothing

Response.write "<Meta http-equiv='Refresh' Content='0; Url=http://sms.powereasy.net/MessageGate/MessageReceive.aspx?UserName=" & SMSUserName & "&MD5String=" & MD5String & "'>"
%>
