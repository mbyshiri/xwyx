<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

ChannelID = 0
If EnableUserReg <> True Then
	FoundErr = True
	ErrMsg = ErrMsg & "<li>�Բ��𣬱�վ��ͣ���û�ע�����</li>"
	Call WriteErrMsg(ErrMsg, ComeUrl)
	Response.End
End If
%>