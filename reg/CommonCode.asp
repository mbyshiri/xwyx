<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 0
If EnableUserReg <> True Then
	FoundErr = True
	ErrMsg = ErrMsg & "<li>对不起，本站暂停新用户注册服务！</li>"
	Call WriteErrMsg(ErrMsg, ComeUrl)
	Response.End
End If
%>