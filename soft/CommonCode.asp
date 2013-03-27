<!--#include file="../Start.asp"-->
<!--#include file="Channel_Config.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Class.asp"-->
<!--#include file="../Include/PowerEasy.Special.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Common.Purview.asp"-->
<!--#include file="../Include/PowerEasy.ConsumeLog.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

UserLogined = CheckUserLogined()

Dim PE_Content
Set PE_Content = New Soft
PE_Content.Init

If CheckPurview_Channel(ChannelPurview, ChannelArrGroupID, UserLogined, GroupID) = False Then
    FoundErr = True
    ErrMsg = ErrMsg & XmlText("BaseText", "ChannelPurviewErr", "<li>对不起，您没有浏览此频道内容的权限！</li>")
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
%>
