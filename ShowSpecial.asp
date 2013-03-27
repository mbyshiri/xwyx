<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<!--#include file="Include/PowerEasy.Class.asp"-->
<!--#include file="Include/PowerEasy.Special.asp"-->
<!--#include file="Include/PowerEasy.Article.asp"-->
<!--#include file="Include/PowerEasy.Soft.asp"-->
<!--#include file="Include/PowerEasy.Photo.asp"-->
<!--#include file="Include/PowerEasy.Product.asp"-->
<!--#include file="Include/PowerEasy.SiteSpecial.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
ChannelID = 0
SpecialID = PE_CLng(Trim(request("SpecialID")))
If SpecialID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>请指定SpecialID！</li>"
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
Call GetSpecial
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

MaxPerPage = MaxPerPage_Special
strFileName = "ShowSpecial.asp?SpecialID=" & SpecialID
PageTitle = ""
strHtml = GetTemplate(0, 30, TemplateID)
Call GetHtml_Special

Response.Write strHtml
Call CloseConn
%>
