<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

SpecialID = PE_CLng(Trim(Request("SpecialID")))
If SpecialID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ��SpecialID��</li>"
Else
    Call GetSpecial
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

MaxPerPage = MaxPerPage_Special
strFileName = ChannelUrl_ASPFile & "/ShowSpecial.asp?ClassID=" & ClassID & "&SpecialID=" & SpecialID
PageTitle = ""
strHtml = GetTemplate(ChannelID, 4, TemplateID)
Call PE_Content.GetHtml_Special

Response.Write strHtml
Set PE_Content = Nothing
Call CloseConn
%>
