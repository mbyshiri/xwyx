<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

SoftID = PE_CLng(Trim(Request("SoftID")))
If SoftID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ��SoftID��</li>"
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

Dim sql
PageTitle = ChannelShortName & "��Ϣ"
ItemID = SoftID
strFileName = "ShowSoft.asp"
    
Conn.Execute ("update PE_Soft set BrowseTimes=BrowseTimes+1 where SoftID=" & SoftID & "")

sql = "select * from PE_Soft where ChannelID=" & ChannelID & " and Deleted=" & PE_False & " and Status=3 and SoftID=" & SoftID & " and ChannelID=" & ChannelID & ""
Set rsSoft = Conn.Execute(sql)
If rsSoft.BOF And rsSoft.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��Ҫ�ҵ�" & ChannelShortName & "�����ڣ������Ѿ�������Աɾ����</li>"
Else
    ClassID = rsSoft("ClassID")
    If ClassID > 0 Then
        Call GetClass
    Else
        EnableProtect = True
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    rsSoft.Close
    Set rsSoft = Nothing
    Response.End
End If

SoftName = Replace(Replace(Replace(Replace(rsSoft("SoftName") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")

SkinID = GetIDByDefault(rsSoft("SkinID"), DefaultItemSkin)
TemplateID = GetIDByDefault(rsSoft("TemplateID"), DefaultItemTemplate)

strHtml = GetTemplate(ChannelID, 3, TemplateID)
Call PE_Content.GetHtml_Soft
Response.Write strHtml
rsSoft.Close
Set rsSoft = Nothing
Set PE_Content = Nothing
Call CloseConn
%>
