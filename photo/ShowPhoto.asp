<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

PhotoID = PE_CLng(Trim(Request("PhotoID")))
If PhotoID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ��PhotoID��</li>"
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

Dim sql
PageTitle = ChannelShortName & "��Ϣ"
ItemID = PhotoID
strFileName = ChannelUrl_ASPFile & "/ShowPhoto.asp"
    
sql = "select * from PE_Photo where Deleted=" & PE_False & " and Status=3 and PhotoID=" & PhotoID & " and ChannelID=" & ChannelID & ""
Set rsPhoto = Conn.Execute(sql)
If rsPhoto.BOF And rsPhoto.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��Ҫ�ҵ�" & ChannelShortName & "�����ڣ������Ѿ�������Աɾ����</li>"
Else
    ClassID = rsPhoto("ClassID")
    If ClassID > 0 Then
        Call GetClass
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    rsPhoto.Close
    Set rsPhoto = Nothing
    Response.End
End If
    
Dim sqlHits, LastHitTime
LastHitTime = rsPhoto("LastHitTime")
sqlHits = "update PE_Photo set Hits=Hits+1,LastHitTime=" & PE_Now & ""
If DateDiff("D", LastHitTime, Now()) <= 0 Then
    sqlHits = sqlHits & ",DayHits=DayHits+1"
Else
    sqlHits = sqlHits & ",DayHits=1"
End If
If DateDiff("ww", LastHitTime, Now()) <= 0 Then
    sqlHits = sqlHits & ",WeekHits=WeekHits+1"
Else
    sqlHits = sqlHits & ",WeekHits=1"
End If
If DateDiff("m", LastHitTime, Now()) <= 0 Then
    sqlHits = sqlHits & ",MonthHits=MonthHits+1"
Else
    sqlHits = sqlHits & ",MonthHits=1"
End If
sqlHits = sqlHits & " where PhotoID=" & PhotoID
Conn.Execute (sqlHits)

PhotoName = Replace(Replace(Replace(Replace(rsPhoto("PhotoName") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
SkinID = GetIDByDefault(rsPhoto("SkinID"), DefaultItemSkin)
TemplateID = GetIDByDefault(rsPhoto("TemplateID"), DefaultItemTemplate)

strHtml = GetTemplate(ChannelID, 3, TemplateID)
Call PE_Content.GetHtml_Photo
Call PE_Content.ReplaceViewPhoto
Response.Write strHtml

rsPhoto.Close
Set rsPhoto = Nothing
Set PE_Content = Nothing
Call CloseConn
%>
