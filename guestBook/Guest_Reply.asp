<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

ReplyId = PE_CLng(Request("TopicID"))
If ReplyId = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "ShowReply/NoTopicID", "��ָ����������ID��") & "</li>"
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
Dim strTemp, strTopUser, strFriendSite, arrTemp, strAnnounce, strPopAnnouce, iCols, iClassID
Dim ArticleList_ChildClass, ArticleList_ChildClass2
Dim strPicList, strList
Dim sqlAD, rsAD, ImgUrl, strAD
Dim DefaultReply, strGuestBookCheck

sqlGuest = "select * from PE_GuestBook where GuestIsPassed=" & PE_True & " and TopicID=" & PE_CLng(ReplyId) & " order by GuestId asc "
Set rsGuest = Server.CreateObject("adodb.recordset")
rsGuest.Open sqlGuest, Conn, 1, 1
If rsGuest.BOF And rsGuest.EOF Then
    totalPut = 0
    FoundErr = True
    ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "EditGuest/Err3", "�Ҳ���ָ�������ԣ����ܴ����Ի�δͨ����˻����Ѿ���ɾ����") & "</li>"
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

Conn.Execute ("update PE_GuestBook set Hits=Hits+1 where GuestID=" & ReplyId & "")
MaxPerPage = ReplyMaxPerPage

strFileName = "Guest_Reply.asp?TopicID=" & ReplyId
SkinID = DefaultSkinID
PageTitle = "�鿴����"
strPageTitle = strPageTitle & " >> " & PageTitle
strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
If UserLogined = False Then
    WriteType = 0
Else
    WriteType = 1
End If
WriteName = UserName
WriteSex = "1"
WriteFace = "1"
WriteImages = "01"
WriteHomepage = "http://"
WriteIsPrivate = False

strHtml = GetTemplate(ChannelID, 4, 0)


DefaultReply = DefaultTemplate("Reply")
strHtml = Replace(strHtml, "{$ReplyGuest}", DefaultReply)
strHtml = Replace(strHtml, "{$TopicID}", ReplyId)
Call ReplaceCommon

strHtml = Replace(strHtml, "{$WriteTitle}", "Re: " & rsGuest("GuestTitle"))
strHtml = Replace(strHtml, "{$WriteKindID}", rsGuest("KindId") & "")

Dim GuestBookLList, GuestBookFace
regEx.Pattern = "��GuestList1\((.*?)\)��([\s\S]*?)��\/GuestList1��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    strHtml = PE_Replace(strHtml, Match.value, GetRepeatGuestBook(Match.SubMatches(0), Match.SubMatches(1)))
Next


strHtml = Replace(strHtml, "{$ShowJS_Guest}", ShowJS_Guest())
strHtml = Replace(strHtml, "{$GuestFace}", GuestFace())
strHtml = Replace(strHtml, "{$GuestContent}", GuestContent())

strHtml = Replace(strHtml, "{$saveedit}", SaveEdit)
strHtml = Replace(strHtml, "{$saveeditid}", SaveEditId)
strHtml = Replace(strHtml, "{$ReplyId}", ReplyId)

regEx.Pattern = "��GuestBookCheck��([\s\S]*?)��\/GuestBookCheck��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    strGuestBookCheck = Match.value
Next
If EnableGuestBookCheck = False Then
    strHtml = Replace(strHtml, strGuestBookCheck, "")
End If
strHtml = Replace(strHtml, "��GuestBookCheck��", "")
strHtml = Replace(strHtml, "��/GuestBookCheck��", "")

regEx.Pattern = "��GuestBookFace��([\s\S]*?)��\/GuestBookFace��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    GuestBookFace = Match.value
Next
If WriteType = 0 Then
     strHtml = Replace(strHtml, GuestBookFace, "")
End If
strHtml = Replace(strHtml, "��GuestBookFace��", "")
strHtml = Replace(strHtml, "��/GuestBookFace��", "")

regEx.Pattern = "��GuestBookLList��([\s\S]*?)��\/GuestBookLList��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    GuestBookLList = Match.value
Next
If WriteType <> 0 Then
     strHtml = Replace(strHtml, GuestBookLList, "")
End If
strHtml = Replace(strHtml, "��GuestBookLList��", "")
strHtml = Replace(strHtml, "��/GuestBookLList��", "")

If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Guest", "ShowReply/PageChar", "������"), False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Guest", "ShowReply/PageChar", "������"), False))

strHtml = Replace(strHtml, "value= ", "value='' ")
strHtml = Replace(strHtml, "Value= ", "value='' ")

Response.Write strHtml
Set rsGuest = Nothing
Call CloseConn

%>
