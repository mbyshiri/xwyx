<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If GuestBook_EnableVisitor = False Then
    If UserLogined = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "ShowWrite/Notpermission", "����δ��¼�����¼���ٽ������Ĳ�����") & "</li>"
    Else
        If GroupType < 1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "SaveGuest/Err8", "�Բ�������δͨ���ʼ���֤�����ܷ������ԣ�") & "</li>"
        ElseIf GroupType = 1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "SaveGuest/Err9", "�Բ�������δͨ������Ա��ˣ����ܷ������ԣ�") & "</li>"
        End If
    End If 
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

KindName = ReplaceBadChar(Trim(Request("KindName")))
SkinID = DefaultSkinID
If Action = "edit" Then
    PageTitle = "�༭����"
Else
    If KindName <> "" Then
        PageTitle = KindName & " >> " & "ǩд����"
    Else
        PageTitle = "ǩд����"
    End If
End If

strPageTitle = strPageTitle & " >> " & PageTitle
strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
If SaveEdit <> 1 Then
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
End If

strHtml = GetTemplate(ChannelID, 3, 0)


Dim strTemp, strTopUser, strFriendSite, arrTemp, strAnnounce, strPopAnnouce, iCols, iClassID
Dim ArticleList_ChildClass, ArticleList_ChildClass2
Dim strPicList, strList
Dim sqlAD, rsAD, ImgUrl, strAD

'strHTML = Replace(strHTML, "{$WriteGuest}", WriteGuest)
Dim DefaultWrite
DefaultWrite = DefaultTemplate("strWrite")
strHtml = Replace(strHtml, "{$WriteGuest}", DefaultWrite)
Call ReplaceCommon

Dim GuestEditList, strParameter, GuestEditListContent
regEx.Pattern = "��GuestList1\((.*?)\)��([\s\S]*?)��\/GuestList1��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    GuestEditList = Match.value
	strParameter = Match.SubMatches(0)
	GuestEditListContent = Match.SubMatches(1)
Next

If Action = "edit" Then
	If UserLogined = False Then
		FoundErr = True
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "SaveGuest/Err6", "�οͲ��ܱ༭���ԣ��������Ҫ�༭���ԣ������û���ݷ������ԣ�") & "</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
		Response.End
	End If
    Dim EditId
    EditId = Request("guestid")
    If EditId = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "EditGuest/Err1", "��ָ��Ҫ�༭������ID��") & "</li>"
    Else
        EditId = PE_CLng(EditId)
        sqlGuest = "select * from PE_GuestBook where GuestId=" & EditId & " and GuestName='" & UserName & "'"
    End If
    Set rsGuest = Server.CreateObject("adodb.recordset")
    rsGuest.Open sqlGuest, Conn, 1, 1
    If rsGuest.BOF And rsGuest.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "EditGuest/NoFound", "�Ҳ�����ָ�������ԣ�") & "</li>"
    End If

    If rsGuest("GuestName") = UserName And rsGuest("GuestIsPassed") = False Then
        WriteTopicID = rsGuest("TopicID")
        WriteName = rsGuest("GuestName")
        WriteType = rsGuest("GuestType")
        WriteSex = rsGuest("GuestSex")
        WriteEmail = rsGuest("GuestEmail") & ""
        WriteOicq = rsGuest("GuestOicq") & ""
        WriteIcq = rsGuest("GuestIcq") & ""
        WriteMsn = rsGuest("GuestMsn") & ""
        WriteHomepage = rsGuest("GuestHomepage") & ""
        WriteFace = rsGuest("GuestFace")
        WriteImages = rsGuest("GuestImages")
        WriteTitle = rsGuest("GuestTitle")
        WriteContent = rsGuest("GuestContent")
        WriteIsPrivate = rsGuest("GuestIsPrivate")
        SaveEdit = 1
        SaveEditId = EditId
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "EditGuest/Err2", "�û�ֻ���Ա༭�Լ��������δͨ��������ԣ�") & "</li>"
    End If
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Response.End
    End If
Else
    strHtml = PE_Replace(strHtml, GuestEditList, "")
End If

strHtml = PE_Replace(strHtml, GuestEditList, GetRepeatGuestBook(strParameter, GuestEditListContent))
strHtml = Replace(strHtml, "{$ShowJS_Guest}", ShowJS_Guest)
strHtml = Replace(strHtml, "{$GuestFace}", GuestFace)
strHtml = Replace(strHtml, "{$WriteName}", WriteName)
If GuestBook_IsAssignSort = True Then
    strHtml = Replace(strHtml, "{$GetGKind_Option}", GetGKind_Option(3, 0))
Else
    strHtml = Replace(strHtml, "{$GetGKind_Option}", GetGKind_Option(1, 0))
End If
strHtml = Replace(strHtml, "{$WriteEmail}", WriteEmail)
strHtml = Replace(strHtml, "{$WriteOicq}", WriteOicq)
strHtml = Replace(strHtml, "{$WriteIcq}", WriteIcq)
strHtml = Replace(strHtml, "{$WriteMsn}", WriteMsn)
strHtml = Replace(strHtml, "{$WriteHomepage}", WriteHomepage)
strHtml = Replace(strHtml, "{$GuestContent}", GuestContent)
strHtml = Replace(strHtml, "{$WriteTitle}", WriteTitle)
strHtml = Replace(strHtml, "{$saveedit}", SaveEdit)
strHtml = Replace(strHtml, "{$saveeditid}", SaveEditId)
strHtml = Replace(strHtml, "{$ReplyId}", ReplyId)

Dim GuestBookList, GuestBookKind, GuestBookFace, strGuestBookCheck

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
If WriteType <> 1 Then
    strHtml = Replace(strHtml, GuestBookFace, "")
End If
strHtml = Replace(strHtml, "��GuestBookFace��", "")
strHtml = Replace(strHtml, "��/GuestBookFace��", "")

regEx.Pattern = "��GuestBookList��([\s\S]*?)��\/GuestBookList��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    GuestBookList = Match.value
Next
If WriteType <> 0 Then
    strHtml = Replace(strHtml, GuestBookList, "")
End If
strHtml = Replace(strHtml, "��GuestBookList��", "")
strHtml = Replace(strHtml, "��/GuestBookList��", "")

regEx.Pattern = "��GuestBookKind��([\s\S]*?)��\/GuestBookKind��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    GuestBookKind = Match.value
Next

'���Ǳ༭״̬�ֲ������⣬����ʾ���ѡ��
If Action = "edit" And WriteTopicID <> SaveEditId Then
    strHtml = Replace(strHtml, GuestBookKind, "")
End If

'��������״̬�µ��ǩд���ԣ��������������
If KindID <> 0 Then
    strHtml = Replace(strHtml, GuestBookKind, "<Input type=hidden value='" & KindID & "' name=KindID>")
End If
strHtml = Replace(strHtml, "��GuestBookKind��", "")
strHtml = Replace(strHtml, "��/GuestBookKind��", "")

strHtml = Replace(strHtml, "value= ", "value='' ")
strHtml = Replace(strHtml, "Value= ", "value='' ")

Response.Write strHtml
Call CloseConn
%>
