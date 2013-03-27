<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

SkinID = DefaultSkinID
strFileName = "Search.asp?Field=" & strField & "&Keyword=" & Keyword & "&KindID=" & KindID
If KindID = 0 Then
    PageTitle = XmlText("Guest", "ShowIndex/SearchEd", "�������")
Else
    Dim rsKind
    Set rsKind = Conn.Execute("select KindName from PE_Guestkind where KindID=" & KindID)
    If rsKind.BOF And rsKind.EOF Then
        Response.Write "�������𲢲����ڣ�"
        Response.End
    Else
        KindName = rsKind(0)
    End If
    Set rsKind = Nothing
    PageTitle = "<font class=Channel_font>" & KindName & "</font>"
End If
strPageTitle = strPageTitle & " >> " & PageTitle
strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
        
strHtml = GetTemplate(ChannelID, 5, 0)


Dim strTemp, strTopUser, strFriendSite, arrTemp, strAnnounce, strPopAnnouce
Dim strPicList, strList
Dim sqlAD, rsAD, ImgUrl, strAD

strHtml = Replace(strHtml, "{$ResultTitle}", GetResultTitle())

'strHTML = Replace(strHTML, "{$SearchResult}", GetSearchResult())
Dim DefaultSearch
DefaultSearch = DefaultTemplate("Search")
strHtml = Replace(strHtml, "{$SearchResult}", DefaultSearch)
Call ReplaceCommon

Dim strParameter1, GuestList1, GuestListContent1
Dim strParameter2, GuestList2, GuestListContent2

regEx.Pattern = "��GuestList1\((.*?)\)��([\s\S]*?)��\/GuestList1��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
	GuestList1 = Match.value
	strParameter1 = Match.SubMatches(0)
	GuestListContent1 = Match.SubMatches(1)
Next

regEx.Pattern = "��GuestList2\((.*?)\)��([\s\S]*?)��\/GuestList2��"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
	GuestList2 = Match.value
	strParameter2 = Match.SubMatches(0)
	GuestListContent2 = Match.SubMatches(1)
Next

If ShowGStyle = 2 Then
    strHtml = PE_Replace(strHtml, GuestList1, GetRepeatGuestBook(strParameter1, GuestListContent1))
    strHtml = PE_Replace(strHtml, GuestList2, "")
Else

    strHtml = PE_Replace(strHtml, GuestList1, "")
    strHtml = PE_Replace(strHtml, GuestList2, GetRepeatDiscussion(strParameter2, GuestListContent2))
End If

If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Guest", "ShowIndex/PageChar", "������"), False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Guest", "ShowIndex/PageChar", "������"), False))
Response.Write strHtml
Call CloseConn

'=================================================
'��������GetResultTitle()
'��  �ã���ȡ�����������
'��  ������
'=================================================
Private Function GetResultTitle()
    Dim strTitle
    If Keyword = "" Then
        strTitle = XmlText("Guest", "SearchResult/t1", "��������")
    Else
        Select Case strField
        Case "Title"
            strTitle = Replace(XmlText("Guest", "SearchResult/t2", "���⺬�� <font class=Channel_font>{$Keyword}</font>������"), "{$Keyword}", Keyword)
        Case "Content"
            strTitle = Replace(XmlText("Guest", "SearchResult/t3", "���ݺ��� <font class=Channel_font>{$Keyword}</font>������"), "{$Keyword}", Keyword)
        Case "Name"
            strTitle = Replace(XmlText("Guest", "SearchResult/t4", "�����˺��� <font class=Channel_font>{$Keyword}</font>������"), "{$Keyword}", Keyword)
        Case "Reply"
            strTitle = Replace(XmlText("Guest", "SearchResult/t5", "����Ա�ظ����� <font class=Channel_font>{$Keyword}</font>������"), "{$Keyword}", Keyword)
        Case Else
            If IsDate(Trim(Request("keyword"))) = False Then
                Exit Function
            Else
                strTitle = Replace(XmlText("Guest", "SearchResult/t6", "����ʱ��Ϊ<font class=Channel_font>{$Keyword}</font> ������"), "{$Keyword}", Trim(Request("keyword")))
            End If
        End Select
    End If
 
    GetResultTitle = strTitle
End Function
%>
