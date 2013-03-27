<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
If Action<>"" Then
     If FoundInArr(LCase("user,savewrite,del,quintessence"), LCase(Action), ",") = False Then Action=""
End If
strFileName = FileName & "?action=" & Action & "&KindID=" & KindID
SkinID = DefaultSkinID
Select Case Action
Case "savewrite"
    PageTitle = "保存留言"
Case "del"
    PageTitle = "删除留言"
Case Else
    If KindID = 0 Then
        If Action = "user" Then
            If ReplaceBadChar(TopicType) = "participation" Then
                PageTitle = "我回复的留言"
            Else
                PageTitle = "我发表的留言"
            End If
        Else
            PageTitle = XmlText("Guest", "FirstPage", "留言首页")
        End If
    Else
        Dim KindNam, rsKind
        Set rsKind = Conn.Execute("select KindName from PE_Guestkind where KindID=" & KindID)
        If rsKind.BOF And rsKind.EOF Then
            FoundErr = True
            Response.Write XmlText("Guest", "Err1", "输入的类别并不存在！")
        Else
            KindName = rsKind(0)
        End If
        Set rsKind = Nothing
        PageTitle = KindName
    End If
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
strPageTitle = strPageTitle & " >> " & PageTitle
strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
If Action <> "" Or KindID > 0 Or CurrentPage > 1 Then
    Call GetHTML_Index
Else
    If PE_Cache.CacheIsEmpty(ChannelID & "_HTML_Index" & ShowGStyle) Then
        Call GetHTML_Index
        PE_Cache.SetValue ChannelID & "_HTML_Index" & ShowGStyle, strHtml
    Else
        strHtml = PE_Cache.GetValue(ChannelID & "_HTML_Index" & ShowGStyle)
    End If
End If
Response.Write strHtml
Call CloseConn

'=================================================
'过程名：GetHTML_Index()
'作  用：留言首页模板替换函数
'参  数：无
'=================================================
Sub GetHTML_Index()

    Dim strTemp, strTopUser, strFriendSite, arrTemp, strAnnounce, strPopAnnouce
    Dim ArticleList_ChildClass, ArticleList_ChildClass2
    Dim strPicList, strList
    Dim sqlAD, rsAD, ImgUrl, strAD
   
    
    strHtml = GetTemplate(ChannelID, 1, 0)

    'strHTML = Replace(strHTML, "{$GuestMain}", GuestMain())
    
    Dim DefaultIndex
    DefaultIndex = DefaultTemplate("Index")
    strHtml = Replace(strHtml, "{$GuestMain}", DefaultIndex)
    strHtml = Replace(strHtml, "{$KindID}", KindID)
    Call ReplaceCommon
    
    Dim strParameter1, GuestList1, GuestListContent1
    Dim strParameter2, GuestList2, GuestListContent2
    
    regEx.Pattern = "【GuestList1\((.*?)\)】([\s\S]*?)【\/GuestList1】"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        GuestList1 = Match.value
		strParameter1 = Match.SubMatches(0)
		GuestListContent1 = Match.SubMatches(1)
    Next

    regEx.Pattern = "【GuestList2\((.*?)\)】([\s\S]*?)【\/GuestList2】"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        GuestList2 = Match.value
		strParameter2 = Match.SubMatches(0)
		GuestListContent2 = Match.SubMatches(1)
    Next
   
    Select Case Action
    Case "savewrite"
        strHtml = PE_Replace(strHtml, GuestList1, "")
        strHtml = PE_Replace(strHtml, GuestList2, SaveWriteGuest())

    Case "del"
        strHtml = PE_Replace(strHtml, GuestList1, "")
        strHtml = PE_Replace(strHtml, GuestList2, DelGuest())

    Case Else
        If ShowGStyle = 2 Then
            strHtml = PE_Replace(strHtml, GuestList2, "")
            strHtml = PE_Replace(strHtml, GuestList1, GetRepeatGuestBook(strParameter1, GuestListContent1))
        Else
            strHtml = PE_Replace(strHtml, GuestList1, "")
            strHtml = PE_Replace(strHtml, GuestList2, GetRepeatDiscussion(strParameter2, GuestListContent2))
        End If
    End Select

    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Guest", "HTML_Index/PageChar", "条主题"), False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Guest", "HTML_Index/PageChar", "条主题"), False))
End Sub

'=================================================
'函数名：SaveWriteGuest()
'作  用：保存留言
'参  数：无
'=================================================
Private Function SaveWriteGuest()
    Dim GuestName, GuestSex, GuestOicq, GuestEmail, GuestHomepage, GuestFace, GuestImages, GuestIcq, GuestMsn
    Dim GuestTitle, GuestContent, GuestIsPrivate, GuestIsPassed, CheckCode
    Dim sqlMaxId, rsMaxId, MaxId, Saveinfo
    ReplyId = Trim(Request("ReplyId"))

    If ReplyId = "" Then
        ReplyId = 0
    Else
        ReplyId = PE_CLng(ReplyId)
    End If
    
    If GuestBook_EnableVisitor = False Then
        If UserLogined = False Then
            SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Notpermission", "您尚未登录，请登录后再进行您的操作！") & "</li>")
            Exit Function
        Else
            If GroupType < 1 Then
                SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err8", "对不起，您尚未通过邮件验证，不能发表留言！") & "</li>")
                Exit Function
            ElseIf GroupType = 1 Then
                SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err9", "对不起，您尚未通过管理员审核，不能发表留言！") & "</li>")
                Exit Function
            End If
        End If
    End If
    GuestContent = ReplaceBadUrl(ReplaceText(FilterJS(Request("GuestContent")), 4)) '过滤非法系统URL
    If GuestBook_EnableManageRubbish = True And ManageRubbishContent(GuestBook_ManageRubbish, GuestContent) Then
        SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/ForbiddenAD", "您输入的内容涉嫌广告，禁止留言！") & "</li>")
        Exit Function
    End If

    '过虑非法SQL字符,过滤 Jscript代码
    If UserLogined = False Then
        GuestName = PE_HTMLEncode(Trim(Request("GuestName")))
        GuestSex = PE_HTMLEncode(Trim(Request("GuestSex"))) '对方可以用伪造表单提交
        GuestOicq = PE_HTMLEncode(Trim(Request("GuestOicq")))
        GuestIcq = PE_HTMLEncode(Trim(Request("GuestIcq")))
        GuestMsn = PE_HTMLEncode(Trim(Request("GuestMsn")))
        GuestEmail = PE_HTMLEncode(Trim(Request("GuestEmail")))
        GuestHomepage = PE_HTMLEncode(Trim(Request("GuestHomepage")))
        If GuestHomepage = "http://" Or IsNull(GuestHomepage) Then GuestHomepage = ""
    Else
        GuestName = UserName
    End If
    GuestImages = PE_HTMLEncode(Trim(Request("GuestImages")))
    GuestFace = PE_HTMLEncode(Trim(Request("GuestFace")))
    GuestTitle = ReplaceText(PE_HTMLEncode(Trim(Request("GuestTitle"))), 4)
    GuestIsPrivate = Trim(Request("GuestIsPrivate"))
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
    If GuestIsPrivate = "yes" Then
        GuestIsPrivate = True
    Else
        GuestIsPrivate = False
    End If
    
    If CheckLevel = 0 Or NeedlessCheck = 1 Then
        GuestIsPassed = True
    Else
        GuestIsPassed = False
    End If
    
    SaveEdit = Request("saveedit")
    If EnableGuestBookCheck = True Then
        If CheckCode = "" Then
            SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err1", "验证码不能为空！") & "</li>")
            Exit Function
        End If
        If Trim(Session("CheckCode")) = "" Then
            SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err2", "你发留言时间过长，请重发留言。") & "</li>")
            Exit Function
        End If
        If CheckCode <> LCase(Session("CheckCode")) Then
            SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err3", "您输入的确认码和系统产生的不一致！请重新输入。") & "</li>")
            Exit Function
        End If
    End If
    If GuestName = "" Or GuestTitle = "" Or GuestContent = "" Then
        SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err4", "留言发送失败！请将必要的信息填写完整！") & "</li>")
        Exit Function
    End If
    Dim mrs, intMaxID
    Set mrs = Conn.Execute("select max(GuestID) from PE_GuestBook")
    If IsNull(mrs(0)) Then
        intMaxID = 0
    Else
        intMaxID = mrs(0)
    End If
    Set mrs = Nothing
    If SaveEdit = 1 Then
		If UserLogined = False Then
			SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err6", "游客不能编辑留言！如果您需要编辑留言，请以用户身份发表留言！") & "</li>")
			Exit Function
		End If
        SaveEditId = Request("saveeditid")
        If SaveEditId = "" Then
            SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err5", "请指定要编辑的留言ID！") & "</li>")
            Exit Function
        Else
				
			sqlMaxId = "select max(GuestMaxId) as MaxId from PE_GuestBook"
            Set rsMaxId = Conn.Execute(sqlMaxId)
            MaxId = rsMaxId("MaxId")
            Set rsMaxId = Nothing
            If MaxId = "" Or IsNull(MaxId) Then MaxId = 0
            Set rsGuest = Server.CreateObject("adodb.recordset")
            sqlGuest = "select * from PE_GuestBook where GuestID=" & PE_CLng(SaveEditId) & " and GuestName = '" & UserName & "'"
            rsGuest.Open sqlGuest, Conn, 1, 3
            'rsGuest("GuestName") = GuestName
            rsGuest("GuestSex") = GuestSex
            rsGuest("GuestOicq") = GuestOicq
            rsGuest("GuestIcq") = GuestIcq
            rsGuest("GuestMsn") = GuestMsn
            rsGuest("GuestEmail") = GuestEmail
            rsGuest("GuestHomepage") = GuestHomepage
            rsGuest("GuestIP") = UserTrueIP
            rsGuest("GuestTitle") = GuestTitle
            rsGuest("KindID") = KindID
            rsGuest("ReplyNum") = 0
            rsGuest("GuestFace") = GuestFace
            rsGuest("GuestContent") = GuestContent
            rsGuest("GuestDatetime") = Now()
            rsGuest("GuestImages") = GuestImages
            rsGuest("GuestMaxId") = MaxId + 1
            rsGuest("GuestIsPrivate") = GuestIsPrivate
            rsGuest("GuestIsPassed") = GuestIsPassed
            rsGuest("GuestContentLength") = Len(GuestContent)
            rsGuest.Update
            If CheckLevel = 0 Or NeedlessCheck = 1 Then
                SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/SurMsg1", "留言编辑成功！") & "</li>")
            Else
                SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/SurMsg2", "留言编辑成功！只有管理员审核通过的留言才会显示出来。") & "</li>")
            End If
            rsGuest.Close
            Set rsGuest = Nothing
            Call ClearSiteCache(ChannelID)
        End If
    Else
        If GuestContent <> Session("OldGuestContent") Then
            Session("OldGuestContent") = GuestContent
            sqlMaxId = "select max(GuestMaxId) as MaxId from PE_GuestBook"
            Set rsMaxId = Conn.Execute(sqlMaxId)
            MaxId = rsMaxId("MaxId")
            Set rsMaxId = Nothing
            If MaxId = "" Or IsNull(MaxId) Then MaxId = 0
            Set rsGuest = Server.CreateObject("adodb.recordset")
            sqlGuest = "select * from PE_GuestBook"
            rsGuest.Open sqlGuest, Conn, 1, 3
            rsGuest.addnew
            If UserLogined = False Then
                rsGuest("GuestType") = 0
            Else
                rsGuest("GuestType") = 1
            End If
            rsGuest("GuestName") = GuestName
            rsGuest("GuestSex") = GuestSex
            rsGuest("GuestOicq") = GuestOicq
            rsGuest("GuestIcq") = GuestIcq
            rsGuest("GuestMsn") = GuestMsn
            rsGuest("GuestEmail") = GuestEmail
            rsGuest("GuestHomepage") = GuestHomepage
            rsGuest("GuestIP") = UserTrueIP
            rsGuest("GuestTitle") = GuestTitle
            rsGuest("KindID") = KindID
            If ReplyId <> 0 Then
                rsGuest("TopicID") = ReplyId
            Else
                rsGuest("TopicID") = intMaxID + 1
            End If
            rsGuest("ReplyNum") = 0
            rsGuest("GuestFace") = GuestFace
            rsGuest("GuestContent") = GuestContent
            rsGuest("GuestDatetime") = Now()
            rsGuest("GuestImages") = GuestImages
            rsGuest("GuestId") = intMaxID + 1
            rsGuest("GuestMaxId") = MaxId + 1
            rsGuest("GuestIsPrivate") = GuestIsPrivate
            rsGuest("GuestIsPassed") = GuestIsPassed
            rsGuest("GuestContentLength") = Len(GuestContent)
            rsGuest.Update
            If CheckLevel = 0 Or NeedlessCheck = 1 Then
                Saveinfo = "<li>" & XmlText("Guest", "SaveGuest/SurMsg3", "您的留言已经发送成功！") & "</li>"
            Else
                Saveinfo = "<li>" & XmlText("Guest", "SaveGuest/SurMsg4", "您的留言已经发送成功！只有管理员审核通过的留言才会显示出来。") & "</li>"
            End If

            rsGuest.Close
            Set rsGuest = Nothing
            If ReplyId <> 0 And (CheckLevel = 0 Or NeedlessCheck = 1) Then
                'GuestContent = ReplaceBadChar(GuestContent)
                'GuestName = ReplaceBadChar(GuestName)
                'GuestTitle = ReplaceBadChar(GuestTitle)
                'Conn.Execute ("update PE_GuestBook set LastReplyContent='" & GuestContent & "',LastReplyGuest='" & GuestName & "',LastReplyTitle='" & GuestTitle & "',LastReplyTime='" & Now() & "',GuestMaxId=" & MaxId & "+1,ReplyNum=ReplyNum+1 where GuestId=" & ReplyId & "")
                Dim sql, rs, rsReplyNum
                Set rs = Server.CreateObject("adodb.recordset")
                sql = "select top 1 * from PE_GuestBook where GuestId=" & ReplyId
                rs.Open sql, Conn, 1, 3
                If rs.EOF And rs.BOF Then
                    Saveinfo = "<li>" & XmlText("Guest", "SaveGuest/Err7", "找不到您回复的主题！") & "</li>"
                Else
                    rsReplyNum = rs("ReplyNum")
                    rs("LastReplyContent") = GuestContent
                    rs("LastReplyGuest") = GuestName
                    rs("LastReplyTitle") = GuestTitle
                    rs("LastReplyTime") = Now()
                    rs("ReplyNum") = rsReplyNum + 1
                    rs("GuestMaxId") = MaxId + 1
                    rs.Update
                End If
                rs.Close
            End If
            SaveWriteGuest = Guest_info(Saveinfo)
            Call ClearSiteCache(ChannelID)
            Exit Function
        Else
            SaveWriteGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err6", "请不要连续发送内容相同的留言或空留言！") & "</li>")
        End If
    End If
    
End Function

'=================================================
'函数名：DelGuest()
'作  用：删除留言
'参  数：无
'=================================================
Private Function DelGuest()
	If UserLogined = False Then
		DelGuest = Guest_info("<li>" & XmlText("Guest", "SaveGuest/Err7", "游客不能删除留言！") & "</li>")
		Exit Function
	End If
    Dim Delid
    Delid = Trim(Request("guestid"))
    If IsValidID(Delid) = False Then
        Delid = ""
    End If
    If Delid = "" Then
        DelGuest = Guest_info("<li>" & XmlText("Guest", "DelGuest/Err1", "请指定要删除的留言ID！") & "</li>")
        Exit Function
    End If
    If InStr(Delid, ",") > 0 Then
        sqlGuest = "Select * from PE_GuestBook where GuestID in (" & Delid & ")"
    Else
        sqlGuest = "select * from PE_GuestBook where GuestID=" & Delid
    End If
    Set rsGuest = Server.CreateObject("Adodb.RecordSet")
    rsGuest.Open sqlGuest, Conn, 1, 3
    If rsGuest.BOF And rsGuest.EOF Then
        DelGuest = Guest_info("<li>" & XmlText("Guest", "DelGuest/NoFound", "找不到您指定的留言！") & "</li>")
        Exit Function
    End If

    If rsGuest("GuestName") <> UserName Or rsGuest("GuestIsPassed") = True Then
        DelGuest = Guest_info("<li>" & XmlText("Guest", "DelGuest/Err2", "您没有使用此功能的权限！") & "</li>")
    Else
        Do While Not rsGuest.EOF
            rsGuest.Delete
            rsGuest.Update
            rsGuest.MoveNext
        Loop
        DelGuest = Guest_info("<li>" & XmlText("Guest", "DelGuest/SurMsg", "删除留言成功！") & "</li>")
    End If
    rsGuest.Close
    Set rsGuest = Nothing
    Call ClearSiteCache(ChannelID)
End Function
%>
