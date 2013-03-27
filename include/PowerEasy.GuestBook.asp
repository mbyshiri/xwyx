<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 4

'Private GuestBook_ManageRubbish, GuestBook_EnableVisitor,
Dim IndexMaxPerPage1, IndexMaxPerPage2, ReplyMaxPerPage, TreeMaxPerPage
'Private EnableGuestBookCheck, GuestBook_IsAssignSort, GuestBook_ShowIP, GuestBook_EnableManageRubbish

'Private NeedlessCheck, arrClass_Input, UserSetting

Private testHTML

Private rsGuest, sqlGuest
Private ReplyId, ShowGStyle

Private WriteName, WriteType, WriteSex, WriteEmail, WriteOicq, WriteIcq, WriteMsn, WriteTopicID
Private WriteHomepage, WriteFace, WriteImages, WriteTitle, WriteContent, WriteIsPrivate, WriteKindId
Private SaveEdit, SaveEditId

Private GImagePath, GFacePath
Private KindName
Dim TopicType
Dim arrMaxPerPage

XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

If IsNull(GuestBook_MaxPerPage) Then
    GuestBook_MaxPerPage = "20|||10|||10|||5"
End If
arrMaxPerPage = Split(GuestBook_MaxPerPage, "|||")
'讨论区
IndexMaxPerPage1 = PE_CLng(arrMaxPerPage(0))
'留言本
IndexMaxPerPage2 = PE_CLng(arrMaxPerPage(1))
'回复页
ReplyMaxPerPage = PE_CLng(arrMaxPerPage(2))
'展开树
TreeMaxPerPage = arrMaxPerPage(3)


'检查用户是否登录
UserLogined = CheckUserLogined()
strNavPath = XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
strPageTitle = SiteTitle

Call GetChannel(ChannelID)

If Trim(ChannelName) <> "" And ShowChannelName = True Then
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a href='" & InstallDir & ChannelDir & "/Index.asp'>" & ChannelName & "</a>"
    strPageTitle = strPageTitle & " >> " & ChannelName
End If

'留言板内部参数
SaveEdit = 0
GImagePath = InstallDir & "GuestBook/Images/"
GFacePath = InstallDir & "GuestBook/Images/Face/"
FileName = "index.asp"
KindID = PE_CLng(Trim(Request("KindID")))

TopicType = Trim(Request("topictype"))

'读取查看方式
ShowGStyle = GuestStyle()

If ShowGStyle = 2 Then
    MaxPerPage = IndexMaxPerPage2
Else
    MaxPerPage = IndexMaxPerPage1
End If

Private Sub ReplaceCommon()
    
    Call ReplaceCommonLabel
    
    strHtml = Replace(strHtml, "{$GuestBook_Search}", GuestBook_Search())
    strHtml = Replace(strHtml, "{$GuestBook_top}", GuestBook_Top())
    strHtml = Replace(strHtml, "{$GuestBook_Mode}", GuestBook_Mode())
    strHtml = Replace(strHtml, "{$GetGKindList}", GetGKindList())
    strHtml = Replace(strHtml, "{$ShowGueststyle}", ShowGueststyle())
    strHtml = Replace(strHtml, "{$GuestBook_See}", GuestBook_See())
    strHtml = Replace(strHtml, "{$GuestBook_Appear}", GuestBook_Appear())
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & InstallDir & "Rss.asp?ChannelID=" & ChannelID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
End Sub
'=================================================
'过程名：GuestBook_Top()
'作  用：显示顶部留言功能
'参  数：无
'=================================================
Private Function GuestBook_Top()
    Dim strTop
    strTop = ""
    If UserLogined = True Then
        strTop = Replace(XmlText("Guest", "GuestBook_Top/Logined", "<a href='{$FileName}?action=user'><img align=absmiddle src='images/Guest_user.gif' alt='我发表的留言' border='0'></a>&nbsp;<a href='{$FileName}?action=user&topictype=participation'><img align=absmiddle src='images/Guest_participation.gif' alt='我回复的留言' border='0'></a>&nbsp;"), "{$FileName}", FileName)
    End If
    strTop = strTop & Replace(Replace(Replace(XmlText("Guest", "GuestBook_Top/NoLogin", "<a href='{$FileName}'><img align=absmiddle src='images/Guest_all.gif' alt='查看所有留言' border='0'></a>&nbsp;<a href='Guest_Write.asp?KindID={$KindID}&KindName={$KindName}'><img align=absmiddle src='images/Guest_write.gif' alt='签写新的留言' border='0'></a>"), "{$FileName}", FileName), "{$KindID}", KindID), "{$KindName}", KindName) & vbCrLf
    GuestBook_Top = strTop
End Function
'=================================================
'过程名：GuestBook_Mode()
'作  用：显示顶部留言功能
'参  数：无
'=================================================
Private Function GuestBook_Mode()
    Dim strTop
    strTop = ""
    If UserLogined = True Then
        strTop = strTop & XmlText("Guest", "GuestBook_Mode/Mode1", "用户模式") & vbCrLf
    Else
        strTop = strTop & XmlText("Guest", "GuestBook_Mode/Mode2", "游客模式") & vbCrLf
    End If
    GuestBook_Mode = strTop
End Function
'=================================================
'过程名：GuestBook_See()
'作  用：显示顶部留言功能
'参  数：无
'=================================================
Private Function GuestBook_See()
    Dim strTop
    strTop = ""
    If ShowGStyle = 1 Then
        strTop = strTop & XmlText("Guest", "GuestBook_See/Mode1", "讨论区方式") & vbCrLf
    Else
        strTop = strTop & XmlText("Guest", "GuestBook_See/Mode2", "留言板方式") & vbCrLf
    End If

    GuestBook_See = strTop
End Function
'=================================================
'过程名：GuestBook_Appear()
'作  用：显示顶部留言功能
'参  数：无
'=================================================
Private Function GuestBook_Appear()
    Dim strTop
    If CheckLevel = 0 Or NeedlessCheck = 1 Then
        strTop = strTop & XmlText("Guest", "GuestBook_Appear/Mode1", "直接发表") & vbCrLf
    Else
        strTop = strTop & XmlText("Guest", "GuestBook_Appear/Mode2", "审核发表") & vbCrLf
        Dim grs
        Set grs = Conn.Execute("select count(*) from PE_GuestBook where GuestIsPassed=" & PE_False & "")
        strTop = strTop & "&nbsp;&nbsp;" & Replace(XmlText("Guest", "GuestBook_Appear/Count", "有{$GuestNo}条待审核"), "{$GuestNo}", grs(0)) & vbCrLf
        Set grs = Nothing
    End If
    GuestBook_Appear = strTop
End Function

'=================================================
'过程名：GuestBook_Search()
'作  用：显示留言搜索
'参  数：无
'=================================================
Private Function GuestBook_Search()
    Dim strGuestSearch
    'If GuestBook_IsAssignSort = True Then
        'strGuestSearch = Replace(XmlText("Guest", "GuestBook_Search", "<table border='0' cellpadding='0' cellspacing='0'><form method='post' name='SearchForm' action='Search.asp'><tr><td height='30' >&nbsp;&nbsp;<select name='Field' id='1'><option value='Title' selected>留言主题</option><option value='Content'>留言内容</option><option value='Name'>留言人</option><option value='GuestTime'>留言时间</option><option value='Reply'>管理员回复</option></select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<select name='KindID' id='KindID'>{$KindID}</select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<input type='text' name='keyword'  size='15' value='关键字' maxlength='45' onFocus='this.select();'>&nbsp;<input type='submit' name='Submit'  value='搜索'></td></tr></form></table>"), "{$KindID}", GetGKind_Option(3, KindID))
    'Else
        strGuestSearch = Replace(XmlText("Guest", "GuestBook_Search", "<table border='0' cellpadding='0' cellspacing='0'><form method='post' name='SearchForm' action='Search.asp'><tr><td height='30' >&nbsp;&nbsp;<select name='Field' id='1'><option value='Title' selected>留言主题</option><option value='Content'>留言内容</option><option value='Name'>留言人</option><option value='GuestTime'>留言时间</option><option value='Reply'>管理员回复</option></select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<select name='KindID' id='KindID'>{$KindID}</select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<input type='text' name='keyword'  size='15' value='关键字' maxlength='45' onFocus='this.select();'>&nbsp;<input type='submit' name='Submit'  value='搜索'></td></tr></form></table>"), "{$KindID}", GetGKind_Option(1, KindID))
    'End If
    GuestBook_Search = strGuestSearch
End Function




'=================================================
'过程名：ShowAllGuest()
'作  用：分页显示所有留言
'参  数：ShowType-----  0为显示所有
'                       1为显示已通过审核及用户自己发表的留言
'                       2为显示已通过审核的留言（用于游客显示）
'                       3为显示用户自己发表的留言
'                       4为显示推荐精华的留言
'                       5为要编辑的留言
'                       6为回复页的留言
'=================================================
Private Sub ShowAllGuest(ShowType)
    Select Case ShowType
    Case 1
        sqlGuest = "select * from PE_GuestBook where (GuestIsPassed=" & PE_True & " or GuestName='" & UserName & "')"
    Case 2
        sqlGuest = "select * from PE_GuestBook where GuestIsPassed=" & PE_True & ""
    Case 3
        If TopicType <> "" Then
            TopicType = ReplaceBadChar(TopicType)
        End If
        If TopicType = "participation" Then
            sqlGuest = "select * from PE_GuestBook where GuestID in (select TopicID from PE_GuestBook where GuestName='" & UserName & "' and TopicID<>GuestId)"
        Else
            sqlGuest = "select * from PE_GuestBook where GuestName='" & UserName & "'"
        End If
    Case 4
        sqlGuest = "select * from PE_GuestBook where GuestIsPassed=" & PE_True & " and Quintessence=1"
    Case 5
        sqlGuest = "select * from PE_GuestBook where GuestId=" & PE_CLng(Request("guestid"))
    Case 6
        sqlGuest = "select * from PE_GuestBook where GuestIsPassed=" & PE_True & " and TopicID=" & PE_CLng(ReplyId) & " order by GuestId asc "
    Case Else
        sqlGuest = "select * from PE_GuestBook where 1=1"
    End Select
    If Keyword <> "" Then
        Select Case strField
            Case "Title"
                sqlGuest = sqlGuest & " and GuestTitle like '%" & Keyword & "%' "
            Case "Content"
                sqlGuest = sqlGuest & " and GuestContent like '%" & Keyword & "%' "
            Case "Name"
                sqlGuest = sqlGuest & " and GuestName like '%" & Keyword & "%' "
            Case "Reply"
                sqlGuest = sqlGuest & " and GuestReply like '%" & Keyword & "%' "
            Case Else
                If IsDate(Trim(Request("keyword"))) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "ShowAllGuest/Err1", "输入的关键字不是有效日期！") & "</li>"
                    Exit Sub
                Else
                    If SystemDatabaseType = "SQL" Then
                        sqlGuest = sqlGuest & " and GuestDatetime = '" & Trim(Request("keyword")) & "' "
                    Else
                        sqlGuest = sqlGuest & " and GuestDatetime = #" & Trim(Request("keyword")) & "# "
                    End If
                End If
        End Select
    End If
    If KindID <> "" And KindID <> "0" Then
        sqlGuest = sqlGuest & " and KindID =" & KindID
    End If

    If strField = "" And ShowType <> 5 And ShowType <> 6 Then
        sqlGuest = sqlGuest & " and TopicID =GuestId"
    End If
    If ShowType <> 6 Then
        sqlGuest = sqlGuest & " order by Ontop desc,GuestMaxId desc"
    End If

    Set rsGuest = Server.CreateObject("adodb.recordset")
    rsGuest.Open sqlGuest, Conn, 1, 1
    If rsGuest.BOF And rsGuest.EOF Then
        totalPut = 0
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "ShowAllGuest/NoFound", "没有任何留言") & "</li>"
        Exit Sub
    Else
        totalPut = rsGuest.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsGuest.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    End If
End Sub

'=================================================
'过程名：ShowJS_Guest()
'作  用：提交留言的输入判断
'参  数：无
'=================================================
Private Function ShowJS_Guest()
    Dim strJS
    strJS = "<script language = 'JavaScript'>" & vbCrLf
    strJS = strJS & "function changeimage()" & vbCrLf
    strJS = strJS & "{" & vbCrLf
    strJS = strJS & "  document.myform.GuestImages.value=document.myform.Image.value;" & vbCrLf
    strJS = strJS & "  document.myform.showimages.src='" & GFacePath & "'+document.myform.Image.value+'.gif';" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function guestpreview()" & vbCrLf
    strJS = strJS & "{" & vbCrLf
    strJS = strJS & "  document.preview.content.value=document.myform.GuestContent.value;" & vbCrLf
    strJS = strJS & "  var popupWin = window.open('GuestPreview.asp', 'GuestPreview', 'scrollbars=yes,width=620,height=230');" & vbCrLf
    strJS = strJS & "  document.preview.submit();" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function CheckForm()" & vbCrLf
    strJS = strJS & "{" & vbCrLf
    If UserLogined = False Then
        strJS = strJS & "    if(document.myform.GuestName.value==''){" & vbCrLf
        strJS = strJS & "      alert('姓名不能为空！');" & vbCrLf
        strJS = strJS & "      document.myform.GuestName.focus();" & vbCrLf
        strJS = strJS & "      return(false) ;" & vbCrLf
        strJS = strJS & "    }" & vbCrLf
    End If
    strJS = strJS & "  if(document.myform.GuestTitle.value==''){" & vbCrLf
    strJS = strJS & "    alert('主题不能为空！');" & vbCrLf
    strJS = strJS & "    document.myform.GuestTitle.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  if(document.myform.GuestTitle.value.length>30){" & vbCrLf
    strJS = strJS & "    alert('主题不能超过30字符！');" & vbCrLf
    strJS = strJS & "    document.myform.GuestTitle.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "   var IframeContent=document.getElementById(""editor"").contentWindow;" & vbCrLf
    strJS = strJS & "   IframeContent.HtmlEdit.focus();" & vbCrLf
    strJS = strJS & "   IframeContent.HtmlEdit.document.execCommand('selectAll');" & vbCrLf
    strJS = strJS & "   IframeContent.HtmlEdit.document.execCommand('copy');" & vbCrLf
    strJS = strJS & "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    strJS = strJS & "  if (CurrentMode==0){" & vbCrLf
    strJS = strJS & "       document.myform.GuestContent.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  else if(CurrentMode==1){" & vbCrLf
    strJS = strJS & "       document.myform.GuestContent.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  if(document.myform.GuestContent.value==''){" & vbCrLf
    strJS = strJS & "    alert('内容不能为空！');" & vbCrLf
    strJS = strJS & "    editor.HtmlEdit.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  if(document.myform.GuestContent.value.length>65536){" & vbCrLf
    strJS = strJS & "    alert('内容不能超过64K！');" & vbCrLf
    strJS = strJS & "    editor.HtmlEdit.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    If EnableGuestBookCheck = True Then
        strJS = strJS & "  if(document.myform.CheckCode.value==''){" & vbCrLf
        strJS = strJS & "    alert('请输入您的验证码！');" & vbCrLf
        strJS = strJS & "    document.myform.CheckCode.focus();" & vbCrLf
        strJS = strJS & "    return(false);" & vbCrLf
        strJS = strJS & "  }" & vbCrLf
    End If
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "</script>" & vbCrLf
    ShowJS_Guest = strJS
End Function

'**************************************************
'函数名：KeywordReplace
'作  用：标示搜索关键字
'参  数：strChar-----要转换的字符
'返回值：转换后的字符
'**************************************************
Private Function KeywordReplace(strChar)
    If strChar = "" Then
        KeywordReplace = ""
    Else
        KeywordReplace = PE_Replace(strChar, "" & Keyword & "", "<font class=Channel_font>" & Keyword & "</font>")
    End If
    If IsNull(KeywordReplace) Then KeywordReplace = ""
End Function


'=================================================
'函数名：Guest_info()
'作  用：留言操作信息
'参  数：info 提示信息内容
'=================================================
Private Function Guest_info(info)
    Dim strInfo
    'strInfo = Replace(Replace(XmlText("Guest", "Guest_info", "<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr align='center'><td class='Guest_title_760'>留言操作反馈信息</td></tr><tr><td class='main_tdbg_575'><table cellpadding=5 cellspacing=0 border=0 width=100% align=center><tr><td height='100' valign='top'>{$info}</td></tr><tr align='center' class='tdbg'><td><a href='{$FileName}'>【查看留言】</a><a href='Guest_Write.asp'>【签写留言】</a></td></tr></table></td></tr></table><br>"), "{$info}", info), "{$FileName}", FileName)
    strInfo = Replace(Replace(XmlText("Guest", "Guest_info", "<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr align='center'><td class='Guest_title_760'>留言操作反馈信息</td></tr><tr><td class='main_tdbg_575'><table cellpadding=5 cellspacing=0 border=0 width=100% align=center><tr><td height='100' valign='top'>{$info}</td></tr><tr align='center' class='tdbg'><td><a href='{$FileName}'>【查看留言】</a><a href='javascript:history.go(-1)'>【签写留言】</a></td></tr></table></td></tr></table><br>"), "{$info}", info), "{$FileName}", FileName)
    Guest_info = strInfo
End Function

'=================================================
'过程名：GetGKind_Option()
'作  用：下拉框留言类别
'参  数：ShowType 显示类型
'        KindID   类别
'=================================================
Private Function GetGKind_Option(ShowType, KindID)
    Dim sqlGKind, rsGKind, strOption
    If ShowType = 3 Then
        strOption = ""
    Else
        strOption = "<option value='0'"
        If KindID = 0 Then
            strOption = strOption & " selected"
        End If
        strOption = strOption & ">不指定类别</option>"
    End If
    sqlGKind = "select * from PE_Guestkind order by OrderID"
    Set rsGKind = Conn.Execute(sqlGKind)
    Do While Not rsGKind.EOF
        If rsGKind("KindID") = KindID Then
            strOption = strOption & "<option value='" & rsGKind("KindID") & "' selected>" & rsGKind("KindName") & "</option>"
        Else
            strOption = strOption & "<option value='" & rsGKind("KindID") & "'>" & rsGKind("KindName") & "</option>"
        End If
        rsGKind.MoveNext
    Loop
    rsGKind.Close
    Set rsGKind = Nothing
    GetGKind_Option = strOption
End Function
'=================================================
'过程名：GetGKindList()
'作  用：横向显示留言类别
'参  数：无
'=================================================
Private Function GetGKindList()
    Dim rsGKind, sqlGKind, strGKind, i
    sqlGKind = "select * from PE_Guestkind order by OrderID"
    Set rsGKind = Conn.Execute(sqlGKind)
    If rsGKind.BOF And rsGKind.EOF Then
        strGKind = "| " & XmlText("Guest", "KindList/Nofound", "没有任何类别")
    Else
        i = 1
        strGKind = "| "
        Do While Not rsGKind.EOF
            strGKind = strGKind & "<a href='index.asp?KindID=" & rsGKind("KindID") & "'>" & rsGKind("KindName") & "</a>"
            strGKind = strGKind & " | "
            i = i + 1
            If i Mod 10 = 0 Then
                strGKind = strGKind & "<br>"
            End If
            rsGKind.MoveNext
        Loop
    End If
    rsGKind.Close
    Set rsGKind = Nothing
    'If GuestBook_IsAssignSort = False Then
        'strGKind = strGKind & "<a href='index.asp?KindID=0'>" & XmlText("xxxxx", "xxxxxx", "不属任何类别") & "</a> |"
    'End If
    GetGKindList = strGKind
End Function

'=================================================
'函数名：ShowGueststyle()
'作  用：获取查看方式
'参  数：无
'=================================================
Private Function GuestStyle()
    ShowGStyle = Request.Cookies("ShowGStyle")
    If ShowGStyle = "" Or Not IsNumeric(ShowGStyle) Then
        ShowGStyle = 1
    Else
        ShowGStyle = Int(ShowGStyle)
    End If
    GuestStyle = ShowGStyle
End Function
'=================================================
'函数名：ShowGueststyle()
'作  用：显示切换方式
'参  数：无
'=================================================
Private Function ShowGueststyle()
    Dim Shtm
    If ShowGStyle = 1 Then
        Shtm = "<a class=Guest href=ShowGuestStyle.asp?ShowGStyle=2>" & XmlText("Guest", "ShowGueststyle/Mode1", "切换到留言本方式") & "</a>"
    Else
        Shtm = "<a class=Guest href=ShowGuestStyle.asp?ShowGStyle=1>" & XmlText("Guest", "ShowGueststyle/Mode2", "切换到讨论区方式") & "</a>"
    End If
    ShowGueststyle = Shtm
End Function
'=================================================
'函数名：TransformTime()
'作  用：格式化时间
'参  数：时间
'=================================================
Private Function TransformTime(GuestDatetime)
    If Not IsDate(GuestDatetime) Then Exit Function
    Dim thour, tminute, tday, nowday, dnt, dayshow, pshow
    thour = Hour(GuestDatetime)
    tminute = Minute(GuestDatetime)
    tday = DateValue(GuestDatetime)
    nowday = DateValue(Now)
    If thour < 10 Then
        thour = "0" & thour
    End If
    If tminute < 10 Then
        tminute = "0" & tminute
    End If
    dnt = DateDiff("d", tday, nowday)
    If dnt > 2 Then
       dayshow = Year(GuestDatetime)
       If (Month(GuestDatetime) < 10) Then
           dayshow = dayshow & "-0" & Month(GuestDatetime)
       Else
           dayshow = dayshow & "-" & Month(GuestDatetime)
       End If
       If (Day(GuestDatetime) < 10) Then
           dayshow = dayshow & "-0" & Day(GuestDatetime)
       Else
           dayshow = dayshow & "-" & Day(GuestDatetime)
       End If
       TransformTime = dayshow
       Exit Function
    ElseIf dnt = 0 Then
       dayshow = XmlText("Guest", "TransformTime/d1", "今天 ")
    ElseIf dnt = 1 Then
       dayshow = XmlText("Guest", "TransformTime/d2", "昨天 ")
    ElseIf dnt = 2 Then
       dayshow = XmlText("Guest", "TransformTime/d3", "前天 ")
    End If
    TransformTime = dayshow & pshow & thour & ":" & tminute
End Function

'=================================================
'函数名：TransformIP()
'作  用：格式化IP
'参  数：IP
'=================================================
Private Function TransformIP(GuestIP)
    Dim arrIp
    arrIp = Split(GuestIP, ".")
    If UBound(arrIp) > 0 Then
        TransformIP = arrIp(0) & "." & arrIp(1) & ".*"
    Else
        TransformIP = "*"
    End If
End Function

'=================================================
'函数名：ShowTip()
'作  用：鼠标经过显示提示
'参  数：无
'=================================================
Private Function ShowTip()
    Dim strTip
    strTip = "<div id=toolTipLayer style='position: absolute; visibility: hidden'></div>" & vbCrLf
    strTip = strTip & "<SCRIPT language=JavaScript>" & vbCrLf
    strTip = strTip & "var ns4 = document.layers;" & vbCrLf
    strTip = strTip & "var ns6 = document.getElementById && !document.all;" & vbCrLf
    strTip = strTip & "var ie4 = document.all;" & vbCrLf
    strTip = strTip & "offsetX = 0;" & vbCrLf
    strTip = strTip & "offsetY = 20;" & vbCrLf
    strTip = strTip & "var toolTipSTYLE='';" & vbCrLf
    strTip = strTip & "function initToolTips()" & vbCrLf
    strTip = strTip & "{" & vbCrLf
    strTip = strTip & "  if(ns4||ns6||ie4)" & vbCrLf
    strTip = strTip & "  {" & vbCrLf
    strTip = strTip & "    if(ns4) toolTipSTYLE = document.toolTipLayer;" & vbCrLf
    strTip = strTip & "    else if(ns6) toolTipSTYLE = document.getElementById('toolTipLayer').style;" & vbCrLf
    strTip = strTip & "    else if(ie4) toolTipSTYLE = document.all.toolTipLayer.style;" & vbCrLf
    strTip = strTip & "    if(ns4) document.captureEvents(Event.MOUSEMOVE);" & vbCrLf
    strTip = strTip & "    else" & vbCrLf
    strTip = strTip & "    {" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.visibility = 'visible';" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.display = 'none';" & vbCrLf
    strTip = strTip & "    }" & vbCrLf
    strTip = strTip & "    document.onmousemove = moveToMouseLoc;" & vbCrLf
    strTip = strTip & "  }" & vbCrLf
    strTip = strTip & "}" & vbCrLf
    strTip = strTip & "function toolTip(msg, fg, bg)" & vbCrLf
    strTip = strTip & "{" & vbCrLf
    strTip = strTip & "  if(toolTip.arguments.length < 1)" & vbCrLf
    strTip = strTip & "  {" & vbCrLf
    strTip = strTip & "    if(ns4) toolTipSTYLE.visibility = 'hidden';" & vbCrLf
    strTip = strTip & "    else toolTipSTYLE.display = 'none';" & vbCrLf
    strTip = strTip & "  }" & vbCrLf
    strTip = strTip & "  else" & vbCrLf
    strTip = strTip & "  {" & vbCrLf
    strTip = strTip & "    if(!fg) fg = '#333333';" & vbCrLf
    strTip = strTip & "    if(!bg) bg = '#FFFFFF';" & vbCrLf
    strTip = strTip & "    var content = '<table border=""0"" cellspacing=""0"" cellpadding=""1"" bgcolor=""' + fg + '""><td>' + '<table border=""0"" cellspacing=""0"" cellpadding=""1"" bgcolor=""' + bg + '""><td align=""left"" nowrap style=""line-height: 120%""><font color=""' + fg + '"">' + msg + '&nbsp\;</font></td></table></td></table>';" & vbCrLf
    strTip = strTip & "    if(ns4)" & vbCrLf
    strTip = strTip & "    {" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.document.write(content);" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.document.close();" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.visibility = 'visible';" & vbCrLf
    strTip = strTip & "    }" & vbCrLf
    strTip = strTip & "    if(ns6)" & vbCrLf
    strTip = strTip & "    {" & vbCrLf
    strTip = strTip & "      document.getElementById('toolTipLayer').innerHTML = content;" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.display='block'" & vbCrLf
    strTip = strTip & "    }" & vbCrLf
    strTip = strTip & "    if(ie4)" & vbCrLf
    strTip = strTip & "    {" & vbCrLf
    strTip = strTip & "      document.all('toolTipLayer').innerHTML=content;" & vbCrLf
    strTip = strTip & "      toolTipSTYLE.display='block'" & vbCrLf
    strTip = strTip & "    }" & vbCrLf
    strTip = strTip & "  }" & vbCrLf
    strTip = strTip & "}" & vbCrLf
    strTip = strTip & "function moveToMouseLoc(e)" & vbCrLf
    strTip = strTip & "{" & vbCrLf
    strTip = strTip & "  if(ns4||ns6)" & vbCrLf
    strTip = strTip & "  {" & vbCrLf
    strTip = strTip & "    x = e.pageX;" & vbCrLf
    strTip = strTip & "    y = e.pageY;" & vbCrLf
    strTip = strTip & "  }" & vbCrLf
    strTip = strTip & "  else" & vbCrLf
    strTip = strTip & "  {" & vbCrLf
    strTip = strTip & "    x = event.x + document.body.scrollLeft;" & vbCrLf
    strTip = strTip & "    y = event.y + document.body.scrollTop;" & vbCrLf
    strTip = strTip & "  }" & vbCrLf
    strTip = strTip & "  toolTipSTYLE.left = x + offsetX;" & vbCrLf
    strTip = strTip & "  toolTipSTYLE.top = y + offsetY;" & vbCrLf
    strTip = strTip & "  return true;" & vbCrLf
    strTip = strTip & "}" & vbCrLf
    strTip = strTip & "initToolTips();" & vbCrLf
    strTip = strTip & "</SCRIPT>" & vbCrLf
    ShowTip = strTip
End Function


'=================================================
'函数名：GetRepeatGuestBook()
'作  用：留言本方式时替换要循环的标签
'参  数：要替换的值
'=================================================
Private Function GetRepeatGuestBook(strParameter, strList)
    Dim strTemp, arrTemp
    Dim strSoftPic, strPicTemp, arrPicTemp

    If strParameter = "" Then
        GetRepeatGuestBook = ""
        Exit Function
    End If
    
    arrTemp = PE_CLng(strParameter)
    
    If arrTemp = 0 Then
        Select Case Action
            Case "user"
                Call ShowAllGuest(3)
            Case "Quintessence"
                Call ShowAllGuest(4)
            Case Else
                If UserLogined = True Then
                    Call ShowAllGuest(1)
                Else
                    Call ShowAllGuest(2)
                End If
        End Select
    ElseIf arrTemp = 1 Then
        Call ShowAllGuest(5)
    ElseIf arrTemp = 2 Then
        Call ShowAllGuest(6)
    ElseIf arrTemp = 3 Then
        If UserLogined = True Then
            ShowAllGuest (1)
        Else
            ShowAllGuest (2)
        End If
    End If
    
    If FoundErr = True Then
        GetRepeatGuestBook = Guest_info(ErrMsg)
        Exit Function
    End If


    Dim UserGuestName, UserType, UserSex, UserEmail, UserHomepage, UserOicq, UserIcq, UserMsn
    Dim GuestNum, GuestTip, TipName, TipSex, TipEmail, TipOicq, TipHomepage
    Dim GtbDel, GtbNoEnter, GtbMan, GtbGirl, GtbTip1, GtbTip2, Gtbp1, Gtbp2, Gtbp3, Gtbp4, Gtbp5, Gtbp6, GtbGuestImages, GtbUser, GtbGuest
    Dim GtbHide1, GtbHide2, GtbHide3, GtbReply4, GtbReply5, GtbReply6, GtbReply7, GtbReply8, GtbReply9, GtbReply10, GtbReply11
    GtbDel = XmlText("Guest", "GuestBookShow/Del", "（已删除）")
    GtbNoEnter = XmlText("BaseText", "NoEnter", "未填")
    GtbMan = XmlText("BaseText", "Man", "男")
    GtbGirl = XmlText("BaseText", "Girl", "女")
    GtbTip1 = XmlText("Guest", "GuestBookShow/Tip1", " 姓名：{$Name} {$Sex}<br> 主页：{$Homepage}<br> OICQ：{$Oicq}<br> 信箱：{$Email}<br> 地址：{$GuestIP}<br> 时间：{$Time}")
    GtbTip2 = XmlText("Guest", "GuestBookShow/Tip2", "用户相关资料保密。")
    Gtbp1 = XmlText("Guest", "GuestBookShow/p1", "固顶留言")
    Gtbp2 = XmlText("Guest", "GuestBookShow/p2", "精华留言")
    Gtbp3 = XmlText("Guest", "GuestBookShow/p3", "有回复")
    Gtbp4 = XmlText("Guest", "GuestBookShow/p4", "无回复")
    Gtbp5 = XmlText("Guest", "GuestBookShow/p5", "回复：")
    Gtbp6 = XmlText("Guest", "GuestBookShow/p6", "主题：")
    GtbGuestImages = XmlText("Guest", "GuestBookShow/GuestImages", "<img src='{$GuestImages}.gif' width='80' height='90' onMouseOut=toolTip() onMouseOver=""toolTip('{$GuestTip}')"">")
    GtbUser = XmlText("Guest", "GuestBookShow/User", "用户")
    GtbGuest = XmlText("Guest", "GuestBookShow/Guest", "游客")
    GtbHide1 = XmlText("Guest", "GuestBookShow/Hide1", " **************************************<br> * 隐藏留言，管理员和留言用户可以看到 *<br> **************************************")
    GtbHide2 = XmlText("Guest", "GuestBookShow/Hide2", "[隐藏]")
    GtbHide3 = XmlText("Guest", "GuestBookShow/Hide3", " *********************************************<br> * 隐藏管理员回复，管理员和留言用户可以看到 *<br> *********************************************")
    GtbReply4 = XmlText("Guest", "GuestBookShow/Reply4", "回复这条留言")
    GtbReply5 = XmlText("Guest", "GuestBookShow/Reply5", "编辑这条留言")
    GtbReply6 = XmlText("Guest", "GuestBookShow/Reply6", "确定要删除此留言吗？")
    GtbReply7 = XmlText("Guest", "GuestBookShow/Reply7", "删除这条留言")
    GtbReply8 = XmlText("Guest", "GuestBookShow/Reply8", "查看全部回复")
    GtbReply9 = XmlText("Guest", "GuestBookShow/Reply9", "共有回复{$ReplyNum}条")
    GtbReply10 = XmlText("Guest", "GuestBookShow/Reply10", "回复这条留言")
    GtbReply11 = XmlText("Guest", "GuestBookShow/Reply11", "返回列表")

    GuestNum = 0
    Do While Not rsGuest.EOF
        UserGuestName = rsGuest("GuestName")
        UserSex = rsGuest("GuestSex")
        UserEmail = rsGuest("GuestEmail")
        UserOicq = rsGuest("GuestOicq")
        UserIcq = rsGuest("GuestIcq")
        UserMsn = rsGuest("GuestMsn")
        UserHomepage = rsGuest("GuestHomepage")
        TipName = UserGuestName
        If UserEmail = "" Or IsNull(UserEmail) Then
            TipEmail = GtbNoEnter
        Else
            TipEmail = UserEmail
        End If
        If UserOicq = "" Or IsNull(UserOicq) Then
            TipOicq = GtbNoEnter
        Else
            TipOicq = UserOicq
        End If
        If UserHomepage = "" Or IsNull(UserHomepage) Then
            TipHomepage = GtbNoEnter
        Else
            TipHomepage = UserHomepage
        End If
        If UserIcq = "" Or IsNull(UserIcq) Then UserIcq = GtbNoEnter
        If UserMsn = "" Or IsNull(UserMsn) Then UserMsn = GtbNoEnter
        If UserSex = "1" Then
            TipSex = "(" & GtbMan & ")"
        ElseIf UserSex = "0" Then
            TipSex = "(" & GtbGirl & ")"
        Else
            TipSex = ""
        End If
        If GuestBook_ShowIP = True Then
            GuestTip = Replace(Replace(Replace(Replace(Replace(Replace(Replace(GtbTip1, "{$Name}", TipName), "{$Sex}", TipSex), "{$Homepage}", TipHomepage), "{$Oicq}", TipOicq), "{$Email}", TipEmail), "{$GuestIP}", rsGuest("GuestIP")), "{$Time}", rsGuest("GuestDatetime"))
        Else
            GuestTip = Replace(Replace(Replace(Replace(Replace(Replace(Replace(GtbTip1, "{$Name}", TipName), "{$Sex}", TipSex), "{$Homepage}", TipHomepage), "{$Oicq}", TipOicq), "{$Email}", TipEmail), "{$GuestIP}", TransformIP(rsGuest("GuestIP"))), "{$Time}", rsGuest("GuestDatetime"))
        End If
        
        strTemp = strList
        
        If rsGuest("OnTop") = 1 Then
            strTemp = Replace(strTemp, "{$IsTitlePic}", "<img border='0' src='" & GImagePath & "ontop.gif' title=" & Gtbp1 & ">")
        ElseIf rsGuest("Quintessence") = 1 Then
            strTemp = Replace(strTemp, "{$IsTitlePic}", "<img border='0' src='" & GImagePath & "pith.gif' title=" & Gtbp2 & ">")
        ElseIf rsGuest("ReplyNum") > 0 Then
            strTemp = Replace(strTemp, "{$IsTitlePic}", "<img border='0' src='" & GImagePath & "yes.gif' title=" & Gtbp3 & ">")
        Else
            strTemp = Replace(strTemp, "{$IsTitlePic}", "<img border='0' src='" & GImagePath & "no.gif' title=" & Gtbp4 & ">")
        End If
        
        If ReplyId = "" Then
            If strField <> "" And rsGuest("GuestID") <> rsGuest("TopicID") Then
                strTemp = Replace(strTemp, "{$GuestType}", Gtbp5)
                strTemp = Replace(strTemp, "{$GuestTitle}", "<a class='Guest' href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & ">" & KeywordReplace(rsGuest("GuestTitle")) & "</a>")
            Else
                If Action = "edit" And rsGuest("GuestID") <> rsGuest("TopicID") Then
                    strTemp = Replace(strTemp, "{$GuestType}", Gtbp5)
                    strTemp = Replace(strTemp, "{$GuestTitle}", "<a class='Guest' href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & ">" & KeywordReplace(rsGuest("GuestTitle")) & "</a>")
                Else
                    strTemp = Replace(strTemp, "{$GuestType}", Gtbp6)
                    strTemp = Replace(strTemp, "{$GuestTitle}", "<a class='Guest' href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & ">" & KeywordReplace(rsGuest("GuestTitle")) & "</a>")
                End If
            End If
        Else
            If rsGuest("GuestID") = rsGuest("TopicID") Then
                strTemp = Replace(strTemp, "{$GuestType}", Gtbp6)
                strTemp = Replace(strTemp, "{$GuestTitle}", "<a class='Guest' href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & ">" & KeywordReplace(rsGuest("GuestTitle")) & "</a>")
            Else
                strTemp = Replace(strTemp, "{$GuestType}", Gtbp5)
                strTemp = Replace(strTemp, "{$GuestTitle}", "<a class='Guest' href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & ">" & KeywordReplace(rsGuest("GuestTitle")) & "</a>")
            End If
        End If

    
        strTemp = Replace(strTemp, "{$GuestTime}", rsGuest("GuestDatetime"))

        strTemp = Replace(strTemp, "{$GuestHead}", Replace(Replace(GtbGuestImages, "{$GuestImages}", GFacePath & rsGuest("GuestImages")), "{$GuestTip}", GuestTip))
        'strTemp = Replace(strTemp, "{$GuestHead}", "                        <img src='" & GFacePath & rsGuest("GuestImages") & ".gif' width='80' height='90' onMouseOut=toolTip() onMouseOver=""toolTip('" & GuestTip & "')"">")
        If rsGuest("GuestType") = 1 Then
            strTemp = Replace(strTemp, "{$GuestNameType}", GtbUser)
        Else
            strTemp = Replace(strTemp, "{$GuestNameType}", GtbGuest)
        End If
        strTemp = Replace(strTemp, "{$GuestName}", KeywordReplace(UserGuestName))

        strTemp = Replace(strTemp, "{$GuestFaceShow}", "<img src='" & GImagePath & "face" & rsGuest("GuestFace") & ".gif' width='19' height='19'>")
        
        Dim ContentShow, AdminReplyShow, LastReplyShow
        '替换留言内容
        regEx.Pattern = "【ContentShow】([\s\S]*?)【\/ContentShow】"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            ContentShow = Match.value
        Next
        If rsGuest("GuestIsPrivate") = True And rsGuest("GuestName") <> UserName And rsGuest("ReplyIsPrivate") = True Then
            strTemp = Replace(strTemp, ContentShow, "<br><br><font class=Guest_font>" & GtbHide1 & "</font>")
        End If
        strTemp = Replace(strTemp, "【ContentShow】", "")
        strTemp = Replace(strTemp, "【/ContentShow】", "")

        If rsGuest("GuestIsPrivate") = True And rsGuest("GuestName") <> UserName Then
            strTemp = Replace(strTemp, "{$IsHiddenShow}", "                        <font class=Guest_font>" & GtbHide2 & "</font>&nbsp;")
            strTemp = Replace(strTemp, "{$GuestContentShow}", "")
        Else
            strTemp = Replace(strTemp, "{$IsHiddenShow}", "")
        End If
        strTemp = Replace(strTemp, "{$GuestContentShow}", KeywordReplace(FilterJS(rsGuest("GuestContent"))))
        
   
        '替换用户最后回复
        regEx.Pattern = "【LastReplyShow】([\s\S]*?)【\/LastReplyShow】"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            LastReplyShow = Match.value
        Next

        If rsGuest("LastReplyGuest") = "" Or IsNull(rsGuest("LastReplyGuest")) Or ReplyId <> "" Then
            strTemp = Replace(strTemp, LastReplyShow, "")
        End If

        strTemp = Replace(strTemp, "【LastReplyShow】", "")
        strTemp = Replace(strTemp, "【/LastReplyShow】", "")
        strTemp = Replace(strTemp, "{$LastReplyContent}", KeywordReplace(rsGuest("LastReplyContent")))
        strTemp = Replace(strTemp, "{$LastReplyGuest}", KeywordReplace(rsGuest("LastReplyGuest")))
        strTemp = Replace(strTemp, "{$LastReplyTitle}", KeywordReplace(rsGuest("LastReplyTitle")))
        strTemp = Replace(strTemp, "{$LastReplyTime}", KeywordReplace(rsGuest("LastReplyTime")))
        '替换用户最后回复完毕
        
        '替换管理员回复

        regEx.Pattern = "【AdminReplyShow】([\s\S]*?)【\/AdminReplyShow】"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            AdminReplyShow = Match.value
        Next
        
        If rsGuest("GuestReply") = "" Or IsNull(rsGuest("GuestReply")) Then
            strTemp = Replace(strTemp, AdminReplyShow, "")
        ElseIf rsGuest("ReplyIsPrivate") = True And rsGuest("GuestName") <> UserName Then
            strTemp = Replace(strTemp, AdminReplyShow, "<font class=Guest_font>" & GtbHide3 & "</font>")
        Else
                        strTemp = Replace(strTemp, "【AdminReplyShow】", "")
                        strTemp = Replace(strTemp, "【/AdminReplyShow】", "")
                        strTemp = Replace(strTemp, "{$ReplyAdmin}", KeywordReplace(rsGuest("GuestReplyAdmin")))
                        strTemp = Replace(strTemp, "{$AdminReplyTime}", KeywordReplace(rsGuest("GuestReplyDatetime")))
                        strTemp = Replace(strTemp, "{$AdminReplyContent}", KeywordReplace(rsGuest("GuestReply")))
        End If
        '替换管理员回复完毕
        
        '替换留言内容完毕


        If UserHomepage = "" Or IsNull(UserHomepage) Then
            strTemp = Replace(strTemp, "{$HomePagePic}", "<img src=" & GImagePath & "nourl.gif width=45 height=16 border=0>")
        Else
            strTemp = Replace(strTemp, "{$HomePagePic}", "<a href=" & UserHomepage & " target=""_blank""><img src=" & GImagePath & "url.gif width=45 height=16 alt=" & UserHomepage & " border=0></a>")
        End If
        If UserOicq = "" Or IsNull(UserOicq) Then
            strTemp = Replace(strTemp, "{$OicqPic}", "<img src=" & GImagePath & "nooicq.gif width=45 height=16 border=0>")
        Else
            strTemp = Replace(strTemp, "{$OicqPic}", "<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & UserOicq & " target='_blank'><img src=" & GImagePath & "oicq.gif width=45 height=16 alt=" & UserOicq & " border=0 ></a>")
        End If
        If UserEmail = "" Or IsNull(UserEmail) Then
            strTemp = Replace(strTemp, "{$EmailPic}", "<img src=" & GImagePath & "noemail.gif width=45 height=16 border=0>")
        Else
            strTemp = Replace(strTemp, "{$EmailPic}", "<a href=mailto:" & UserEmail & "><img src=" & GImagePath & "email.gif width=45 height=16 border=0 alt=" & UserEmail & "></a>")
        End If
        If GuestBook_ShowIP = True Then
            strTemp = Replace(strTemp, "{$OtherPic}", "<img src=" & GImagePath & "other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;Icq：" & UserIcq & "<br>&nbsp;Msn：" & UserMsn & "<br>&nbsp;I P：" & rsGuest("GuestIP") & "')"">")
        Else
            strTemp = Replace(strTemp, "{$OtherPic}", "<img src=" & GImagePath & "other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;Icq：" & UserIcq & "<br>&nbsp;Msn：" & UserMsn & "<br>&nbsp;I P：" & TransformIP(rsGuest("GuestIP")) & "')"">")
        End If
        If rsGuest("GuestIsPassed") = False Then
             strTemp = Replace(strTemp, "{$ReplyPic}", "")
        End If
        If ReplyId = "" And rsGuest("GuestID") = rsGuest("TopicID") Then

            strTemp = Replace(strTemp, "{$ReplyPic}", "<a href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & "><img src=" & GImagePath & "reply.gif width=45 height=16 border=0 alt=" & GtbReply4 & "></a>")
        Else
            strTemp = Replace(strTemp, "{$ReplyPic}", "")
        End If
        
        If rsGuest("GuestName") = UserName And rsGuest("GuestIsPassed") = False Then

            strTemp = Replace(strTemp, "{$EditPic}", "<a href=Guest_Write.asp?action=edit&guestid=" & rsGuest("guestid") & "><img src=" & GImagePath & "edit.gif width=45 height=16 border=0 alt=" & GtbReply5 & "></a>")

            strTemp = Replace(strTemp, "{$DelPic}", "<a href=" & FileName & "?action=del&guestid=" & rsGuest("guestid") & " onClick=""return confirm('" & GtbReply6 & "');""><img src=" & GImagePath & "del.gif width=45 height=16  alt=" & GtbReply7 & " border=0></a></td>")
        Else
            strTemp = Replace(strTemp, "{$EditPic}", "")
            strTemp = Replace(strTemp, "{$DelPic}", "")
        End If

        If rsGuest("GuestIsPassed") = False Then
            strTemp = Replace(strTemp, "{$InfoShow}", "")
        ElseIf rsGuest("ReplyNum") > 0 Then
            If ReplyId = "" Then
                strTemp = Replace(strTemp, "{$InfoShow}", "<a href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & "><img src=" & GImagePath & "reply0.gif width=15 height=16 border=0>&nbsp;" & GtbReply8 & "</a>(共" & rsGuest("ReplyNum") & "条)")
            Else
                strTemp = Replace(strTemp, "{$InfoShow}", Replace(GtbReply9, "{$ReplyNum}", rsGuest("ReplyNum")))
            End If
        Else
            If ReplyId = "" Then
                If rsGuest("GuestID") = rsGuest("TopicID") Then
                    strTemp = Replace(strTemp, "{$InfoShow}", "<a href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & "><img src=" & GImagePath & "reply0.gif width=15 height=16 border=0>&nbsp;" & GtbReply10 & "</a>")
                Else
                    strTemp = Replace(strTemp, "{$InfoShow}", "")
                End If
            Else
                strTemp = Replace(strTemp, "{$InfoShow}", "<a href='" & FileName & "'><img src=" & GImagePath & "home.gif width=15 height=16 border=0>&nbsp;" & GtbReply11 & "</a>")
            End If
        End If

                testHTML = testHTML & strTemp

        rsGuest.MoveNext
        GuestNum = GuestNum + 1
        If GuestNum >= MaxPerPage Then Exit Do
        
    Loop

    testHTML = testHTML & ShowTip()
    
    rsGuest.Close
    Set rsGuest = Nothing
    
    GetRepeatGuestBook = testHTML
End Function
'=================================================
'函数名：GetRepeatDiscussion()
'作  用：讨论区方式时替换要循环的标签
'参  数：要替换的值
'=================================================
Private Function GetRepeatDiscussion(strParameter, strList)
    Dim strTemp, arrTemp, strBeg, strEnd, strSource
    Dim strSoftPic, strPicTemp, arrPicTemp

    If strParameter = "" Or IsNull(strParameter) Then
		GetRepeatDiscussion = ""
        Exit Function
    End If
	strSource = strList

    '替换讨论区方式列表头部
    regEx.Pattern = "【GuestList2_Beg】([\s\S]*?)【\/GuestList2_Beg】"
    Set Matches = regEx.Execute(strSource)
    For Each Match In Matches
        strBeg = Match.SubMatches(0)
        strSource = Replace(strSource, Match.value, "")
    Next

    '替换讨论区方式列表尾部
    regEx.Pattern = "【GuestList2_End】([\s\S]*?)【\/GuestList2_End】"
    Set Matches = regEx.Execute(strSource)
    For Each Match In Matches
        strEnd = Match.SubMatches(0)
        strSource = Replace(strSource, Match.value, "")
    Next
    
    arrTemp = PE_CLng(strParameter)
    
    If arrTemp = 0 Then
        Select Case Action
		Case "user"
			Call ShowAllGuest(3)
		Case "Quintessence"
			Call ShowAllGuest(4)
		Case Else
			If UserLogined = True Then
				ShowAllGuest (1)
			Else
				ShowAllGuest (2)
			End If
        End Select
    ElseIf arrTemp = 1 Then
        If UserLogined = True Then
            ShowAllGuest (1)
        Else
            ShowAllGuest (2)
        End If
    End If
    
    If FoundErr = True Then
        GetRepeatDiscussion = Guest_info(ErrMsg)
        Exit Function
    End If

    Dim strHTM, strXml
    strXml = Split(XmlText("Guest", "discussionShow/Text", "发言主题|||所有主题|||留言人|||回复|||阅读|||最后回复|||固顶留言|||精华留言|||展开主题回复的列表|||无回复|||点击查看记录具体信息"), "|||")

    Dim i, GtbUser, GtbGuest
    GtbUser = XmlText("Guest", "GuestBookShow/User", "用户")
    GtbGuest = XmlText("Guest", "GuestBookShow/Guest", "游客")
    i = 0
    Do While Not rsGuest.EOF
        strTemp = strSource
        If rsGuest("OnTop") = 1 Then
            strTemp = Replace(strTemp, "{$GuestFaceShow}", "<img border='0' src='" & GImagePath & "ontop.gif' title=" & strXml(6) & ">")
        ElseIf rsGuest("Quintessence") = 1 Then
            strTemp = Replace(strTemp, "{$GuestFaceShow}", "<img border='0' src='" & GImagePath & "pith.gif' title=" & strXml(7) & ">")
        Else
            strTemp = Replace(strTemp, "{$GuestFaceShow}", "  <img src='" & GImagePath & "face" & rsGuest("GuestFace") & ".gif' width='19' height='19'>")
        End If

        If rsGuest("ReplyNum") > 0 Then
            strTemp = Replace(strTemp, "{$IsTitlePic}", "<span id='FollowImg" & rsGuest("GuestID") & "'><a href='ListingTree.asp?TopicID=" & rsGuest("GuestID") & "&Action=show' target='hiddeniframe' title='" & strXml(8) & "'><img border='0' src='" & GImagePath & "yes.gif'></a></span>")
        Else
            strTemp = Replace(strTemp, "{$IsTitlePic}", "<img border='0' src='" & GImagePath & "no.gif' title=" & strXml(9) & ">")
        End If

        If rsGuest("GuestType") = 1 Then
            strTemp = Replace(strTemp, "{$GuestNameType}", GtbUser)
        Else
            strTemp = Replace(strTemp, "{$GuestNameType}", GtbGuest)
        End If

        strTemp = Replace(strTemp, "{$GuestTitle}", "  <a href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & ">" & rsGuest("GuestTitle") & "</a>")
        strTemp = Replace(strTemp, "{$GuestContentLength}", rsGuest("GuestContentLength") & "")
        strTemp = Replace(strTemp, "{$GuestName}", rsGuest("GuestName") & "")
        strTemp = Replace(strTemp, "{$ReplyNum}", rsGuest("ReplyNum") & "")
        strTemp = Replace(strTemp, "{$Hits}", rsGuest("Hits") & "")
        If rsGuest("LastReplyTime") <> "" Then
            strTemp = Replace(strTemp, "{$GuestTime}", TransformTime(rsGuest("LastReplyTime")))
            strTemp = Replace(strTemp, "{$LastReplyGuest}", rsGuest("LastReplyGuest"))
        ElseIf rsGuest("GuestReplyDatetime") <> "" Then
            strTemp = Replace(strTemp, "{$GuestTime}", TransformTime(rsGuest("GuestReplyDatetime")))
            strTemp = Replace(strTemp, "{$LastReplyGuest}", rsGuest("GuestReplyAdmin"))
        Else
            strTemp = Replace(strTemp, "{$GuestTime}", TransformTime(rsGuest("GuestDatetime")))
            strTemp = Replace(strTemp, "{$LastReplyGuest}", rsGuest("GuestName"))
        End If
        strTemp = Replace(strTemp, "{$GuestID}", rsGuest("GuestID"))
        
        testHTML = testHTML & strTemp

        i = i + 1
        If i >= MaxPerPage Then Exit Do
        rsGuest.MoveNext
    Loop
    testHTML = strBeg & testHTML & strEnd & "<br><iframe with='0' height='0' src='' name='hiddeniframe'></iframe>"

   
    rsGuest.Close
    Set rsGuest = Nothing
    
    GetRepeatDiscussion = testHTML
 
End Function

'=================================================
'函数名：GuestFace()
'作  用：留言心情选择
'参  数：无
'=================================================
Private Function GuestFace()
    Dim i, strHTM
    'For i = 1 To 30
    For i = 1 To 20
        strHTM = strHTM & "<input type='radio' name='GuestFace' value='" & i & "'"
        If i = PE_CLng(WriteFace) Then strHTM = strHTM & " checked"
        strHTM = strHTM & " style='BORDER:0px;width:19;'>"
        strHTM = strHTM & "<img src='" & GImagePath & "face" & i & ".gif' width='19' height='19'>" & vbCrLf
        If i Mod 10 = 0 Then strHTM = strHTM & "<br>"
    Next

    GuestFace = strHTM
End Function

'=================================================
'函数名：ManageRubbishContent()
'作  用：屏蔽垃圾广告子函数
'参  数：无
'=================================================
Private Function ManageRubbishContent(ByVal GuestBook_ManageRubbish, ByVal GuestContent)
    Dim RubbishContent
    RubbishContent = False
    ManageRubbishContent = RubbishContent
    Dim i, obj
    If GuestBook_ManageRubbish = "" Or IsNull(GuestBook_ManageRubbish) Then
        RubbishContent = False
        Exit Function
    End If
    obj = Split(GuestBook_ManageRubbish, "$$$")
    If GuestContent = "" Then Exit Function
    For i = 0 To UBound(obj)
        If Trim(obj(i)) <> "" And InStr(GuestContent, Trim(obj(i))) > 0 Then
            RubbishContent = True
            If RubbishContent Then Exit For
        End If
    Next
    ManageRubbishContent = RubbishContent
End Function
'=================================================
'函数名：GuestContent()
'作  用：留言内容
'参  数：无
'=================================================
Private Function GuestContent()
    Dim strHTM
    strHTM = "<textarea name='GuestContent' id='GuestContent' style='display:none' >" & Server.HTMLEncode(FilterJS(WriteContent)) & "</textarea>" & vbCrLf
    strHTM = strHTM & "<iframe ID='editor' src='../editor.asp?ChannelID=4&ShowType=2&tContentid=GuestContent' frameborder='1' scrolling='no' width='480' height='280' ></iframe>" & vbCrLf
    GuestContent = strHTM
End Function

'=================================================
'函数名：DefaultTemplate()
'作  用：得到默认模板设置，与以前版本兼容
'参  数：strType 页面类别
'=================================================
Private Function DefaultTemplate(strType)
    Dim TemplateType, strTemplate
    TemplateType = Trim(strType)
    
    If TemplateType = "" Or IsNull(TemplateType) Then
        DefaultTemplate = ""
        Exit Function
    End If
    
    Select Case TemplateType
        Case "Index"
            strTemplate = TemplateDiscission("Index") & TemplateGuestBook("Index")
        Case "strWrite"
            strTemplate = TemplateGuestBook("strWrite")
        Case "Reply"
            strTemplate = TemplateGuestBook("Reply")
        Case "Search"
            strTemplate = TemplateDiscission("Search") & TemplateGuestBook("Search")
    End Select
    
    DefaultTemplate = strTemplate
End Function

'=================================================
'函数名：TemplateDiscission()
'作  用：得到讨论区方式默认模板设置
'参  数：strstlye1 页面类别
'=================================================
Private Function TemplateDiscission(strstlye1)
    Dim strTemplate
    strTemplate = ""
    If strstlye1 = "Index" Then
        strTemplate = strTemplate & "     <!--讨论区方式循环显示留言结束-->   【GuestList2(0)】" & vbCrLf
    Else
        strTemplate = strTemplate & "     <!--讨论区方式循环显示留言结束-->   【GuestList2(1)】" & vbCrLf
    End If
    strTemplate = strTemplate & "             【GuestList2_Beg】<table width='100%' class='Guest_border' border='0' cellspacing='1' cellpadding='0' align='center'>" & vbCrLf
    strTemplate = strTemplate & "      <tr class='Guest_title'>" & vbCrLf
    strTemplate = strTemplate & "        <td width='58%' colspan='3'> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>发言主题</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>" & vbCrLf
    strTemplate = strTemplate & "        <td width='10%' nowrap> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>留言人</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>" & vbCrLf
    strTemplate = strTemplate & "        <td width='5%' nowrap> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>回复</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>   " & vbCrLf
    strTemplate = strTemplate & "        <td width='5%' nowrap> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>阅读</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>   " & vbCrLf
    strTemplate = strTemplate & "        <td width='22%' nowrap>  " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>最后回复</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>    " & vbCrLf
    strTemplate = strTemplate & "      </tr>【/GuestList2_Beg】" & vbCrLf
    strTemplate = strTemplate & "<tr class='Guest_tdbg'>" & vbCrLf
    strTemplate = strTemplate & "<td width='5%' align='center'>" & vbCrLf
    strTemplate = strTemplate & "  {$GuestFaceShow}" & vbCrLf
    strTemplate = strTemplate & "  </td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='5%'  align='center'>" & vbCrLf
    strTemplate = strTemplate & "{$IsTitlePic}" & vbCrLf
    strTemplate = strTemplate & "  </td><td width='48%' title='点击查看记录具体信息' align='left'>{$GuestTitle}<I><font color=gray>({$GuestContentLength}字)</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='10%' align='center'>{$GuestName}</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='5%' align='center'>{$ReplyNum}</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='5%' align='center'>{$Hits}</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='22%' align='left'>{$GuestTime}<font class=Channel_font> | </font>{$LastReplyGuest}</td></tr>" & vbCrLf
    strTemplate = strTemplate & "  <tr id='FollowTr{$GuestID}' style='display:none;'><td id='FollowTd{$GuestID}' colspan='7'></td></tr>" & vbCrLf
    strTemplate = strTemplate & "【GuestList2_End】</table>【/GuestList2_End】【/GuestList2】" & vbCrLf
    strTemplate = strTemplate & "     <!--讨论区方式循环显示留言结束-->" & vbCrLf
    TemplateDiscission = strTemplate
End Function

'=================================================
'函数名：TemplateGuestBook()
'作  用：得到留言本方式默认模板设置
'参  数：strstlye2 页面类别
'=================================================
Private Function TemplateGuestBook(strstlye2)
    Dim strTemplate
    strTemplate = ""
    strTemplate = strTemplate & "                     <!--留言本方式循环显示留言开始-->" & vbCrLf
    If strstlye2 = "Index" Then
        strTemplate = strTemplate & "     【GuestList1(0)】" & vbCrLf
    ElseIf strstlye2 = "strWrite" Then
        strTemplate = strTemplate & "     【GuestList1(1)】" & vbCrLf
    ElseIf strstlye2 = "Reply" Then
        strTemplate = strTemplate & "     【GuestList1(2)】" & vbCrLf
    Else
        strTemplate = strTemplate & "     【GuestList1(3)】" & vbCrLf
    End If
    strTemplate = strTemplate & "          <table width='100%' border='0' cellpadding='0' cellspacing='1' class='Guest_border'>" & vbCrLf
    strTemplate = strTemplate & "        <tr>" & vbCrLf
    strTemplate = strTemplate & "          <td align='center' valign='top'>" & vbCrLf
    strTemplate = strTemplate & "            <table width='100%' border='0' cellspacing='0' cellpadding='0' class='Guest_title'>" & vbCrLf
    strTemplate = strTemplate & "              <tr>" & vbCrLf
    strTemplate = strTemplate & "                <td>" & vbCrLf
    strTemplate = strTemplate & "{$IsTitlePic}<strong>{$GuestType}：</strong>{$GuestTitle}" & vbCrLf
    strTemplate = strTemplate & "                </td>" & vbCrLf
    strTemplate = strTemplate & "                <td width='180'>" & vbCrLf
    strTemplate = strTemplate & "                  <img src='{$InstallDir}Images/posttime.gif' width='11' height='11' align='absmiddle'>：{$GuestTime}" & vbCrLf
    strTemplate = strTemplate & "                </td>" & vbCrLf
    strTemplate = strTemplate & "              </tr>" & vbCrLf
    strTemplate = strTemplate & "            </table>" & vbCrLf
    strTemplate = strTemplate & "          </td>" & vbCrLf
    strTemplate = strTemplate & "        </tr>" & vbCrLf
    strTemplate = strTemplate & "        <tr>" & vbCrLf
    strTemplate = strTemplate & "          <td align='center' height='153' valign='top' class='Guest_tdbg'>" & vbCrLf
    strTemplate = strTemplate & "            <table width='100%' border='0' cellpadding='0' cellspacing='3'>" & vbCrLf
    strTemplate = strTemplate & "              <tr>" & vbCrLf
    strTemplate = strTemplate & "                <td width='130' align='center' height='130' valign='top'>" & vbCrLf
    strTemplate = strTemplate & "{$GuestHead}<br>" & vbCrLf
    strTemplate = strTemplate & "                        <br>" & vbCrLf
    strTemplate = strTemplate & "【{$GuestNameType}】<br>{$GuestName}                </td>" & vbCrLf
    strTemplate = strTemplate & "                <td align='center' height='153' width='1' class='Guest_tdbg_1px'></td>" & vbCrLf
    strTemplate = strTemplate & "                <td>" & vbCrLf
    strTemplate = strTemplate & "                  <table width='100%' border='0' cellpadding='6' cellspacing='0' height='125' style='TABLE-LAYOUT: fixed'>" & vbCrLf
    strTemplate = strTemplate & "                    <tr>" & vbCrLf
    strTemplate = strTemplate & "                      <td align='left' valign='top'>{$GuestFaceShow}【ContentShow】" & vbCrLf
    strTemplate = strTemplate & "                     {$IsHiddenShow}" & vbCrLf
    strTemplate = strTemplate & "{$GuestContentShow}" & vbCrLf
    strTemplate = strTemplate & "【LastReplyShow】<table width='98%' align='right'  cellpadding='5' cellspacing='0' class='Guest_border2'>" & vbCrLf
    strTemplate = strTemplate & "  <tr><td align='left' valign='top' class='Guest_ReplyUser'> 回复主题：{$LastReplyTitle}     回复人:{$LastReplyGuest}</td>       </tr>       <tr>     <td colspan=2>" & vbCrLf
    strTemplate = strTemplate & "{$LastReplyContent}</td></tr></table>【/LastReplyShow】                     </td>" & vbCrLf
    strTemplate = strTemplate & "                    </tr>" & vbCrLf
    strTemplate = strTemplate & "                    <tr>" & vbCrLf
    strTemplate = strTemplate & "                      <td align='left' valign='bottom'>" & vbCrLf
    strTemplate = strTemplate & "                     【AdminReplyShow】" & vbCrLf
    strTemplate = strTemplate & "                                             <table width='100%' border='0' cellspacing='0' cellpadding='2'>" & vbCrLf
    strTemplate = strTemplate & "                          <tr>" & vbCrLf
    strTemplate = strTemplate & "                            <td height='1' class='Guest_tdbg_1px'></td>" & vbCrLf
    strTemplate = strTemplate & "                          </tr>" & vbCrLf
    strTemplate = strTemplate & "                          <tr>" & vbCrLf
    strTemplate = strTemplate & "                            <td valign='top'>" & vbCrLf
    strTemplate = strTemplate & "                              <table width='100%' border='0' cellpadding='0' cellspacing='0' style='TABLE-LAYOUT: fixed' class='Guest_border2'>" & vbCrLf
    strTemplate = strTemplate & "                                <tr>" & vbCrLf
    strTemplate = strTemplate & "                                  <td class='Guest_ReplyAdmin'> 管理员[{$ReplyAdmin}]回复:</td>" & vbCrLf
    strTemplate = strTemplate & "                                </tr>" & vbCrLf
    strTemplate = strTemplate & "                                <tr>" & vbCrLf
    strTemplate = strTemplate & "                                  <td valign='bottom'>{$AdminReplyContent}    回复时间:" & vbCrLf
    strTemplate = strTemplate & "{$AdminReplyTime}</td>" & vbCrLf
    strTemplate = strTemplate & "                                </tr>" & vbCrLf
    strTemplate = strTemplate & "                              </table>" & vbCrLf
    strTemplate = strTemplate & "                            </td>" & vbCrLf
    strTemplate = strTemplate & "                          </tr>" & vbCrLf
    strTemplate = strTemplate & "                        </table>【/AdminReplyShow】" & vbCrLf
    strTemplate = strTemplate & "                       " & vbCrLf
    strTemplate = strTemplate & "                      【/ContentShow】</td>" & vbCrLf
    strTemplate = strTemplate & "                    </tr>" & vbCrLf
    strTemplate = strTemplate & "                  </table>" & vbCrLf
    strTemplate = strTemplate & "                  <table width='100%' height='1' border='0' cellpadding='0' cellspacing='0' class='Guest_tdbg_1px'>" & vbCrLf
    strTemplate = strTemplate & "                    <tr>" & vbCrLf
    strTemplate = strTemplate & "                      <td></td>" & vbCrLf
    strTemplate = strTemplate & "                    </tr>" & vbCrLf
    strTemplate = strTemplate & "                  </table>" & vbCrLf
    strTemplate = strTemplate & "                  <table width=100% border=0 cellpadding=0 cellspacing=3>" & vbCrLf
    strTemplate = strTemplate & "                    <tr>" & vbCrLf
    strTemplate = strTemplate & "                      <td>" & vbCrLf
    strTemplate = strTemplate & "{$HomePagePic}" & vbCrLf
    strTemplate = strTemplate & "{$OicqPic}" & vbCrLf
    strTemplate = strTemplate & "{$EmailPic}" & vbCrLf
    strTemplate = strTemplate & "{$OtherPic}{$ReplyPic}{$EditPic}{$DelPic}" & vbCrLf
    strTemplate = strTemplate & "               <td align='right'> " & vbCrLf
    strTemplate = strTemplate & "{$InfoShow}" & vbCrLf
    strTemplate = strTemplate & "                      </td>" & vbCrLf
    strTemplate = strTemplate & "                    </tr>" & vbCrLf
    strTemplate = strTemplate & "                  </table>" & vbCrLf
    strTemplate = strTemplate & "                </td>" & vbCrLf
    strTemplate = strTemplate & "              </tr>" & vbCrLf
    strTemplate = strTemplate & "            </table>" & vbCrLf
    strTemplate = strTemplate & "          </td>" & vbCrLf
    strTemplate = strTemplate & "        </tr>" & vbCrLf
    strTemplate = strTemplate & "      </table>" & vbCrLf
    strTemplate = strTemplate & "      <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    strTemplate = strTemplate & "        <tr>" & vbCrLf
    strTemplate = strTemplate & "          <td class='main_shadow'>" & vbCrLf
    strTemplate = strTemplate & "          </td>" & vbCrLf
    strTemplate = strTemplate & "        </tr>" & vbCrLf
    strTemplate = strTemplate & "      </table>" & vbCrLf
    strTemplate = strTemplate & "     【/GuestList1】 " & vbCrLf
    strTemplate = strTemplate & "     <!--留言本方式循环显示留言结束-->" & vbCrLf
    TemplateGuestBook = strTemplate
End Function

%>
