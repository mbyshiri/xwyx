<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
'������
IndexMaxPerPage1 = PE_CLng(arrMaxPerPage(0))
'���Ա�
IndexMaxPerPage2 = PE_CLng(arrMaxPerPage(1))
'�ظ�ҳ
ReplyMaxPerPage = PE_CLng(arrMaxPerPage(2))
'չ����
TreeMaxPerPage = arrMaxPerPage(3)


'����û��Ƿ��¼
UserLogined = CheckUserLogined()
strNavPath = XmlText("BaseText", "Nav", "�����ڵ�λ�ã�") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
strPageTitle = SiteTitle

Call GetChannel(ChannelID)

If Trim(ChannelName) <> "" And ShowChannelName = True Then
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a href='" & InstallDir & ChannelDir & "/Index.asp'>" & ChannelName & "</a>"
    strPageTitle = strPageTitle & " >> " & ChannelName
End If

'���԰��ڲ�����
SaveEdit = 0
GImagePath = InstallDir & "GuestBook/Images/"
GFacePath = InstallDir & "GuestBook/Images/Face/"
FileName = "index.asp"
KindID = PE_CLng(Trim(Request("KindID")))

TopicType = Trim(Request("topictype"))

'��ȡ�鿴��ʽ
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
'��������GuestBook_Top()
'��  �ã���ʾ�������Թ���
'��  ������
'=================================================
Private Function GuestBook_Top()
    Dim strTop
    strTop = ""
    If UserLogined = True Then
        strTop = Replace(XmlText("Guest", "GuestBook_Top/Logined", "<a href='{$FileName}?action=user'><img align=absmiddle src='images/Guest_user.gif' alt='�ҷ��������' border='0'></a>&nbsp;<a href='{$FileName}?action=user&topictype=participation'><img align=absmiddle src='images/Guest_participation.gif' alt='�һظ�������' border='0'></a>&nbsp;"), "{$FileName}", FileName)
    End If
    strTop = strTop & Replace(Replace(Replace(XmlText("Guest", "GuestBook_Top/NoLogin", "<a href='{$FileName}'><img align=absmiddle src='images/Guest_all.gif' alt='�鿴��������' border='0'></a>&nbsp;<a href='Guest_Write.asp?KindID={$KindID}&KindName={$KindName}'><img align=absmiddle src='images/Guest_write.gif' alt='ǩд�µ�����' border='0'></a>"), "{$FileName}", FileName), "{$KindID}", KindID), "{$KindName}", KindName) & vbCrLf
    GuestBook_Top = strTop
End Function
'=================================================
'��������GuestBook_Mode()
'��  �ã���ʾ�������Թ���
'��  ������
'=================================================
Private Function GuestBook_Mode()
    Dim strTop
    strTop = ""
    If UserLogined = True Then
        strTop = strTop & XmlText("Guest", "GuestBook_Mode/Mode1", "�û�ģʽ") & vbCrLf
    Else
        strTop = strTop & XmlText("Guest", "GuestBook_Mode/Mode2", "�ο�ģʽ") & vbCrLf
    End If
    GuestBook_Mode = strTop
End Function
'=================================================
'��������GuestBook_See()
'��  �ã���ʾ�������Թ���
'��  ������
'=================================================
Private Function GuestBook_See()
    Dim strTop
    strTop = ""
    If ShowGStyle = 1 Then
        strTop = strTop & XmlText("Guest", "GuestBook_See/Mode1", "��������ʽ") & vbCrLf
    Else
        strTop = strTop & XmlText("Guest", "GuestBook_See/Mode2", "���԰巽ʽ") & vbCrLf
    End If

    GuestBook_See = strTop
End Function
'=================================================
'��������GuestBook_Appear()
'��  �ã���ʾ�������Թ���
'��  ������
'=================================================
Private Function GuestBook_Appear()
    Dim strTop
    If CheckLevel = 0 Or NeedlessCheck = 1 Then
        strTop = strTop & XmlText("Guest", "GuestBook_Appear/Mode1", "ֱ�ӷ���") & vbCrLf
    Else
        strTop = strTop & XmlText("Guest", "GuestBook_Appear/Mode2", "��˷���") & vbCrLf
        Dim grs
        Set grs = Conn.Execute("select count(*) from PE_GuestBook where GuestIsPassed=" & PE_False & "")
        strTop = strTop & "&nbsp;&nbsp;" & Replace(XmlText("Guest", "GuestBook_Appear/Count", "��{$GuestNo}�������"), "{$GuestNo}", grs(0)) & vbCrLf
        Set grs = Nothing
    End If
    GuestBook_Appear = strTop
End Function

'=================================================
'��������GuestBook_Search()
'��  �ã���ʾ��������
'��  ������
'=================================================
Private Function GuestBook_Search()
    Dim strGuestSearch
    'If GuestBook_IsAssignSort = True Then
        'strGuestSearch = Replace(XmlText("Guest", "GuestBook_Search", "<table border='0' cellpadding='0' cellspacing='0'><form method='post' name='SearchForm' action='Search.asp'><tr><td height='30' >&nbsp;&nbsp;<select name='Field' id='1'><option value='Title' selected>��������</option><option value='Content'>��������</option><option value='Name'>������</option><option value='GuestTime'>����ʱ��</option><option value='Reply'>����Ա�ظ�</option></select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<select name='KindID' id='KindID'>{$KindID}</select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<input type='text' name='keyword'  size='15' value='�ؼ���' maxlength='45' onFocus='this.select();'>&nbsp;<input type='submit' name='Submit'  value='����'></td></tr></form></table>"), "{$KindID}", GetGKind_Option(3, KindID))
    'Else
        strGuestSearch = Replace(XmlText("Guest", "GuestBook_Search", "<table border='0' cellpadding='0' cellspacing='0'><form method='post' name='SearchForm' action='Search.asp'><tr><td height='30' >&nbsp;&nbsp;<select name='Field' id='1'><option value='Title' selected>��������</option><option value='Content'>��������</option><option value='Name'>������</option><option value='GuestTime'>����ʱ��</option><option value='Reply'>����Ա�ظ�</option></select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<select name='KindID' id='KindID'>{$KindID}</select>&nbsp;</td><td height='30' >&nbsp;&nbsp;<input type='text' name='keyword'  size='15' value='�ؼ���' maxlength='45' onFocus='this.select();'>&nbsp;<input type='submit' name='Submit'  value='����'></td></tr></form></table>"), "{$KindID}", GetGKind_Option(1, KindID))
    'End If
    GuestBook_Search = strGuestSearch
End Function




'=================================================
'��������ShowAllGuest()
'��  �ã���ҳ��ʾ��������
'��  ����ShowType-----  0Ϊ��ʾ����
'                       1Ϊ��ʾ��ͨ����˼��û��Լ����������
'                       2Ϊ��ʾ��ͨ����˵����ԣ������ο���ʾ��
'                       3Ϊ��ʾ�û��Լ����������
'                       4Ϊ��ʾ�Ƽ�����������
'                       5ΪҪ�༭������
'                       6Ϊ�ظ�ҳ������
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
                    ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "ShowAllGuest/Err1", "����Ĺؼ��ֲ�����Ч���ڣ�") & "</li>"
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
        ErrMsg = ErrMsg & "<li>" & XmlText("Guest", "ShowAllGuest/NoFound", "û���κ�����") & "</li>"
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
'��������ShowJS_Guest()
'��  �ã��ύ���Ե������ж�
'��  ������
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
        strJS = strJS & "      alert('��������Ϊ�գ�');" & vbCrLf
        strJS = strJS & "      document.myform.GuestName.focus();" & vbCrLf
        strJS = strJS & "      return(false) ;" & vbCrLf
        strJS = strJS & "    }" & vbCrLf
    End If
    strJS = strJS & "  if(document.myform.GuestTitle.value==''){" & vbCrLf
    strJS = strJS & "    alert('���ⲻ��Ϊ�գ�');" & vbCrLf
    strJS = strJS & "    document.myform.GuestTitle.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  if(document.myform.GuestTitle.value.length>30){" & vbCrLf
    strJS = strJS & "    alert('���ⲻ�ܳ���30�ַ���');" & vbCrLf
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
    strJS = strJS & "    alert('���ݲ���Ϊ�գ�');" & vbCrLf
    strJS = strJS & "    editor.HtmlEdit.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  if(document.myform.GuestContent.value.length>65536){" & vbCrLf
    strJS = strJS & "    alert('���ݲ��ܳ���64K��');" & vbCrLf
    strJS = strJS & "    editor.HtmlEdit.focus();" & vbCrLf
    strJS = strJS & "    return(false);" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    If EnableGuestBookCheck = True Then
        strJS = strJS & "  if(document.myform.CheckCode.value==''){" & vbCrLf
        strJS = strJS & "    alert('������������֤�룡');" & vbCrLf
        strJS = strJS & "    document.myform.CheckCode.focus();" & vbCrLf
        strJS = strJS & "    return(false);" & vbCrLf
        strJS = strJS & "  }" & vbCrLf
    End If
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "</script>" & vbCrLf
    ShowJS_Guest = strJS
End Function

'**************************************************
'��������KeywordReplace
'��  �ã���ʾ�����ؼ���
'��  ����strChar-----Ҫת�����ַ�
'����ֵ��ת������ַ�
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
'��������Guest_info()
'��  �ã����Բ�����Ϣ
'��  ����info ��ʾ��Ϣ����
'=================================================
Private Function Guest_info(info)
    Dim strInfo
    'strInfo = Replace(Replace(XmlText("Guest", "Guest_info", "<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr align='center'><td class='Guest_title_760'>���Բ���������Ϣ</td></tr><tr><td class='main_tdbg_575'><table cellpadding=5 cellspacing=0 border=0 width=100% align=center><tr><td height='100' valign='top'>{$info}</td></tr><tr align='center' class='tdbg'><td><a href='{$FileName}'>���鿴���ԡ�</a><a href='Guest_Write.asp'>��ǩд���ԡ�</a></td></tr></table></td></tr></table><br>"), "{$info}", info), "{$FileName}", FileName)
    strInfo = Replace(Replace(XmlText("Guest", "Guest_info", "<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr align='center'><td class='Guest_title_760'>���Բ���������Ϣ</td></tr><tr><td class='main_tdbg_575'><table cellpadding=5 cellspacing=0 border=0 width=100% align=center><tr><td height='100' valign='top'>{$info}</td></tr><tr align='center' class='tdbg'><td><a href='{$FileName}'>���鿴���ԡ�</a><a href='javascript:history.go(-1)'>��ǩд���ԡ�</a></td></tr></table></td></tr></table><br>"), "{$info}", info), "{$FileName}", FileName)
    Guest_info = strInfo
End Function

'=================================================
'��������GetGKind_Option()
'��  �ã��������������
'��  ����ShowType ��ʾ����
'        KindID   ���
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
        strOption = strOption & ">��ָ�����</option>"
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
'��������GetGKindList()
'��  �ã�������ʾ�������
'��  ������
'=================================================
Private Function GetGKindList()
    Dim rsGKind, sqlGKind, strGKind, i
    sqlGKind = "select * from PE_Guestkind order by OrderID"
    Set rsGKind = Conn.Execute(sqlGKind)
    If rsGKind.BOF And rsGKind.EOF Then
        strGKind = "| " & XmlText("Guest", "KindList/Nofound", "û���κ����")
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
        'strGKind = strGKind & "<a href='index.asp?KindID=0'>" & XmlText("xxxxx", "xxxxxx", "�����κ����") & "</a> |"
    'End If
    GetGKindList = strGKind
End Function

'=================================================
'��������ShowGueststyle()
'��  �ã���ȡ�鿴��ʽ
'��  ������
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
'��������ShowGueststyle()
'��  �ã���ʾ�л���ʽ
'��  ������
'=================================================
Private Function ShowGueststyle()
    Dim Shtm
    If ShowGStyle = 1 Then
        Shtm = "<a class=Guest href=ShowGuestStyle.asp?ShowGStyle=2>" & XmlText("Guest", "ShowGueststyle/Mode1", "�л������Ա���ʽ") & "</a>"
    Else
        Shtm = "<a class=Guest href=ShowGuestStyle.asp?ShowGStyle=1>" & XmlText("Guest", "ShowGueststyle/Mode2", "�л�����������ʽ") & "</a>"
    End If
    ShowGueststyle = Shtm
End Function
'=================================================
'��������TransformTime()
'��  �ã���ʽ��ʱ��
'��  ����ʱ��
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
       dayshow = XmlText("Guest", "TransformTime/d1", "���� ")
    ElseIf dnt = 1 Then
       dayshow = XmlText("Guest", "TransformTime/d2", "���� ")
    ElseIf dnt = 2 Then
       dayshow = XmlText("Guest", "TransformTime/d3", "ǰ�� ")
    End If
    TransformTime = dayshow & pshow & thour & ":" & tminute
End Function

'=================================================
'��������TransformIP()
'��  �ã���ʽ��IP
'��  ����IP
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
'��������ShowTip()
'��  �ã���꾭����ʾ��ʾ
'��  ������
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
'��������GetRepeatGuestBook()
'��  �ã����Ա���ʽʱ�滻Ҫѭ���ı�ǩ
'��  ����Ҫ�滻��ֵ
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
    GtbDel = XmlText("Guest", "GuestBookShow/Del", "����ɾ����")
    GtbNoEnter = XmlText("BaseText", "NoEnter", "δ��")
    GtbMan = XmlText("BaseText", "Man", "��")
    GtbGirl = XmlText("BaseText", "Girl", "Ů")
    GtbTip1 = XmlText("Guest", "GuestBookShow/Tip1", " ������{$Name} {$Sex}<br> ��ҳ��{$Homepage}<br> OICQ��{$Oicq}<br> ���䣺{$Email}<br> ��ַ��{$GuestIP}<br> ʱ�䣺{$Time}")
    GtbTip2 = XmlText("Guest", "GuestBookShow/Tip2", "�û�������ϱ��ܡ�")
    Gtbp1 = XmlText("Guest", "GuestBookShow/p1", "�̶�����")
    Gtbp2 = XmlText("Guest", "GuestBookShow/p2", "��������")
    Gtbp3 = XmlText("Guest", "GuestBookShow/p3", "�лظ�")
    Gtbp4 = XmlText("Guest", "GuestBookShow/p4", "�޻ظ�")
    Gtbp5 = XmlText("Guest", "GuestBookShow/p5", "�ظ���")
    Gtbp6 = XmlText("Guest", "GuestBookShow/p6", "���⣺")
    GtbGuestImages = XmlText("Guest", "GuestBookShow/GuestImages", "<img src='{$GuestImages}.gif' width='80' height='90' onMouseOut=toolTip() onMouseOver=""toolTip('{$GuestTip}')"">")
    GtbUser = XmlText("Guest", "GuestBookShow/User", "�û�")
    GtbGuest = XmlText("Guest", "GuestBookShow/Guest", "�ο�")
    GtbHide1 = XmlText("Guest", "GuestBookShow/Hide1", " **************************************<br> * �������ԣ�����Ա�������û����Կ��� *<br> **************************************")
    GtbHide2 = XmlText("Guest", "GuestBookShow/Hide2", "[����]")
    GtbHide3 = XmlText("Guest", "GuestBookShow/Hide3", " *********************************************<br> * ���ع���Ա�ظ�������Ա�������û����Կ��� *<br> *********************************************")
    GtbReply4 = XmlText("Guest", "GuestBookShow/Reply4", "�ظ���������")
    GtbReply5 = XmlText("Guest", "GuestBookShow/Reply5", "�༭��������")
    GtbReply6 = XmlText("Guest", "GuestBookShow/Reply6", "ȷ��Ҫɾ����������")
    GtbReply7 = XmlText("Guest", "GuestBookShow/Reply7", "ɾ����������")
    GtbReply8 = XmlText("Guest", "GuestBookShow/Reply8", "�鿴ȫ���ظ�")
    GtbReply9 = XmlText("Guest", "GuestBookShow/Reply9", "���лظ�{$ReplyNum}��")
    GtbReply10 = XmlText("Guest", "GuestBookShow/Reply10", "�ظ���������")
    GtbReply11 = XmlText("Guest", "GuestBookShow/Reply11", "�����б�")

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
        '�滻��������
        regEx.Pattern = "��ContentShow��([\s\S]*?)��\/ContentShow��"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            ContentShow = Match.value
        Next
        If rsGuest("GuestIsPrivate") = True And rsGuest("GuestName") <> UserName And rsGuest("ReplyIsPrivate") = True Then
            strTemp = Replace(strTemp, ContentShow, "<br><br><font class=Guest_font>" & GtbHide1 & "</font>")
        End If
        strTemp = Replace(strTemp, "��ContentShow��", "")
        strTemp = Replace(strTemp, "��/ContentShow��", "")

        If rsGuest("GuestIsPrivate") = True And rsGuest("GuestName") <> UserName Then
            strTemp = Replace(strTemp, "{$IsHiddenShow}", "                        <font class=Guest_font>" & GtbHide2 & "</font>&nbsp;")
            strTemp = Replace(strTemp, "{$GuestContentShow}", "")
        Else
            strTemp = Replace(strTemp, "{$IsHiddenShow}", "")
        End If
        strTemp = Replace(strTemp, "{$GuestContentShow}", KeywordReplace(FilterJS(rsGuest("GuestContent"))))
        
   
        '�滻�û����ظ�
        regEx.Pattern = "��LastReplyShow��([\s\S]*?)��\/LastReplyShow��"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            LastReplyShow = Match.value
        Next

        If rsGuest("LastReplyGuest") = "" Or IsNull(rsGuest("LastReplyGuest")) Or ReplyId <> "" Then
            strTemp = Replace(strTemp, LastReplyShow, "")
        End If

        strTemp = Replace(strTemp, "��LastReplyShow��", "")
        strTemp = Replace(strTemp, "��/LastReplyShow��", "")
        strTemp = Replace(strTemp, "{$LastReplyContent}", KeywordReplace(rsGuest("LastReplyContent")))
        strTemp = Replace(strTemp, "{$LastReplyGuest}", KeywordReplace(rsGuest("LastReplyGuest")))
        strTemp = Replace(strTemp, "{$LastReplyTitle}", KeywordReplace(rsGuest("LastReplyTitle")))
        strTemp = Replace(strTemp, "{$LastReplyTime}", KeywordReplace(rsGuest("LastReplyTime")))
        '�滻�û����ظ����
        
        '�滻����Ա�ظ�

        regEx.Pattern = "��AdminReplyShow��([\s\S]*?)��\/AdminReplyShow��"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            AdminReplyShow = Match.value
        Next
        
        If rsGuest("GuestReply") = "" Or IsNull(rsGuest("GuestReply")) Then
            strTemp = Replace(strTemp, AdminReplyShow, "")
        ElseIf rsGuest("ReplyIsPrivate") = True And rsGuest("GuestName") <> UserName Then
            strTemp = Replace(strTemp, AdminReplyShow, "<font class=Guest_font>" & GtbHide3 & "</font>")
        Else
                        strTemp = Replace(strTemp, "��AdminReplyShow��", "")
                        strTemp = Replace(strTemp, "��/AdminReplyShow��", "")
                        strTemp = Replace(strTemp, "{$ReplyAdmin}", KeywordReplace(rsGuest("GuestReplyAdmin")))
                        strTemp = Replace(strTemp, "{$AdminReplyTime}", KeywordReplace(rsGuest("GuestReplyDatetime")))
                        strTemp = Replace(strTemp, "{$AdminReplyContent}", KeywordReplace(rsGuest("GuestReply")))
        End If
        '�滻����Ա�ظ����
        
        '�滻�����������


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
            strTemp = Replace(strTemp, "{$OtherPic}", "<img src=" & GImagePath & "other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;Icq��" & UserIcq & "<br>&nbsp;Msn��" & UserMsn & "<br>&nbsp;I P��" & rsGuest("GuestIP") & "')"">")
        Else
            strTemp = Replace(strTemp, "{$OtherPic}", "<img src=" & GImagePath & "other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;Icq��" & UserIcq & "<br>&nbsp;Msn��" & UserMsn & "<br>&nbsp;I P��" & TransformIP(rsGuest("GuestIP")) & "')"">")
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
                strTemp = Replace(strTemp, "{$InfoShow}", "<a href=Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & "><img src=" & GImagePath & "reply0.gif width=15 height=16 border=0>&nbsp;" & GtbReply8 & "</a>(��" & rsGuest("ReplyNum") & "��)")
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
'��������GetRepeatDiscussion()
'��  �ã���������ʽʱ�滻Ҫѭ���ı�ǩ
'��  ����Ҫ�滻��ֵ
'=================================================
Private Function GetRepeatDiscussion(strParameter, strList)
    Dim strTemp, arrTemp, strBeg, strEnd, strSource
    Dim strSoftPic, strPicTemp, arrPicTemp

    If strParameter = "" Or IsNull(strParameter) Then
		GetRepeatDiscussion = ""
        Exit Function
    End If
	strSource = strList

    '�滻��������ʽ�б�ͷ��
    regEx.Pattern = "��GuestList2_Beg��([\s\S]*?)��\/GuestList2_Beg��"
    Set Matches = regEx.Execute(strSource)
    For Each Match In Matches
        strBeg = Match.SubMatches(0)
        strSource = Replace(strSource, Match.value, "")
    Next

    '�滻��������ʽ�б�β��
    regEx.Pattern = "��GuestList2_End��([\s\S]*?)��\/GuestList2_End��"
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
    strXml = Split(XmlText("Guest", "discussionShow/Text", "��������|||��������|||������|||�ظ�|||�Ķ�|||���ظ�|||�̶�����|||��������|||չ������ظ����б�|||�޻ظ�|||����鿴��¼������Ϣ"), "|||")

    Dim i, GtbUser, GtbGuest
    GtbUser = XmlText("Guest", "GuestBookShow/User", "�û�")
    GtbGuest = XmlText("Guest", "GuestBookShow/Guest", "�ο�")
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
'��������GuestFace()
'��  �ã���������ѡ��
'��  ������
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
'��������ManageRubbishContent()
'��  �ã�������������Ӻ���
'��  ������
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
'��������GuestContent()
'��  �ã���������
'��  ������
'=================================================
Private Function GuestContent()
    Dim strHTM
    strHTM = "<textarea name='GuestContent' id='GuestContent' style='display:none' >" & Server.HTMLEncode(FilterJS(WriteContent)) & "</textarea>" & vbCrLf
    strHTM = strHTM & "<iframe ID='editor' src='../editor.asp?ChannelID=4&ShowType=2&tContentid=GuestContent' frameborder='1' scrolling='no' width='480' height='280' ></iframe>" & vbCrLf
    GuestContent = strHTM
End Function

'=================================================
'��������DefaultTemplate()
'��  �ã��õ�Ĭ��ģ�����ã�����ǰ�汾����
'��  ����strType ҳ�����
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
'��������TemplateDiscission()
'��  �ã��õ���������ʽĬ��ģ������
'��  ����strstlye1 ҳ�����
'=================================================
Private Function TemplateDiscission(strstlye1)
    Dim strTemplate
    strTemplate = ""
    If strstlye1 = "Index" Then
        strTemplate = strTemplate & "     <!--��������ʽѭ����ʾ���Խ���-->   ��GuestList2(0)��" & vbCrLf
    Else
        strTemplate = strTemplate & "     <!--��������ʽѭ����ʾ���Խ���-->   ��GuestList2(1)��" & vbCrLf
    End If
    strTemplate = strTemplate & "             ��GuestList2_Beg��<table width='100%' class='Guest_border' border='0' cellspacing='1' cellpadding='0' align='center'>" & vbCrLf
    strTemplate = strTemplate & "      <tr class='Guest_title'>" & vbCrLf
    strTemplate = strTemplate & "        <td width='58%' colspan='3'> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>��������</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>" & vbCrLf
    strTemplate = strTemplate & "        <td width='10%' nowrap> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>������</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>" & vbCrLf
    strTemplate = strTemplate & "        <td width='5%' nowrap> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>�ظ�</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>   " & vbCrLf
    strTemplate = strTemplate & "        <td width='5%' nowrap> " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>�Ķ�</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>   " & vbCrLf
    strTemplate = strTemplate & "        <td width='22%' nowrap>  " & vbCrLf
    strTemplate = strTemplate & "          <div align='center'><b>���ظ�</b></div>" & vbCrLf
    strTemplate = strTemplate & "        </td>    " & vbCrLf
    strTemplate = strTemplate & "      </tr>��/GuestList2_Beg��" & vbCrLf
    strTemplate = strTemplate & "<tr class='Guest_tdbg'>" & vbCrLf
    strTemplate = strTemplate & "<td width='5%' align='center'>" & vbCrLf
    strTemplate = strTemplate & "  {$GuestFaceShow}" & vbCrLf
    strTemplate = strTemplate & "  </td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='5%'  align='center'>" & vbCrLf
    strTemplate = strTemplate & "{$IsTitlePic}" & vbCrLf
    strTemplate = strTemplate & "  </td><td width='48%' title='����鿴��¼������Ϣ' align='left'>{$GuestTitle}<I><font color=gray>({$GuestContentLength}��)</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='10%' align='center'>{$GuestName}</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='5%' align='center'>{$ReplyNum}</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='5%' align='center'>{$Hits}</td>" & vbCrLf
    strTemplate = strTemplate & "  <td width='22%' align='left'>{$GuestTime}<font class=Channel_font> | </font>{$LastReplyGuest}</td></tr>" & vbCrLf
    strTemplate = strTemplate & "  <tr id='FollowTr{$GuestID}' style='display:none;'><td id='FollowTd{$GuestID}' colspan='7'></td></tr>" & vbCrLf
    strTemplate = strTemplate & "��GuestList2_End��</table>��/GuestList2_End����/GuestList2��" & vbCrLf
    strTemplate = strTemplate & "     <!--��������ʽѭ����ʾ���Խ���-->" & vbCrLf
    TemplateDiscission = strTemplate
End Function

'=================================================
'��������TemplateGuestBook()
'��  �ã��õ����Ա���ʽĬ��ģ������
'��  ����strstlye2 ҳ�����
'=================================================
Private Function TemplateGuestBook(strstlye2)
    Dim strTemplate
    strTemplate = ""
    strTemplate = strTemplate & "                     <!--���Ա���ʽѭ����ʾ���Կ�ʼ-->" & vbCrLf
    If strstlye2 = "Index" Then
        strTemplate = strTemplate & "     ��GuestList1(0)��" & vbCrLf
    ElseIf strstlye2 = "strWrite" Then
        strTemplate = strTemplate & "     ��GuestList1(1)��" & vbCrLf
    ElseIf strstlye2 = "Reply" Then
        strTemplate = strTemplate & "     ��GuestList1(2)��" & vbCrLf
    Else
        strTemplate = strTemplate & "     ��GuestList1(3)��" & vbCrLf
    End If
    strTemplate = strTemplate & "          <table width='100%' border='0' cellpadding='0' cellspacing='1' class='Guest_border'>" & vbCrLf
    strTemplate = strTemplate & "        <tr>" & vbCrLf
    strTemplate = strTemplate & "          <td align='center' valign='top'>" & vbCrLf
    strTemplate = strTemplate & "            <table width='100%' border='0' cellspacing='0' cellpadding='0' class='Guest_title'>" & vbCrLf
    strTemplate = strTemplate & "              <tr>" & vbCrLf
    strTemplate = strTemplate & "                <td>" & vbCrLf
    strTemplate = strTemplate & "{$IsTitlePic}<strong>{$GuestType}��</strong>{$GuestTitle}" & vbCrLf
    strTemplate = strTemplate & "                </td>" & vbCrLf
    strTemplate = strTemplate & "                <td width='180'>" & vbCrLf
    strTemplate = strTemplate & "                  <img src='{$InstallDir}Images/posttime.gif' width='11' height='11' align='absmiddle'>��{$GuestTime}" & vbCrLf
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
    strTemplate = strTemplate & "��{$GuestNameType}��<br>{$GuestName}                </td>" & vbCrLf
    strTemplate = strTemplate & "                <td align='center' height='153' width='1' class='Guest_tdbg_1px'></td>" & vbCrLf
    strTemplate = strTemplate & "                <td>" & vbCrLf
    strTemplate = strTemplate & "                  <table width='100%' border='0' cellpadding='6' cellspacing='0' height='125' style='TABLE-LAYOUT: fixed'>" & vbCrLf
    strTemplate = strTemplate & "                    <tr>" & vbCrLf
    strTemplate = strTemplate & "                      <td align='left' valign='top'>{$GuestFaceShow}��ContentShow��" & vbCrLf
    strTemplate = strTemplate & "                     {$IsHiddenShow}" & vbCrLf
    strTemplate = strTemplate & "{$GuestContentShow}" & vbCrLf
    strTemplate = strTemplate & "��LastReplyShow��<table width='98%' align='right'  cellpadding='5' cellspacing='0' class='Guest_border2'>" & vbCrLf
    strTemplate = strTemplate & "  <tr><td align='left' valign='top' class='Guest_ReplyUser'> �ظ����⣺{$LastReplyTitle}     �ظ���:{$LastReplyGuest}</td>       </tr>       <tr>     <td colspan=2>" & vbCrLf
    strTemplate = strTemplate & "{$LastReplyContent}</td></tr></table>��/LastReplyShow��                     </td>" & vbCrLf
    strTemplate = strTemplate & "                    </tr>" & vbCrLf
    strTemplate = strTemplate & "                    <tr>" & vbCrLf
    strTemplate = strTemplate & "                      <td align='left' valign='bottom'>" & vbCrLf
    strTemplate = strTemplate & "                     ��AdminReplyShow��" & vbCrLf
    strTemplate = strTemplate & "                                             <table width='100%' border='0' cellspacing='0' cellpadding='2'>" & vbCrLf
    strTemplate = strTemplate & "                          <tr>" & vbCrLf
    strTemplate = strTemplate & "                            <td height='1' class='Guest_tdbg_1px'></td>" & vbCrLf
    strTemplate = strTemplate & "                          </tr>" & vbCrLf
    strTemplate = strTemplate & "                          <tr>" & vbCrLf
    strTemplate = strTemplate & "                            <td valign='top'>" & vbCrLf
    strTemplate = strTemplate & "                              <table width='100%' border='0' cellpadding='0' cellspacing='0' style='TABLE-LAYOUT: fixed' class='Guest_border2'>" & vbCrLf
    strTemplate = strTemplate & "                                <tr>" & vbCrLf
    strTemplate = strTemplate & "                                  <td class='Guest_ReplyAdmin'> ����Ա[{$ReplyAdmin}]�ظ�:</td>" & vbCrLf
    strTemplate = strTemplate & "                                </tr>" & vbCrLf
    strTemplate = strTemplate & "                                <tr>" & vbCrLf
    strTemplate = strTemplate & "                                  <td valign='bottom'>{$AdminReplyContent}    �ظ�ʱ��:" & vbCrLf
    strTemplate = strTemplate & "{$AdminReplyTime}</td>" & vbCrLf
    strTemplate = strTemplate & "                                </tr>" & vbCrLf
    strTemplate = strTemplate & "                              </table>" & vbCrLf
    strTemplate = strTemplate & "                            </td>" & vbCrLf
    strTemplate = strTemplate & "                          </tr>" & vbCrLf
    strTemplate = strTemplate & "                        </table>��/AdminReplyShow��" & vbCrLf
    strTemplate = strTemplate & "                       " & vbCrLf
    strTemplate = strTemplate & "                      ��/ContentShow��</td>" & vbCrLf
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
    strTemplate = strTemplate & "     ��/GuestList1�� " & vbCrLf
    strTemplate = strTemplate & "     <!--���Ա���ʽѭ����ʾ���Խ���-->" & vbCrLf
    TemplateGuestBook = strTemplate
End Function

%>
