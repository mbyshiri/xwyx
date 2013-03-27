<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Const AdminType = True
Const EnableGuestCheck = "Yes"

Dim rs, sql, rsGuest, sqlGuest
Dim GuestID, Passed, GImagePath, GFacePath, GEmotPath, i, KindID, KindName

GImagePath = InstallDir & "GuestBook/Images/"
GFacePath = InstallDir & "GuestBook/Images/Face/"
GEmotPath = InstallDir & "GuestBook/Images/Emote/"


'检查管理员操作权限
If AdminPurview > 1 Then
    If AdminPurview_GuestBook = "" Then
        AdminPurview_GuestBook = 5
    Else
        AdminPurview_GuestBook = PE_CLng(AdminPurview_GuestBook)
    End If
    If AdminPurview_GuestBook > 3 Then
        PurviewPassed = False
    Else
        PurviewPassed = True
    End If
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Response.End
    End If
End If

Passed = Trim(Request("Passed"))
GuestID = Trim(Request("GuestID"))
KindID = PE_CLng(Trim(Request("KindID")))

If Passed = "" Then
    Passed = Session("Passed")
End If
Session("Passed") = Passed
If IsValidID(GuestID) = False Then
    GuestID = ""
End If

strFileName = "Admin_GuestBook.asp?Action=" & Action & "&Field=" & strField & "&KindID=" & KindID & "&keyword=" & Keyword
                                                    
Response.Write "<html><head><title>留言管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("网 站 留 言 管 理", 10141)

Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td>"
Response.Write "    <a href='Admin_GuestBook.asp?Passed=All'>网站留言管理</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_GuestBook.asp?Passed=False'>网站留言审核</a>&nbsp;|&nbsp;"
If AdminPurview = 1 Or AdminPurview_GuestBook < 3 Then
    Response.Write "    <a href='Admin_GuestBook.asp?Action=GKind'>留言类别管理</a>&nbsp;|&nbsp;"
    Response.Write "    <a href='Admin_GuestBook.asp?Action=AddGKind'>添加留言类别</a>&nbsp;|&nbsp;"
End If
If AdminPurview = 1 Or AdminPurview_GuestBook < 2 Then
    Response.Write "    <a href='Admin_GuestBook.asp?Action=CreateCode'>首页嵌入代码生成</a>"
End If
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
If Action = "" Then
    Response.Write "<form name='form' method='Post' action='Admin_GuestBook.asp'><tr class='tdbg'>"
    Response.Write "      <td width='70' height='30' ><strong>留言选项：</strong></td><td>"
    Response.Write "  <input name='Passed' type='radio' value='All' onclick='submit();'"
    If Passed = "All" Then Response.Write " checked"
    Response.Write ">所有留言&nbsp;&nbsp;&nbsp;&nbsp;<input name='Passed' type='radio' value='False' onclick='submit();'"
    If Passed = "False" Then Response.Write " checked"
    Response.Write ">未审核的留言&nbsp;&nbsp;&nbsp;&nbsp;<input name='Passed' type='radio' value='True' onclick='submit();'"
    If Passed = "True" Then Response.Write " checked"
    Response.Write ">已审核的留言</td></tr></form>" & vbCrLf
End If
Response.Write "</table>" & vbCrLf


Select Case Action
Case "Modify"
    Call Modify
Case "Show"
    Call Show
Case "SaveModify"
    Call SaveModify
Case "AdminReply"
    Call AdminReply
Case "SaveAdminReply"
    Call SaveAdminReply
Case "Del", "SetPassed", "CancelPassed", "DelReply", "Quintessence", "Cquintessence", "SetOnTop", "CancelOnTop"
    Call SetProperty
Case "GKind"
    Call GKind
Case "AddGKind"
    Call AddGKind
Case "ModifyGKind"
    Call ModifyGKind
Case "DelGKind", "ClearGKind"
    Call DelGKind
Case "SaveAddGKind", "SaveModifyGKind"
    Call SaveGKind
Case "OrderGuestKind"
    Call OrderGuestKind
Case "MoveGuest"
    Call MoveGuest
Case "Move"
    Call Move
Case "BatchMove"
    Call BatchMove
Case "DoBatchMove"
    Call DoBatchMove
Case "CreateCode"
    Call CreateCode
Case "DoCreateCode"
    Call DoCreateCode
Case Else
    Call Main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub Main()
    Dim GKind
    If KindID > 0 Then
        Set GKind = Conn.Execute("select * from PE_GuestKind where KindID=" & KindID)
        If GKind.BOF And GKind.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的类别</li>"
            Exit Sub
        Else
            KindName = GKind("KindName")
        End If
    End If
    
    Call ShowJS_Main("留言")
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetGKindList() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_GuestBook.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "  <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "     <tr class='title'>"
    Response.Write "    <td width='30' height='22' align='center'><strong>属性</strong></td>"
    Response.Write "    <td width='30' height='22' align='center'><strong>选中</strong></td>"
    Response.Write "    <td width='85' height='22' align='center'><strong>留言人</strong></td>"
    Response.Write "    <td height='22' align='center'><strong>留言主题</strong></td>"
    'Response.Write "    <td width='120' height='22' align='center'><strong>留言时间</strong></td>"
    Response.Write "    <td width='30' height='22' align='center'><strong>审核</strong></td>"
    Response.Write "    <td width='328' height='22' align='center'><strong>操作</strong></td>"
    Response.Write "  </tr>"

    sqlGuest = " select G.*,K.KindName from PE_GuestBook G"
    sqlGuest = sqlGuest & " left join PE_GuestKind K on G.KindID=K.KindID where 1=1"
    If Passed = "True" Then
        sqlGuest = sqlGuest & " and GuestIsPassed=1"
    ElseIf Passed = "False" Then
        sqlGuest = sqlGuest & " and GuestIsPassed=0"
    End If
    If KindID > 0 Then
        sqlGuest = sqlGuest & " and G.KindID=" & KindID
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "GuestTitle"
            sqlGuest = sqlGuest & " and GuestTitle like '%" & Keyword & "%' "
        Case "GuestContent"
            sqlGuest = sqlGuest & " and GuestContent like '%" & Keyword & "%' "
        Case "GuestReply"
            sqlGuest = sqlGuest & " and GuestReply like '%" & Keyword & "%' "
        Case "GuestName"
            sqlGuest = sqlGuest & " and GuestName like '%" & Keyword & "%' "
        Case Else
            sqlGuest = sqlGuest & " and GuestTitle like '%" & Keyword & "%' "
        End Select
    End If
    sqlGuest = sqlGuest & " order by G.TopicID desc,G.GuestId asc"
    Set rsGuest = Server.CreateObject("adodb.recordset")
    rsGuest.Open sqlGuest, Conn, 1, 1
    If rsGuest.BOF And rsGuest.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何留言！<br><br></td></tr>"
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

        Dim GuestNum
        GuestNum = 0

        Do While Not rsGuest.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If rsGuest("TopicID") = rsGuest("GuestID") Then
                Response.Write "      <td width='30' align='center'>主题</td>"
            Else
                Response.Write "      <td width='30' align='center' class='tdbg'></td>"
            End If
            Response.Write "      <td width='30' align='center'><input name='GuestID' type='checkbox' onclick='unselectall()' value='" & rsGuest("GuestID") & "'></td>"
            Response.Write "      <td width='85' align='center'><div style='cursor:hand' "
            If rsGuest("GuestType") = 1 Then
                Dim rsUser
                Set rsUser = Conn.Execute("select * from PE_Contacter where ContacterID=(select ContacterID from PE_User where UserName='" & ReplaceBadChar(rsGuest("GuestName")) & "')")
                If Not (rsUser.BOF And rsUser.EOF) Then
                    Dim QQ, icq, msn, Homepage
                    Homepage = rsUser("Homepage")
                    QQ = rsUser("QQ")
                    icq = rsUser("ICQ")
                    msn = rsUser("MSN")
                    Response.Write " title='类型：注册用户" & vbCrLf
                    Response.Write "性别："
                    If rsUser("Sex") = "0" Then
                        Response.Write "女"
                    Else
                        Response.Write "男"
                    End If
                    Response.Write vbCrLf & "信箱：" & rsUser("Email") & vbCrLf & "OICQ：" & QQ & vbCrLf & " ICQ：" & icq & vbCrLf & " MSN：" & msn & vbCrLf & "主页：" & Homepage & vbCrLf & "  IP：" & rsGuest("GuestIP") & "'"
                    '完毕
                End If
                Set rsUser = Nothing
            Else
                Response.Write " title='类型：游客" & vbCrLf
                Response.Write "性别："
                If rsGuest("GuestSex") = "0" Then
                    Response.Write "女"
                Else
                    Response.Write "男"
                End If
                Response.Write vbCrLf & "信箱：" & rsGuest("GuestEmail") & vbCrLf & "OICQ：" & rsGuest("GuestOicq") & vbCrLf & " ICQ：" & rsGuest("GuestIcq") & vbCrLf & " MSN：" & rsGuest("GuestMsn") & vbCrLf & "主页：" & rsGuest("GuestHomepage") & vbCrLf & "  IP：" & rsGuest("GuestIP") & "'"
            End If

            Response.Write " >" & rsGuest("GuestName") & "</div></td>"
            Response.Write "      <td><a href='Admin_GuestBook.asp?Action=Show&GuestID=" & rsGuest("GuestID") & "'>"
            If rsGuest("GuestIsPrivate") = True Then
                Response.Write "<font color=green>【隐藏】</font>" & vbCrLf
            End If
            Dim Title
            Title = rsGuest("GuestTitle")
            If Len(Title) > 18 Then
                Title = Left(Title, 18) & "..."
            End If
            If rsGuest("KindName") <> "" Then
                Response.Write "[" & rsGuest("KindName") & "]" & Title & "</a>"
            Else
                Response.Write "[不指定类别]" & Title & "</a>"
            End If
            'Response.Write "      <td width='120' align='center'>"
            If rsGuest("GuestDatetime") <> "" Then
                Response.Write "(" & TransformTime(FormatDateTime(rsGuest("GuestDatetime"), 0)) & ")"
            End If
            Response.Write "</td>"
            Response.Write "      <td width='30' align='center'>"
            If rsGuest("GuestIsPassed") = True Then
                Response.Write "<b>√</b>"
            Else
                Response.Write "<font color=red><b>×</b></font>"
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='328' align='center'>"
            
            If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(0, rsGuest("KindID")) = True Then
                Response.Write "      <a href='Admin_GuestBook.asp?Action=Modify&GuestID=" & rsGuest("GuestID") & "'>修改</a>"
            End If
            If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(1, rsGuest("KindID")) = True Then
                If rsGuest("TopicID") <> rsGuest("GuestID") Then
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=Del&GuestID=" & rsGuest("GuestID") & "' onClick=""return confirm('确定要删除此回复吗？');"">删除</a>"
                Else
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=Del&GuestID=" & rsGuest("GuestID") & "' onClick=""return confirm('删除此主题将删除其下所有回复，确定要删除此主题吗？');"">删除</a>"
                End If
            End If
            If (AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(2, rsGuest("KindID")) = True) And rsGuest("TopicID") = rsGuest("GuestID") Then
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=Move&GuestID=" & rsGuest("GuestID") & "'>移动</a>"
            End If
            If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(3, rsGuest("KindID")) = True Then
                If rsGuest("GuestIsPassed") = False Then
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=SetPassed&GuestID=" & rsGuest("GuestID") & "'>通过审核</a>"
                Else
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=CancelPassed&GuestID=" & rsGuest("GuestID") & "'>取消审核</a>"
                End If
            End If
            If rsGuest("TopicID") = rsGuest("GuestID") Then
                If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(4, rsGuest("KindID")) = True Then
                    If rsGuest("Quintessence") = 0 Then
                        Response.Write "      <a href='Admin_GuestBook.asp?Action=Quintessence&GuestID=" & rsGuest("GuestID") & "'>推荐精华</a>"
                    Else
                        Response.Write "      <a href='Admin_GuestBook.asp?Action=Cquintessence&GuestID=" & rsGuest("GuestID") & "'>取消精华</a>"
                    End If
                End If
                If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(5, rsGuest("KindID")) = True Then
                    If rsGuest("OnTop") = 0 Then
                        Response.Write "      <a href='Admin_GuestBook.asp?Action=SetOnTop&GuestID=" & rsGuest("GuestID") & "'>固顶</a>"
                    Else
                        Response.Write "      <a href='Admin_GuestBook.asp?Action=CancelOnTop&GuestID=" & rsGuest("GuestID") & "'>解固</a>"
                    End If
                End If
            End If
            If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Or CheckKindPurview(6, rsGuest("KindID")) = True Then
                Response.Write "      <a href='Admin_GuestBook.asp?Action=AdminReply&GuestID=" & rsGuest("GuestID") & "'>回复</a>"
                If rsGuest("GuestReply") <> "" Then
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=DelReply&GuestID=" & rsGuest("GuestID") & "'>清除回复</a>"
                End If
            End If
            Response.Write "      </td>"
            Response.Write "    </tr>"

            GuestNum = GuestNum + 1
            If GuestNum >= MaxPerPage Then Exit Do
            rsGuest.MoveNext
        Loop
    End If
    rsGuest.Close
    Set rsGuest = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='130' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中所有的留言</td><td>"
    
    Response.Write "<input type='submit' value='删除选定的留言' name='submit' onClick=""document.myform.Action.value='Del'"" "
    If CheckKindPurview(1, KindID) = False And AdminPurview = 2 And AdminPurview_GuestBook >= 3 Then Response.Write "disabled"
    Response.Write ">&nbsp;&nbsp;"
        
    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='SetPassed'"" value='审核通过选定的留言' "
    If CheckKindPurview(3, KindID) = False And AdminPurview = 2 And AdminPurview_GuestBook >= 3 Then Response.Write "disabled"
    Response.Write ">&nbsp;&nbsp;"
    Response.Write "<input name='submit2' type='submit' id='submit2' onClick=""document.myform.Action.value='CancelPassed'"" value='取消审核选定的留言' "
    If CheckKindPurview(3, KindID) = False And AdminPurview = 2 And AdminPurview_GuestBook >= 3 Then Response.Write "disabled"
    Response.Write ">"
    If AdminPurview = 1 Or AdminPurview_GuestBook <= 2 Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='submit' name='Submit3' value='批量移动' onClick=""document.myform.Action.value='BatchMove'"">"
    End If
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条留言", True)
    End If
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>留言搜索：</strong></td>"
    Response.Write "   <td>" & GetGuestSearch() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Function GetManagePath()
    Dim strPath
    strPath = "您现在的位置：网站留言管理&nbsp;&gt;&gt;&nbsp;"
    If KindID > 0 Then
        strPath = strPath & "<a href='Admin_GuestBook.asp?KindID=" & KindID & "'>" & KindName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If Keyword = "" Then
        If Passed = "True" Then
            strPath = strPath & "所有<font color=green>已审核</font>的留言"
        ElseIf Passed = "False" Then
            strPath = strPath & "所有<font color=blue>未审核</font>的留言"
        Else
            strPath = strPath & "所有留言"
        End If
    Else
        Select Case strField
            Case "GuestTitle"
                strPath = strPath & "留言主题中含有 <font color=red>" & Keyword & "</font> "
            Case "GuestContent"
                strPath = strPath & "留言内容含有 <font color=red>" & Keyword & "</font> "
            Case "GuestReply"
                strPath = strPath & "回复内容含有 <font color=red>" & Keyword & "</font> "
            Case "GuestName"
                strPath = strPath & "留言人姓名中含有 <font color=red>" & Keyword & "</font> "
            Case Else
                strPath = strPath & "留言主题中含有 <font color=red>" & Keyword & "</font> "
        End Select
        If Passed = "True" Then
            strPath = strPath & "并且<font color=green>已审核</font>的留言"
        ElseIf Passed = "False" Then
            strPath = strPath & "并且<font color=blue>未审核</font>的留言"
        Else
            strPath = strPath & "的留言"
        End If
    End If
    GetManagePath = strPath
End Function

Sub SetProperty()
    Dim sqlProperty, rsProperty
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定留言ID</li>"
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If InStr(GuestID, ",") > 0 Then
        sqlProperty = "select * from PE_GuestBook where GuestID in (" & GuestID & ")"
    Else
        sqlProperty = "select * from PE_GuestBook where GuestID=" & GuestID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Dim ReplyNumCount, rsReplyNum
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetPassed"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(3, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                rsProperty("GuestIsPassed") = True
                If rsProperty("TopicID") <> rsProperty("GuestID") Then
                    Dim sqlMaxId, rsMaxId, MaxId
                    sqlMaxId = "select max(GuestMaxId) as MaxId from PE_GuestBook"
                    Set rsMaxId = Conn.Execute(sqlMaxId)
                    MaxId = rsMaxId("MaxId")
                    If MaxId = "" Or IsNull(MaxId) Then MaxId = 0
                    Set rsMaxId = Nothing
                    Dim sql, rs, rsReplyNumber
                    Set rs = Server.CreateObject("adodb.recordset")
                    sql = "select top 1 * from PE_GuestBook where GuestId=" & rsProperty("TopicID")
                    rs.Open sql, Conn, 1, 3
                    If rs.EOF And rs.BOF Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>找不到发言的主题，可能已被删除！</li>"
                    Else
                        rsReplyNumber = rs("ReplyNum")
                        rs("LastReplyContent") = rsProperty("GuestContent")
                        rs("LastReplyGuest") = rsProperty("GuestName")
                        rs("LastReplyTitle") = rsProperty("GuestTitle")
                        rs("LastReplyTime") = Now()
                        rs("ReplyNum") = rsReplyNumber + 1
                        rs("GuestMaxId") = MaxId + 1
                        rs.Update
                    End If
                    rs.Close
                End If
            End If
        Case "CancelPassed"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(3, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                If rsProperty("TopicID") <> rsProperty("GuestID") Then
                    Set rsReplyNum = Conn.Execute("select count(GuestID) from PE_GuestBook where TopicID =" & rsProperty("TopicID") & " and GuestIsPassed=" & PE_True & "")
                    ReplyNumCount = rsReplyNum(0) - 2
                    Set rsReplyNum = Nothing
                    Conn.Execute ("update PE_GuestBook set ReplyNum=" & ReplyNumCount & " where GuestId=" & rsProperty("TopicID") & "")
                End If
                rsProperty("GuestIsPassed") = False
            End If
        Case "Quintessence"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(4, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                rsProperty("Quintessence") = 1
            End If
        Case "Cquintessence"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(4, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                rsProperty("Quintessence") = 0
            End If
        Case "SetOnTop"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(5, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                rsProperty("OnTop") = 1
            End If
        Case "CancelOnTop"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(5, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                rsProperty("OnTop") = 0
            End If
        Case "DelReply"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(6, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                rsProperty("GuestReply") = ""
            End If
        Case "Del"
            If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(1, rsProperty("KindID")) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对 " & rsProperty("GuestTitle") & " 没有操作权限！</li>"
            Else
                If rsProperty("TopicID") <> rsProperty("GuestID") Then
                    Dim TopicID
                    TopicID = rsProperty("TopicID")
                    rsProperty.Delete

                    Set rs = Server.CreateObject("adodb.recordset")
                    sql = "select top 1 * from PE_GuestBook where GuestId=" & TopicID
                    rs.Open sql, Conn, 1, 3
                    If Not(rs.EOF And rs.BOF) Then
                        Dim trs
                        'Set trs = Conn.Execute("select top 1 * from PE_GuestBook where TopicID=" & TopicID & " order by GuestMaxId desc")
                        Set trs = Conn.Execute("select top 1 * from PE_GuestBook where TopicID=" & TopicID & " AND GuestId<>" & TopicId &" order by GuestMaxId desc")
						If Not(trs.bof And trs.eof) then
                            rs("LastReplyContent") = trs("GuestContent")
                            rs("LastReplyGuest") = trs("GuestName")
                            rs("LastReplyTitle") = trs("GuestTitle")
                            rs("LastReplyTime") = trs("GuestDatetime")
                        Else
                            rs("LastReplyContent") = ""
                            rs("LastReplyGuest") = ""
                            rs("LastReplyTitle") = ""
                            rs("LastReplyTime") = Now()
                        End If
                        trs.close
                        Set trs = nothing

                        rs("ReplyNum") = rs("ReplyNum") - 1
                        rs.Update
                    End If
                    rs.Close
                Else
                    Conn.Execute ("delete from PE_GuestBook where TopicID=" & rsProperty("TopicID") & "")
                End If
            End If
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing
    
    Call ClearSiteCache(4)
    Call CloseConn
    If FoundErr = True Then
        Exit Sub
    End If
    Response.Redirect ComeUrl
End Sub

Sub ShowGuestList()
    Dim UserGuestName, UserType, UserSex, UserEmail, UserHomepage, UserOicq, UserIcq, UserMsn
    Dim GuestNum, GuestTip, TipName, TipSex, TipEmail, TipOicq, TipHomepage, isdelUser
    GuestNum = 0
    Call ShowTip
    Do While Not rsGuest.EOF
        isdelUser = 0
        If rsGuest("GuestType") = 1 Then
            Dim rsUser
            Set rsUser = Conn.Execute("select * from PE_Contacter where ContacterID=(select ContacterID from PE_User where UserName='" & ReplaceBadChar(rsGuest("GuestName")) & "')")
            If Not (rsUser.BOF And rsUser.EOF) Then
                UserGuestName = rsGuest("GuestName")
                UserSex = rsUser("Sex")
                UserEmail = rsUser("Email")
                UserOicq = rsUser("QQ")
                UserIcq = rsUser("ICQ")
                UserMsn = rsUser("MSN")
                UserHomepage = rsUser("Homepage")
            Else
                isdelUser = 1
            End If
        Set rsUser = Nothing
        End If
        If rsGuest("GuestType") <> 1 Or isdelUser = 1 Then
            UserGuestName = rsGuest("GuestName")
            UserSex = rsGuest("GuestSex")
            UserEmail = rsGuest("GuestEmail")
            UserOicq = rsGuest("GuestOicq")
            UserIcq = rsGuest("GuestIcq")
            UserMsn = rsGuest("GuestMsn")
            UserHomepage = rsGuest("GuestHomepage")
        End If
        TipName = UserGuestName
        If isdelUser = 1 Then TipName = TipName & "（已删除）"
        If UserEmail = "" Or IsNull(UserEmail) Then
            TipEmail = "未填"
        Else
            TipEmail = UserEmail
        End If
        If UserOicq = "" Or IsNull(UserOicq) Then
            TipOicq = "未填"
        Else
            TipOicq = UserOicq
        End If
        If UserHomepage = "" Or IsNull(UserHomepage) Then
            TipHomepage = "未填"
        Else
            TipHomepage = UserHomepage
        End If
        If UserIcq = "" Or IsNull(UserIcq) Then UserIcq = "未填"
        If UserMsn = "" Or IsNull(UserMsn) Then UserMsn = "未填"
        If UserSex = "1" Then
            TipSex = "（酷哥）"
        ElseIf UserSex = "0" Then
            TipSex = "(靓妹)"
        Else
            TipSex = ""
        End If
        GuestTip = "&nbsp;姓名：" & TipName & "&nbsp;" & TipSex & "<br>&nbsp;主页：" & TipHomepage & "<br>&nbsp;OICQ：" & TipOicq & "<br>&nbsp;信箱：" & TipEmail & "<br>&nbsp;地址：" & rsGuest("GuestIP") & "<br>&nbsp;时间：" & rsGuest("GuestDatetime")

        Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td align='center' valign='top'>" & vbCrLf
        Response.Write "            <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
        Response.Write "              <tr class='title'>" & vbCrLf
        Response.Write "                <td height='22'>" & vbCrLf
        Response.Write "                  &nbsp;&nbsp;&nbsp;&nbsp;<font color=green>主题：</font> " & rsGuest("GuestTitle") & vbCrLf
        Response.Write "                </td>" & vbCrLf
        Response.Write "                <td width='165'>" & vbCrLf
        Response.Write "                  <img src='" & GImagePath & "posttime.gif' width='11' height='11' align='absmiddle'>" & vbCrLf
        Response.Write "                  <font color='#006633'>：" & rsGuest("GuestDatetime") & "</font>" & vbCrLf
        Response.Write "                </td>" & vbCrLf
        Response.Write "              </tr>" & vbCrLf
        Response.Write "            </table>" & vbCrLf
        Response.Write "          </td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td align='center' height='153' valign='top' class='tdbg'>" & vbCrLf
        Response.Write "            <table width='100%' border='0' cellpadding='0' cellspacing='3'>" & vbCrLf
        If rsGuest("GuestIsPassed") = True Then
            Response.Write "              <tr>" & vbCrLf
        Else
            Response.Write "              <tr>" & vbCrLf
        End If
        Response.Write "                <td width='100' align='center' height='130' valign='top'>" & vbCrLf
        Response.Write "                        <img src='" & GFacePath & rsGuest("GuestImages") & ".gif' width='80' height='90' onMouseOut=toolTip() onMouseOver=""toolTip('" & GuestTip & "')""><br><br>" & vbCrLf
        If rsGuest("GuestType") = 1 Then
            Response.Write "                        <font color='#006633'>【用户】<br>" & UserGuestName & "</font>"
        Else
            Response.Write "                        【游客】<br>" & UserGuestName
        End If
        Response.Write "                </td>" & vbCrLf
        Response.Write "                <td align='center' height='153' width='1' bgcolor='#B4C9E7'>" & vbCrLf
        Response.Write "                </td>" & vbCrLf
        Response.Write "                <td>" & vbCrLf
        Response.Write "                  <table width='100%' border='0' cellpadding='6' cellspacing='0' height='125' style='TABLE-LAYOUT: fixed'>" & vbCrLf
        Response.Write "                    <tr>" & vbCrLf
        Response.Write "                      <td align='left' valign='top'><img src='" & GImagePath & "face" & rsGuest("GuestFace") & ".gif' width='19' height='19'>" & vbCrLf
        If rsGuest("GuestIsPrivate") = True Then
            Response.Write "                        <font color=green>[隐藏]</font>&nbsp;" & vbCrLf
        End If
        Response.Write FilterJS(rsGuest("GuestContent"))
        Response.Write "                      </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr>" & vbCrLf
        Response.Write "                      <td align='left' valign='bottom'>" & vbCrLf
        If rsGuest("GuestReply") <> "" Then
            Response.Write "                        <table width='100%' border='0' cellspacing='0' cellpadding='2'>" & vbCrLf
            Response.Write "                          <tr>" & vbCrLf
            Response.Write "                            <td height='1' bgcolor='#B4C9E7'></td>" & vbCrLf
            Response.Write "                          </tr>" & vbCrLf
            Response.Write "                          <tr>" & vbCrLf
            Response.Write "                            <td valign='top'>" & vbCrLf
            Response.Write "                              <table width='100%' border='0' cellpadding='0' cellspacing='0' style='TABLE-LAYOUT: fixed'>" & vbCrLf
            Response.Write "                                <tr>" & vbCrLf
            Response.Write "                                  <td><font color='#006633'> 管理员<font color='#FF0000'>[" & rsGuest("GuestReplyAdmin") & "]</font>回复:" & vbCrLf & rsGuest("GuestReplyDatetime") & "</font></td>" & vbCrLf
            Response.Write "                                </tr>" & vbCrLf
            Response.Write "                                <tr>" & vbCrLf
            Response.Write "                                  <td valign='bottom'><font color='#006633'>" & rsGuest("GuestReply") & "</font></td>" & vbCrLf
            Response.Write "                                </tr>" & vbCrLf
            Response.Write "                              </table>" & vbCrLf
            Response.Write "                            </td>" & vbCrLf
            Response.Write "                          </tr>" & vbCrLf
            Response.Write "                        </table>" & vbCrLf
        End If
        Response.Write "                      </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                  </table>" & vbCrLf
        Response.Write "                  <table width='100%' height='1' border='0' cellpadding='0' cellspacing='0' bgcolor='#B4C9E7'>" & vbCrLf
        Response.Write "                    <tr>" & vbCrLf
        Response.Write "                      <td></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                  </table>" & vbCrLf
        Response.Write "                  <table width=100% border=0 cellpadding=0 cellspacing=3>" & vbCrLf
        Response.Write "                    <tr>" & vbCrLf
        Response.Write "                      <td>" & vbCrLf
        If UserHomepage = "" Or IsNull(UserHomepage) Then
            Response.Write "<img src=" & GImagePath & "nourl.gif width=45 height=16 alt=" & UserGuestName & "没有留下主页地址 border=0>" & vbCrLf
        Else
            Response.Write "<a href=" & UserHomepage & " target=""_blank"">"
            Response.Write "<img src=" & GImagePath & "url.gif width=45 height=16 alt=" & UserHomepage & " border=0></a>" & vbCrLf
        End If
        If UserOicq = "" Or IsNull(UserOicq) Then
            Response.Write "<img src=" & GImagePath & "nooicq.gif width=45 height=16 alt=" & UserGuestName & "没有留下QQ号码 border=0>" & vbCrLf
        Else
            Response.Write "<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln=" & UserOicq & " target='_blank'>"
            Response.Write "<img src=" & GImagePath & "oicq.gif width=45 height=16 alt=" & UserOicq & " border=0 ></a>" & vbCrLf
        End If
        If UserEmail = "" Or IsNull(UserEmail) Then
            Response.Write "<img src=" & GImagePath & "noemail.gif width=45 height=16 alt=" & UserGuestName & "没有留下Email地址 border=0>" & vbCrLf
        Else
            Response.Write "<a href=mailto:" & UserEmail & ">"
            Response.Write "<img src=" & GImagePath & "email.gif width=45 height=16 border=0 alt=" & UserEmail & "></a>" & vbCrLf
        End If
        Response.Write "<img src=" & GImagePath & "other.gif width=45 height=16 border=0 onMouseOut=toolTip() onMouseOver=""toolTip('&nbsp;Icq：" & UserIcq & "<br>&nbsp;Msn：" & UserMsn & "<br>&nbsp;I P：" & rsGuest("GuestIP") & "')"">" & vbCrLf
        Response.Write "<a href=" & FileName & "?action=reply&guestid=" & rsGuest("GuestId") & ">"
        Response.Write "                      </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                  </table>" & vbCrLf
        Response.Write "                </td>" & vbCrLf
        Response.Write "              </tr>" & vbCrLf
        Response.Write "            </table>" & vbCrLf
        Response.Write "          </td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "      </table>" & vbCrLf
        Response.Write "      <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td class='main_shadow'>" & vbCrLf
        Response.Write "          </td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "      </table>" & vbCrLf
        rsGuest.MoveNext
        GuestNum = GuestNum + 1
        If GuestNum >= MaxPerPage Then Exit Do
    Loop
End Sub

Sub ShowJS_Guest()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function changeimage()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  document.myform.GuestImages.value=document.myform.Image.value;" & vbCrLf
    Response.Write "  document.myform.showimages.src='" & GFacePath & "'+document.myform.Image.value+'.gif';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function guestpreview()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  document.preview.content.value=document.myform.GuestContent.value;" & vbCrLf
    Response.Write "  var popupWin = window.open('GuestPreview.asp', 'GuestPreview', 'scrollbars=yes,width=620,height=230');" & vbCrLf
    Response.Write "  document.preview.submit();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(document.myform.GuestName.value==''){" & vbCrLf
    Response.Write "    alert('姓名不能为空！');" & vbCrLf
    Response.Write "    document.myform.GuestName.focus();" & vbCrLf
    Response.Write "    return(false) ;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.GuestTitle.value==''){" & vbCrLf
    Response.Write "    alert('留言主题不能为空！');" & vbCrLf
    Response.Write "    document.myform.GuestTitle.focus();" & vbCrLf
    Response.Write "    return(false);" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.GuestTitle.value.length>30){" & vbCrLf
    Response.Write "    alert('留言主题不能超过30字符！');" & vbCrLf
    Response.Write "    document.myform.GuestTitle.focus();" & vbCrLf
    Response.Write "    return(false);" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.GuestContent.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "    document.myform.GuestContent.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.GuestContent.value==''){" & vbCrLf
    Response.Write "    alert('留言内容不能为空！');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "    return(false);" & vbCrLf
    Response.Write "  }" & vbCrLf
    
    Response.Write "  if(document.myform.GuestContent.value.length>65536){" & vbCrLf
    Response.Write "    alert('留言内容不能超过64K！');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "    return(false);" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "</script>" & vbCrLf
End Sub

Sub Modify()
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的留言ID！</li>"
        Exit Sub
    Else
        GuestID = PE_CLng(GuestID)
    End If
    sql = "select * from PE_GuestBook where GuestID=" & GuestID
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的留言！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(0, rs("KindID")) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对 " & rs("GuestTitle") & " 没有操作权限！</li>"
        Exit Sub
    End If
    
    Call ShowJS_Guest
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><b>修改留言</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='0'>"
    Response.Write "        <form name='myform' method='post' action='Admin_GuestBook.asp' onSubmit='return CheckForm()'>" & vbCrLf
    If rs("GuestType") = 0 Then
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td width='30%' align='right'>姓 &nbsp;名： </td>" & vbCrLf
        Response.Write "            <td width='30%'>" & vbCrLf
        Response.Write "              <input type='text' name='GuestName' maxlength='14' size='20' value='" & rs("GuestName") & "'>"
        Response.Write "              <font color=red>*</font>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "            <td width='22%'>&nbsp; </td>" & vbCrLf
        Response.Write "            <td colspan='2'>&nbsp; </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='right'>性&nbsp;&nbsp;别： </td>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Response.Write "              <input type='radio' name='GuestSex' value='1' "
        If rs("GuestSex") = "1" Then Response.Write " checked"
        Response.Write " style='BORDER:0px;'>" & vbCrLf
        Response.Write "              男&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "              <input type='radio' name='GuestSex' value='0' "
        If rs("GuestSex") = "0" Then Response.Write " checked"
        Response.Write " style='BORDER:0px;'>" & vbCrLf
        Response.Write "              女 </td>" & vbCrLf
        Response.Write "            <td>&nbsp;&nbsp;" & vbCrLf
        Response.Write "              <select name='Image' size='1' onChange='changeimage();' >" & vbCrLf
        For i = 1 To 9
            Response.Write "                <option value='0" & i & "'>0" & i & "</option>" & vbCrLf
        Next
        For i = 10 To 23
            Response.Write "                  <option value='" & i & "'>" & i & "</option>" & vbCrLf
        Next
        Response.Write "              </select>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "            <td colspan='2'>&nbsp;</td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='right'>E-mail： </td>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Response.Write "              <input type='text' name='GuestEmail' maxlength='30' size='20' value='" & rs("GuestEmail") & "'>"
        Response.Write "            </td>" & vbCrLf
        Response.Write "            <td rowspan='4'>" & vbCrLf
        Response.Write "              <input type='hidden' name='GuestImages' value='01'>" & vbCrLf
        Response.Write "              <img name=showimages src='" & GFacePath & rs("GuestImages") & ".gif' width='80' height='90' border='0' onClick=window.open('../guestbook/guestselect.asp?action=face','face','width=480,height=400,resizable=1,scrollbars=1') title=点击选择头像 style='cursor:hand'>" & vbCrLf
        Response.Write "              </td>" & vbCrLf
        Response.Write "            <td colspan='2' rowspan='4'>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='right'>Oicq： </td>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Response.Write "              <input type='text' name='GuestOicq' maxlength='15' size='20' value='" & rs("GuestOicq") & "'>"
        Response.Write "            </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='right'>Icq： </td>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Response.Write "              <input type='text' name='GuestIcq' maxlength='15' size='20' value='" & rs("GuestIcq") & "'>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='right'>Msn： </td>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Response.Write "              <input type='text' name='GuestMsn' maxlength='40' size='20' value='" & rs("GuestMsn") & "'>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='right'>个人主页： </td>" & vbCrLf
        Response.Write "            <td colspan='4'>" & vbCrLf
        Response.Write "              <input type='text' name='GuestHomepage' maxlength='80' size='37' value='" & rs("GuestHomepage") & "'>" & vbCrLf
        Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        'Response.Write "          <tr class='tdbg'>" & vbCrLf
        'Response.Write "            <td align='center'></td>" & vbCrLf
        'Response.Write "            <td colspan='4'>&nbsp; </td>" & vbCrLf
        'Response.Write "          </tr>" & vbCrLf
    Else
        Response.Write "          <tr class='tdbg'>" & vbCrLf
        Response.Write "            <td align='center'>选择头像： </td>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Response.Write "              <input type='hidden' name='GuestName'  value='" & rs("GuestName") & "'>" & vbCrLf
        Response.Write "              <input type='hidden' name='reg' value='1'>" & vbCrLf
        Response.Write "              <input type='hidden' name='GuestImages' value='" & rs("GuestImages") & "'>" & vbCrLf
        Response.Write "              <img name=showimages src='" & GFacePath & rs("GuestImages") & ".gif' width='80' height='90' border='0' onClick=window.open('guestselect.asp?action=face','face','width=480,height=400,resizable=1,scrollbars=1') title=点击选择头像 style='cursor:hand'>" & vbCrLf
        Response.Write "              <select name='Image' size='1' onChange='changeimage();'>" & vbCrLf
        For i = 1 To 9
          Response.Write "                <option value='0" & i & "'>0" & i & "</option>" & vbCrLf
        Next
        For i = 10 To 23
        Response.Write "                  <option value='" & i & "'>" & i & "</option>" & vbCrLf
        Next
        Response.Write "              </select>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "            <td>&nbsp;</td>" & vbCrLf
        Response.Write "            <td>&nbsp;</td>" & vbCrLf
        Response.Write "            <td>&nbsp;</td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td align='right'>留言主题： </td>" & vbCrLf
    Response.Write "            <td colspan='4'>" & vbCrLf
    Response.Write "              <input type='text' name='GuestTitle' size='37' maxlength='21' value='" & rs("GuestTitle") & "'>" & vbCrLf
    Response.Write "              <font color=red>*</font>" & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td align='right'>留言类别：</td>" & vbCrLf
    Response.Write "            <td colspan='4'><select name='KindID' id='KindID'>" & GetGKind_Option(1, rs("KindID")) & "</select>" & vbCrLf
    Response.Write "            </td></tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td align='right'>现在心情： </td>" & vbCrLf
    Response.Write "            <td colspan='4'>" & vbCrLf
    For i = 1 To 30
        Response.Write "<input type='radio' name='GuestFace' value='" & i & "'"
        If i = PE_CLng(rs("GuestFace")) Then Response.Write " checked"
        Response.Write " style='BORDER:0px;width:19;'>"
        Response.Write "<img src='" & GImagePath & "face" & i & ".gif' width='19' height='19'>" & vbCrLf
        If i Mod 10 = 0 Then Response.Write "<br>"
    Next
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf

    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td valign='middle' align='right'>留言内容：  <br>" & vbCrLf
    Response.Write "              </td>" & vbCrLf
    Response.Write "            <td colspan='4' valign='top'>" & vbCrLf
    'Response.Write "              <textarea name='GuestContent' cols='59' rows='6'    onkeydown=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain); onkeyup=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain);>" & rs("GuestContent") & "</textarea>" & vbCrLf
    Response.Write "              <textarea name='GuestContent' id='GuestContent' style='display:none' >" & Server.HTMLEncode(rs("GuestContent")) & "</textarea>" & vbCrLf
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=GuestContent' frameborder='1' scrolling='no' width='480' height='280' ></iframe>" & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td valign='middle' align='center'></td>" & vbCrLf
    Response.Write "            <td colspan='4' valign='top'>" & vbCrLf
    Response.Write "                <FONT color=green>小提示：</FONT>换行请按Shift+Enter,另起一段请按Enter " & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td valign='middle' align='right'>是否隐藏： </td>" & vbCrLf
    Response.Write "            <td colspan='4' valign='top'>" & vbCrLf
    Response.Write "              <input type='radio' name='GuestIsPrivate' value='no' "
    If rs("GuestIsPrivate") = False Then Response.Write " checked"
    Response.Write " style='BORDER:0px;'>" & vbCrLf
    Response.Write "              正常" & vbCrLf
    Response.Write "              <input type='radio' name='GuestIsPrivate' value='yes' "
    If rs("GuestIsPrivate") = True Then Response.Write " checked"
    Response.Write "    style='BORDER:0px;'>" & vbCrLf
    Response.Write "              隐藏 &nbsp;&nbsp;<font color=#009900>*</font> 选择隐藏后，此留言只有管理员和留言者才可以看到。</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td colspan='5' align='center'  height='40'>" & vbCrLf
    Response.Write "              <input type='hidden' name='GuestID'  value='" & GuestID & "'>"
    Response.Write "              <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "              <input name='Save' type='submit' value='保存修改结果' style='cursor:hand;'>&nbsp;"
    Response.Write "              <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_GuestBook.asp';"" style='cursor:hand;'>"
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </form>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    
    rs.Close
    Set rs = Nothing
End Sub

Sub SaveModify()
    Dim GuestName, GuestSex, GuestOicq, GuestEmail, GuestHomepage, GuestFace, GuestImages, GuestIcq, GuestMsn
    Dim GuestTitle, GuestContent, GuestIsPrivate, GuestIsPassed
    Dim GuestPassword, GuestPasswordConfirm, GuestQuestion, GuestAnswer
    Dim sqlMaxId, rsMaxId, MaxId, Saveinfo, sqlReg, rsReg
    
    KindID = Trim(Request("KindID"))
    KindID = PE_CLng(KindID)
    
    If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(0, KindID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你没有操作权限！</li>"
        Exit Sub
    End If
    
    GuestContent = FilterJS(Request("GuestContent"))
    'If UserLogined = False Then
        GuestName = PE_HTMLEncode(Trim(Request("GuestName")))
        GuestSex = Trim(Request("GuestSex"))
        GuestOicq = PE_HTMLEncode(Trim(Request("GuestOicq")))
        GuestIcq = PE_HTMLEncode(Trim(Request("GuestIcq")))
        GuestMsn = PE_HTMLEncode(Trim(Request("GuestMsn")))
        GuestEmail = PE_HTMLEncode(Trim(Request("GuestEmail")))
        GuestHomepage = PE_HTMLEncode(Trim(Request("GuestHomepage")))
        If GuestHomepage = "http://" Or IsNull(GuestHomepage) Then GuestHomepage = ""
    'Else
    '    GuestName = UserName
    'End If
    GuestImages = Trim(Request("GuestImages"))
    GuestFace = Trim(Request("GuestFace"))
    GuestTitle = PE_HTMLEncode(Trim(Request("GuestTitle")))
    GuestIsPrivate = Trim(Request("GuestIsPrivate"))
    If GuestIsPrivate = "yes" Then
        GuestIsPrivate = True
    Else
        GuestIsPrivate = False
    End If
        
    If GuestName = "" Or GuestTitle = "" Or GuestContent = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>留言保存失败！</li><li>请将必要的信息填写完整！</li>"
        Exit Sub
    End If

    GuestID = Request("GuestID")
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要编辑的留言ID！</li>"
        Exit Sub
    Else
        GuestID = PE_CLng(GuestID)
        sqlMaxId = "select max(GuestMaxId) as MaxId from PE_GuestBook"
        Set rsMaxId = Conn.Execute(sqlMaxId)
        MaxId = rsMaxId("MaxId")
        If MaxId = "" Or IsNull(MaxId) Then MaxId = 0
        Set rsMaxId = Nothing
        Set rsGuest = Server.CreateObject("adodb.recordset")
        sql = "select * from PE_GuestBook where GuestID=" & GuestID
        rsGuest.Open sql, Conn, 1, 3
        rsGuest("KindID") = KindID
        rsGuest("GuestName") = GuestName
        rsGuest("GuestSex") = GuestSex
        rsGuest("GuestOicq") = GuestOicq
        rsGuest("GuestIcq") = GuestIcq
        rsGuest("GuestMsn") = GuestMsn
        rsGuest("GuestEmail") = GuestEmail
        rsGuest("GuestHomepage") = GuestHomepage
        rsGuest("GuestTitle") = GuestTitle
        rsGuest("GuestFace") = GuestFace
        rsGuest("GuestContent") = GuestContent
        rsGuest("GuestImages") = GuestImages
        rsGuest("GuestMaxId") = MaxId + 1
        rsGuest("GuestIsPrivate") = GuestIsPrivate
        rsGuest.Update
    End If
    Call ClearSiteCache(4)
    Call CloseConn
    Response.Redirect "Admin_GuestBook.asp"
End Sub

Sub Show()
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要显示的留言ID！</li>"
        Exit Sub
    Else
        GuestID = PE_CLng(GuestID)
    End If
    sql = "select * from PE_GuestBook where GuestID=" & GuestID
    Set rsGuest = Server.CreateObject("adodb.recordset")
    rsGuest.Open sql, Conn, 1, 1
    If rsGuest.BOF And rsGuest.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的留言！</li>"
        rsGuest.Close
        Set rsGuest = Nothing
        Exit Sub
    End If
    Response.Write "<br>"
    Call ShowGuestList
End Sub

Sub AdminReply()
    Dim GuestReply, ReplyIsPrivate
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的留言ID！</li>"
        Exit Sub
    Else
        GuestID = PE_CLng(GuestID)
    End If
    sql = "select * from PE_GuestBook where GuestID=" & GuestID
    Set rsGuest = Server.CreateObject("adodb.recordset")
    rsGuest.Open sql, Conn, 1, 1
    If rsGuest.BOF And rsGuest.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的留言！</li>"
        rsGuest.Close
        Set rsGuest = Nothing
        Exit Sub
    End If
    
    If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(6, rsGuest("KindID")) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你没有操作权限！</li>"
        Exit Sub
    End If
    
    GuestReply = rsGuest("GuestReply")
    ReplyIsPrivate = rsGuest("ReplyIsPrivate")
    Response.Write "<br>"
    Call ShowGuestList

    Response.Write "<script language=JavaScript>" & vbCrLf
    Response.Write "function check(thisform)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.GuestContent.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "    document.myform.GuestContent.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.GuestContent.value==''){" & vbCrLf
    Response.Write "    alert('留言内容不能为空！');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "    return(false);" & vbCrLf
    Response.Write "  }" & vbCrLf
    
    Response.Write "   if(thisform.GuestContent.value.length>800){" & vbCrLf
    Response.Write "        alert('留言内容不能超过800字符！');" & vbCrLf
    Response.Write "        thisform.GuestContent.focus();" & vbCrLf
    Response.Write "          return(false);" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br><table width='100%' cellpadding='1' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <form name='myform' method='post' action='Admin_GuestBook.asp?action=SaveAdminReply' onSubmit='return check(myform)'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td  height='22' colspan='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color=green>回复留言</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='right'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td colspan='2'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='20%' valign='middle' align='right'>回复内容：  </td>" & vbCrLf
    Response.Write "      <td colspan='2' valign='top'>" & vbCrLf
    'Response.Write "        <textarea name='GuestContent' cols='59' rows='6' >" & GuestReply & "</textarea>"
    Response.Write "        <textarea name='GuestContent' id='GuestContent' style='display:none' >" & Server.HTMLEncode(FilterJS(GuestReply)) & "</textarea>" & vbCrLf

    Response.Write "          <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=GuestContent' frameborder='1' scrolling='no' width='480' height='280' ></iframe>" & vbCrLf
    Response.Write "     </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td width='20%' valign='middle' align='right'>是否隐藏：</td>" & vbCrLf
    Response.Write "            <td vAlign=top colSpan=2>" & vbCrLf
    Response.Write "  <Input style='BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px' type=radio name='ReplyIsPrivate' value='0' " & IsRadioChecked(ReplyIsPrivate, False) & "> 正常 " & vbCrLf

    Response.Write "              <Input style='BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px' type=radio name='ReplyIsPrivate' value='1' " & IsRadioChecked(ReplyIsPrivate, True) & "> 隐藏 <FONT color=red>*</FONT> <FONT color=green>选择隐藏后，此留言只有管理员和留言者才可以看到。</FONT></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td colspan='3' align='center'  height='40'><input name='GuestID' type='hidden' value='" & GuestID & "'>" & vbCrLf
    Response.Write "        <input type='submit' name='Submit' value=' 回 复 ' >" & vbCrLf
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='reset' name='Submit2' value=' 重 填 ' >" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "     </tr>" & vbCrLf
    Response.Write "  </form>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    
    rsGuest.Close
    Set rsGuest = Nothing
End Sub

Sub SaveAdminReply()
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不能确定留言ID</li>"
        Exit Sub
    Else
        GuestID = PE_CLng(GuestID)
    End If
    Dim GuestReply, ReplyIsPrivate
    Dim sqlMaxId, rsMaxId, MaxId
    GuestReply = FilterJS(Request("GuestContent"))
    ReplyIsPrivate = CBool(Trim(Request("ReplyIsPrivate")))
    
    sqlMaxId = "select max(GuestMaxId) as MaxId from PE_GuestBook"
    Set rsMaxId = Conn.Execute(sqlMaxId)
    MaxId = rsMaxId("MaxId")
    If MaxId = "" Or IsNull(MaxId) Then MaxId = 0
    Set rsMaxId = Nothing
    
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_GuestBook where GuestID=" & GuestID
    rs.Open sql, Conn, 1, 3
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的留言！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If AdminPurview = 2 And AdminPurview_GuestBook = 3 And CheckKindPurview(6, rs("KindID")) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你没有操作权限！</li>"
        Exit Sub
    End If
    
    rs("GuestMaxId") = MaxId + 1
    rs("GuestReply") = GuestReply
    rs("GuestReplyAdmin") = AdminName
    rs("GuestReplyDatetime") = Now()
    rs("ReplyIsPrivate") = ReplyIsPrivate
    rs.Update
    rs.Close
    Set rs = Nothing
    Call CloseConn
    Response.Redirect "Admin_GuestBook.asp"
End Sub

Function GetGuestSearch()
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='Admin_GuestBook.asp'>"
    strForm = strForm & "<tr><td height='28' align='center'>"
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='GuestTitle' selected>留言主题</option>"
    strForm = strForm & "<option value='GuestContent'>留言内容</option>"
    strForm = strForm & "<option value='GuestReply'>回复内容</option>"
    strForm = strForm & "<option value='GuestName'>留言人</option>"
    strForm = strForm & "</select>"
    strForm = strForm & " <input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    strForm = strForm & "<input type='submit' name='Submit'  value='搜索'>"
    strForm = strForm & "</td></tr></form></table>"
    GetGuestSearch = strForm
End Function

Sub GKind()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim KindID, rsGKind, sqlGKind
    sqlGKind = "select * from PE_Guestkind order by OrderID"
    Set rsGKind = Conn.Execute(sqlGKind)

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td width='50' align='center'><strong>类别ID</strong></td>"
    Response.Write "    <td width='150' align='center'><strong>类别名称</strong></td>"
    Response.Write "    <td align='center'><strong>类别说明</strong></td>"
    Response.Write "    <td width='150' align='center'><strong>常规操作</strong></td>"
    Response.Write "    <td width='100' align='center'><strong>排序操作</strong></td>" & vbCrLf
    Response.Write "  </tr>"
    If rsGKind.BOF And rsGKind.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='5' align='center'>您还没有添加任何留言类别!</td><tr>" & vbCrLf
    Else
        Do While Not rsGKind.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='50' align='center'>" & rsGKind("KindID") & "</td>"
            Response.Write "    <td width='150' align='center'><a href='Admin_GuestBook.asp?KindID=" & rsGKind("KindID") & "' title='点击进入管理此类别的留言'>" & rsGKind("KindName") & "</a></td>"
            Response.Write "    <td>" & PE_HTMLEncode(rsGKind("ReadMe")) & "</td>"
            Response.Write "    <td width='150' align='center'>"
            Response.Write "<a href='Admin_GuestBook.asp?action=ModifyGKind&KindID=" & rsGKind("KindID") & "'>修改</a>&nbsp;&nbsp;"
            Response.Write "<a href='Admin_GuestBook.asp?Action=DelGKind&KindID=" & rsGKind("KindID") & "' onClick=""return confirm('确定要删除此类别吗？删除此类别后原属于此类别的留言将不属于任何类别。');"">删除</a>&nbsp;&nbsp;"
            Response.Write "<a href='Admin_GuestBook.asp?Action=ClearGKind&KindID=" & rsGKind("KindID") & "' onClick=""return confirm('确定要清空此类别中的留言吗？本操作将原属于此类别的留言改为不属于任何类别。');"">清空</a>"
            Response.Write "</td>"
            Response.Write "<form name='orderform' method='post' action='Admin_GuestBook.asp'>"
            Response.Write "    <td width='100' align='center'>      <input name='OrderID' type='text' id='OrderID' value='" & rsGKind("OrderID") & "' size='4' maxlength='4' style='text-align:center '>"
            Response.Write "      <input name='KindID' type='hidden' id='KindID' value='" & rsGKind("KindID") & "'>"
            Response.Write "    <input type='submit' name='Submit' value='修改'>"
            Response.Write "    <input name='Action' type='hidden' id='Action' value='OrderGuestKind'></td></form>"
            Response.Write "</tr>"
            rsGKind.MoveNext
        Loop
    End If
    Response.Write "</table>"
    rsGKind.Close
    Set rsGKind = Nothing
End Sub

Sub AddGKind()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    Response.Write "<form method='post' action='Admin_GuestBook.asp' name='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>添加留言类别</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>类别名称：</strong></td>"
    Response.Write "      <td class='tdbg'><input name='KindName' type='text' id='KindName' size='49' maxlength='30'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>类别说明</strong><br>鼠标移至类别名称上时将显示设定的说明文字（不支持HTML）</td>"
    Response.Write "      <td class='tdbg'><textarea name='ReadMe' cols='40' rows='5' id='ReadMe'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAddGKind'>"
    Response.Write "        <input  type='submit' name='Submit' value=' 添 加 '>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_GuestBook.asp'"" style='cursor:hand;'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyGKind()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim KindID, rsGKind, sqlGKind
    KindID = Trim(Request("KindID"))
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的类别ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    sqlGKind = "Select * from PE_Guestkind Where KindID=" & KindID
    Set rsGKind = Server.CreateObject("Adodb.RecordSet")
    rsGKind.Open sqlGKind, Conn, 1, 3
    If rsGKind.BOF And rsGKind.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的类别，可能已经被删除！</li>"
    Else
        Response.Write "<form method='post' action='Admin_GuestBook.asp' name='myform'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>修改留言类别</strong></div></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg'><strong>类别名称：</strong></td>"
        Response.Write "      <td class='tdbg'><input name='KindName' type='text' id='KindName' value='" & rsGKind("KindName") & "' size='49' maxlength='30'><input name='KindID' type='hidden' id='KindID' value='" & rsGKind("KindID") & "'></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg'><strong>类别说明</strong><br>鼠标移至类别名称上时将显示设定的说明文字（不支持HTML）</td>"
        Response.Write "      <td class='tdbg'><textarea name='ReadMe' cols='40' rows='5' id='ReadMe'>" & rsGKind("ReadMe") & "</textarea></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModifyGKind'>"
        Response.Write "        <input  type='submit' name='Submit' value='保存修改结果'>&nbsp;&nbsp;"
        Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_GuestBook.asp'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsGKind.Close
    Set rsGKind = Nothing
End Sub

Sub DelGKind()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim KindID
    KindID = Trim(Request("KindID"))
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的类别ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    If FoundErr = True Then Exit Sub

    If Action = "DelGKind" Then
        Conn.Execute ("delete from PE_Guestkind where KindID=" & KindID)
    End If
    Conn.Execute ("update PE_GuestBook set KindID=0 where KindID=" & KindID)
    Call ClearSiteCache(4)
    Call CloseConn
    Response.Redirect "Admin_GuestBook.asp?Action=GKind"
End Sub

Sub SaveGKind()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim KindID, KindName, ReadMe, rs, mrs, intMaxID, OrderID
    KindName = ReplaceBadChar(Trim(Request("KindName")))
    ReadMe = Trim(Request("ReadMe"))
    If KindName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>类别名称不能为空！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If Action = "SaveAddGKind" Then
        Set mrs = Conn.Execute("select max(KindID) from PE_Guestkind")
        If IsNull(mrs(0)) Then
            intMaxID = 0
        Else
            intMaxID = mrs(0)
        End If
        Set mrs = Nothing
        
        Set mrs = Conn.Execute("select max(OrderID) from PE_Guestkind")
        If IsNull(mrs(0)) Then
            OrderID = 1
        Else
            OrderID = mrs(0) + 1
        End If
        Set mrs = Nothing
        
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open "Select * from PE_Guestkind Where KindName='" & KindName & "'", Conn, 1, 3
        If Not (rs.BOF And rs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>类别名称已经存在！</li>"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.addnew
        rs("KindID") = intMaxID + 1
        rs("OrderID") = OrderID
    ElseIf Action = "SaveModifyGKind" Then
        KindID = Trim(Request("KindID"))
        If KindID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的类别ID！</li>"
            Exit Sub
        Else
            KindID = PE_CLng(KindID)
        End If
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open "Select * from PE_Guestkind Where KindID=" & KindID, Conn, 1, 3
        If rs.BOF And rs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的类别，可能已经被删除！</li>"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If
    rs("KindName") = KindName
    rs("ReadMe") = ReadMe
    rs.Update
    rs.Close
    Set rs = Nothing
    Call ClearSiteCache(4)
    Call CloseConn
    Response.Redirect "Admin_GuestBook.asp?Action=GKind"
End Sub

Sub OrderGuestKind()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim KindID, OrderID
    KindID = Trim(Request("KindID"))
    OrderID = Trim(Request("OrderID"))
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定留言类别ID</li>"
    Else
        KindID = PE_CLng(KindID)
    End If
    If OrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定排序顺序ID</li>"
    Else
        OrderID = PE_CLng(OrderID)
    End If
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_Guestkind set OrderID=" & OrderID & " where KindID=" & KindID & "")
    Call ClearSiteCache(4)
    Call CloseConn
    Response.Redirect "Admin_GuestBook.asp?Action=GKind"
End Sub

Function GetGKindList()
    Dim rsGKind, sqlGKind, strGKind, i
    sqlGKind = "select * from PE_Guestkind order by OrderID"
    Set rsGKind = Conn.Execute(sqlGKind)
    If rsGKind.BOF And rsGKind.EOF Then
        strGKind = strGKind & "没有任何类别"
    Else
        i = 1
        strGKind = "| "
        Do While Not rsGKind.EOF
            If rsGKind("KindID") = KindID Then
                strGKind = strGKind & "<a href='Admin_GuestBook.asp?KindID=" & KindID & "'><font color=red>" & rsGKind("KindName") & "</font></a>"
            Else
                strGKind = strGKind & "<a href='Admin_GuestBook.asp?KindID=" & rsGKind("KindID") & "'>" & rsGKind("KindName") & "</a>"
            End If
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
    GetGKindList = strGKind
End Function

Function GetGKind_Option(ShowType, KindID)
    Dim sqlGKind, rsGKind, strOption
    If ShowType = 3 Then
        strOption = ""
    Else
        strOption = "<option value=''"
        If KindID = 0 Then
            strOption = strOption & " selected"
        End If
        strOption = strOption & ">不属于任何类别</option>"
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
'函数名：TransformTime()
'作  用：格式化时间
'参  数：时间
'=================================================
Function TransformTime(GuestDatetime)
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
       dayshow = "今天 "
    ElseIf dnt = 1 Then
       dayshow = "昨天 "
    ElseIf dnt = 2 Then
       dayshow = "前天 "
    End If
    TransformTime = dayshow & pshow & thour & ":" & tminute
End Function

'=================================================
'函数名：Move()
'作  用：移动主题到其他至类别的操作
'参  数：无
'=================================================
Sub Move()
    Dim arrKindID
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要移动的留言！</li>"
        Exit Sub
    End If
    GuestID = PE_CLng(GuestID)
    Set arrKindID = Conn.Execute("select top 1 KindID from PE_GuestBook where GuestID=" & GuestID & " order by GuestID asc ")
    If arrKindID.EOF Or arrKindID.BOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要移动的留言！</li>"
        Exit Sub
    Else
        If AdminPurview = 2 And AdminPurview_GuestBook > 2 And CheckKindPurview(2, arrKindID(0)) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
            Exit Sub
        End If
    End If
    Set arrKindID = Nothing

    sqlGuest = "select B.GuestTitle,K.KindID,K.KindName from PE_GuestBook B"
    sqlGuest = sqlGuest & " left join PE_GuestKind K on B.KindID=K.KindID where B.GuestID=" & GuestID
    Set rsGuest = Conn.Execute(sqlGuest)
    If rsGuest.BOF And rsGuest.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到要移到的留言</li>"
    Else
        Response.Write "<form name='myform' method='post' action='Admin_GuestBook.asp'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
        Response.Write "    <tr class='title'>"
        Response.Write "      <td height='22' align='center'><strong>留言移动</strong></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td>"
        Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='200'><strong>所属类别：</strong></td>"
        If rsGuest("KindName") = "" Or IsNull(rsGuest("KindName")) Then
            Response.Write "            <td>不属于任何类别</td>"
        Else
            Response.Write "            <td>" & rsGuest("KindName") & "</td>"
        End If
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='200'><strong>留言主题：</strong></td>"
        Response.Write "            <td>" & rsGuest("GuestTitle") & "<input name='GuestID' type='hidden' id='GuestID' value='" & GuestID & "'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='200'><strong>移动到：</strong></td>"
        Response.Write "            <td><select name='TargetKindID' size='2'  style='height:300px;width:400px;'>" & GetKind_Option(rsGuest("KindID")) & "</select></td>"
        Response.Write "          </tr>"
        Response.Write "        </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='center'>"
        Response.Write "        <input name='strComeUrl' type='hidden' id='strComeUrl' value='" & ComeUrl & "'>"
        Response.Write "        <input name='Action' type='hidden' id='Action' value='MoveGuest'>"
        Response.Write "        <input type='submit' name='Submit' value=' 确 定 '>&nbsp;"
        Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_GuestBook.asp'"" style='cursor:hand;'>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsGuest.Close
    Set rsGuest = Nothing
    Call ClearSiteCache(4)
End Sub

'=================================================
'函数名：MoveGuest()
'作  用：移动主题到其他至类别的保存操作
'参  数：无
'=================================================
Sub MoveGuest()
    Dim arrKindID
    If GuestID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要移动的留言！</li>"
        Exit Sub
    End If
    GuestID = PE_CLng(GuestID)
    Set arrKindID = Conn.Execute("select top 1 KindID from PE_GuestBook where GuestID=" & GuestID & " order by GuestID asc ")
    If arrKindID.EOF Or arrKindID.BOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要移动的留言！</li>"
        Exit Sub
    Else
        If AdminPurview = 2 And AdminPurview_GuestBook > 2 And CheckKindPurview(2, arrKindID(0)) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
            Exit Sub
        End If
    End If
    Set arrKindID = Nothing
    
    Dim strComeUrl, TargetKindID
    strComeUrl = Trim(Request("strComeUrl"))
    TargetKindID = PE_CLng(Trim(Request("TargetKindID")))
   ' Call CheckKindPurview("saveadd", TargetKindID)
    
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_GuestBook set KindID=" & TargetKindID & " where GuestID in (select GuestID from PE_GuestBook where TopicID=" & GuestID & ")")

    Call CloseConn
    Response.Redirect strComeUrl
End Sub

'=================================================
'函数名：BatchMove()
'作  用：批量移动主题到其他至类别的操作
'参  数：无
'=================================================
Sub BatchMove()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    Dim BatchGuestID
    BatchGuestID = ReplaceBadChar(Request("GuestID"))
    
    Response.Write "<form method='POST' name='myform' action='Admin_GuestBook.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='4' align='center'><b>批量移动留言</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td vlign='top' width='300'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='GuestType' value='1' checked>指定留言ID：<input type='text' name='BatchGuestID' value='" & BatchGuestID & "' size='30'><br>"
    Response.Write "              <input type='radio' name='GuestType' value='2'>指定类别的留言：<br><select name='BatchKindID' size='2' multiple style='height:360px;width:300px;'>" & GetKind_Option(-1) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  选定所有栏目  ' onclick='SelectAll()'>"
    Response.Write "      <input type='button' name='Submit' value='取消选定所有栏目' onclick='UnSelectAll()'></div></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td>移动到&gt;&gt;</td>"
    Response.Write "      <td valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td><br>"

    Response.Write "              目标类别：<br><select name='tKindID' size='2' style='height:360px;width:300px;'>" & GetKind_Option(0) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='BatchMove'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 执行批处理 ' style='cursor:hand;' onClick=""document.myform.Action.value='DoBatchMove';"">&nbsp; "
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_GuestBook.asp?Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
        Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchKindID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchKindID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchKindID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchKindID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'函数名：DoBatchMove()
'作  用：批量移动主题到其他至类别的保存操作
'参  数：无
'=================================================
Sub DoBatchMove()
    If AdminPurview = 2 And AdminPurview_GuestBook > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    Dim GuestType, BatchGuestID, BatchKindID, tKindID
    
    GuestType = PE_CLng(Trim(Request("GuestType")))
    BatchGuestID = Trim(Request.Form("BatchGuestID"))
    BatchKindID = Trim(Request.Form("BatchKindID"))
    tKindID = Trim(Request("tKindID"))
    
    If GuestType = 1 Then
        If IsValidID(BatchGuestID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量移动的留言的ID</li>"
        End If
    Else
        If IsValidID(BatchKindID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量移动的留言的类别</li>"
        End If
    End If

    If tKindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定目标类别！</li>"
    Else
        tKindID = PE_CLng(tKindID)
    End If
    If FoundErr = True Then Exit Sub
            
    If GuestType = 1 Then
        Conn.Execute ("update PE_GuestBook set KindID=" & tKindID & " where GuestID in (select GuestID from PE_GuestBook where TopicID in (" & BatchGuestID & "))")
    Else
        Conn.Execute ("update PE_GuestBook set KindID=" & tKindID & " where GuestID in (select GuestID from PE_GuestBook where KindID in (" & BatchKindID & "))")
    End If
    
    ComeUrl = "Admin_GuestBook.asp?Action=BatchMove"
    Call WriteSuccessMsg("成功将选定的留言移动到目标类别中！", ComeUrl)
    Call ClearSiteCache(4)
End Sub

'=================================================
'函数名：FilterNotTopicID()
'作  用：过滤不是主题的ID
'参  数：BatchGuestID 留言ID
'=================================================
Function FilterNotTopicID(BatchGuestID)
    Dim arrGuestID, arrBatchGuestID
    Set arrGuestID = Conn.Execute("select GuestID from PE_GuestBook where GuestID in (" & arrBatchGuestID & ") and GuestID=TopicID")
    
    Do While Not arrGuestID.EOF
        If FilterNotTopicID = "" Or IsNull(FilterNotTopicID) Then
            FilterNotTopicID = arrGuestID("GuestID")
        Else
            FilterNotTopicID = FilterNotTopicID & "," & arrGuestID("GuestID")
        End If
        arrGuestID.MoveNext
    Loop

    Set arrGuestID = Nothing
End Function

'=================================================
'函数名：GetKind_Option()
'作  用：得到类别
'参  数：CurrentID 所属类别ID
'=================================================
Function GetKind_Option(CurrentID)
    Dim rsKind, sqlKind, strKind_Option
    CurrentID = PE_CLng(CurrentID)
    sqlKind = "Select * from PE_GuestKind order by OrderID,KindID"
    Set rsKind = Conn.Execute(sqlKind)
    If rsKind.BOF And rsKind.EOF Then
        strKind_Option = strKind_Option & "<option value=''>请先添加类别</option>"
    Else
        Do While Not rsKind.EOF
            strKind_Option = strKind_Option & "<option value='" & rsKind("KindID") & "'"
            If rsKind("KindID") = CurrentID Then
                strKind_Option = strKind_Option & " selected"
            End If
            strKind_Option = strKind_Option & ">"
            strKind_Option = strKind_Option & rsKind("KindName")
            strKind_Option = strKind_Option & "</option>"
            rsKind.MoveNext
        Loop
    End If
    rsKind.Close
    Set rsKind = Nothing
    If GuestBook_IsAssignSort = True Then
        strKind_Option = strKind_Option & "<option value='0'"
        If CurrentID = 0 Then strKind_Option = strKind_Option & " selected"
        strKind_Option = strKind_Option & ">不指定类别</option>"
    End If
    GetKind_Option = strKind_Option
End Function

'=================================================
'过程名：CreateCode()
'作  用：生成留言首页嵌入代码
'参  数：无
'=================================================
Sub CreateCode()
    If AdminPurview = 2 And AdminPurview_GuestBook > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "function ShowCommon(Num){" & vbCrLf
    Response.Write "    var commonNum,strHtml;" & vbCrLf
    Response.Write "    commonNum=Num-1;" & vbCrLf
    Response.Write "    //for(i=9;i>0;i--)" & vbCrLf
    Response.Write "       // {" & vbCrLf
    Response.Write "       // document.myform.commonPic+i.style.display='none';" & vbCrLf
    Response.Write "        //}" & vbCrLf
    Response.Write "    document.getElementById(""ShowCommonPic"").style.display='none';" & vbCrLf
    Response.Write "    if(Num>0)" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "        switch(commonNum){" & vbCrLf
    Response.Write "            case 0:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;符号:&nbsp;<font color=#b70000><b>·</b></font></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 1:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式1:&nbsp;<IMG src='../GuestBook/Images/common1.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 2:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式2:&nbsp;<IMG src='../GuestBook/Images/common2.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 3:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式3:&nbsp;<IMG src='../GuestBook/Images/common3.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 4:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式4:&nbsp;<IMG src='../GuestBook/Images/common4.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 5:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式5:&nbsp;<IMG src='../GuestBook/Images/common5.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 6:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式6:&nbsp;<IMG src='../GuestBook/Images/common6.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 7:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式7:&nbsp;<IMG src='../GuestBook/Images/common7.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            case 8:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式8:&nbsp;<IMG src='../GuestBook/Images/common8.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "            default:" & vbCrLf
    Response.Write "                strHtml=""<div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式9:&nbsp;<IMG src='../GuestBook/Images/common9.gif' border='0'></div>"";" & vbCrLf
    Response.Write "                break;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    else{" & vbCrLf
    Response.Write "        strHtml="""";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.getElementById(""ShowCommonPic"").innerHTML=strHtml;" & vbCrLf
    Response.Write "    if(Num>0)" & vbCrLf
    Response.Write "        {document.getElementById(""ShowCommonPic"").style.display='';" & vbCrLf
    Response.Write "    }}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    
    Response.Write "<form method='POST' action='Admin_GuestBook.asp' id='myform' name='myform'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
    Response.Write "    <tr class='topbg'>" & vbCrLf
    Response.Write "      <td colspan='3' align='center'><strong>首页嵌入代码生成</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>调用类别：</strong><br>      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <span>" & vbCrLf
    Response.Write "        <SELECT name='KindId' ID='KindId'>" & vbCrLf
    Response.Write "          <option value='0' selected>调用所有类别</option>" & vbCrLf
    Response.Write "          <option value='10000'>调用精华留言</option>" & vbCrLf
    Dim rsKind, sqlKind
    sqlKind = "Select * from PE_GuestKind order by OrderID,KindID"
    Set rsKind = Conn.Execute(sqlKind)
    If rsKind.BOF And rsKind.EOF Then
        Response.Write "" & vbCrLf
    Else
        Do While Not rsKind.EOF
            Response.Write " <option value='" & rsKind("KindID") & "'>" & rsKind("KindName") & "</option>"
            rsKind.MoveNext
        Loop
    End If
    rsKind.Close
    Set rsKind = Nothing
    Response.Write "      </SELECT>" & vbCrLf
    Response.Write "      </span> </td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>只显示主题留言</span>：</strong><br>" & vbCrLf
    Response.Write "选择是，则不显示回复的主题</td>" & vbCrLf
    Response.Write "      <td><label>" & vbCrLf
    Response.Write "        <input type='radio' name='OnlyTitle' value='1' checked>" & vbCrLf
    Response.Write "是</label>" & vbCrLf
    Response.Write "        <label>" & vbCrLf
    Response.Write "        <input type='radio' name='OnlyTitle' value='0'>" & vbCrLf
    Response.Write "否</label></td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>显示数量</span>：</strong><br>" & vbCrLf
    Response.Write "        <span>列表显示多少条留言主题</span></td>" & vbCrLf
    Response.Write "      <td width='25%'>        <span>" & vbCrLf
    Response.Write "        <INPUT TYPE='text' name='Num' size='4' Maxlength='10' value='8'>条留言" & vbCrLf
    Response.Write "      </span>      </td>" & vbCrLf
    Response.Write "      <td width='25%'>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>留言长度</span>：</strong><br>" & vbCrLf
    Response.Write "        <span>留言标题的长度,显示多少个字，一个汉字=两个英文字符 </span></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <span>" & vbCrLf
    Response.Write "        <INPUT TYPE='text' name='Titlelen' size='4' Maxlength='10' value='14'>" & vbCrLf
    Response.Write "      </span>个</td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>列表排序方式</span>：</strong><br>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <span>" & vbCrLf
    Response.Write "        <SELECT name='order' ID='order'>" & vbCrLf
    Response.Write "          <option value='1'>按时间排序</option>" & vbCrLf
    Response.Write "          <option value='0' selected>按操作排序</option>" & vbCrLf
    Response.Write "        </SELECT>" & vbCrLf
    Response.Write "      </span> </td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>主题图片标志</span>：</strong><br>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <span>" & vbCrLf
    Response.Write "          <select name='ShowPic' id='ShowPic' onchange='ShowCommon(this.value)'>" & vbCrLf
    Response.Write "            <option value='0'>不显示</option> " & vbCrLf
    Response.Write "            <option value='1'>符号</option>    " & vbCrLf
    Response.Write "            <option value='2'>小图片（样式1）</option>" & vbCrLf
    Response.Write "            <option value='3'>小图片（样式2）</option>" & vbCrLf
    Response.Write "            <option value='4'>小图片（样式3）</option>" & vbCrLf
    Response.Write "            <option value='5'>小图片（样式4）</option>" & vbCrLf
    Response.Write "            <option value='6'>小图片（样式5）</option" & vbCrLf
    Response.Write "            ><option value='7'>小图片（样式6）</option>" & vbCrLf
    Response.Write "            <option value='8'>小图片（样式7）</option>" & vbCrLf
    Response.Write "            <option value='9'>小图片（样式8）</option>" & vbCrLf
    Response.Write "            <option value='10' selected>小图片（样式9）</option>" & vbCrLf
    Response.Write "          </select>" & vbCrLf
    Response.Write "        </span>" & vbCrLf
    Response.Write "        " & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      <td id='ShowCommonPic'><div id='commonPic1' name='commonPic1'>&nbsp;&nbsp;样式9:&nbsp;<IMG src='../GuestBook/Images/common9.gif' border='0'></div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>显示类别名称</span>：</strong><br>" & vbCrLf
    Response.Write "选择是，则显示类别名称。如：“[疑难解答]”</td>" & vbCrLf
    Response.Write "      <td><label>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowKindName' value='1'>" & vbCrLf
    Response.Write "是</label>" & vbCrLf
    Response.Write "        <label>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowKindName' value='0' checked>" & vbCrLf
    Response.Write "否</label></td>" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>是否显示留言内容字数</span>：</strong><br>" & vbCrLf
    Response.Write "选择是，则显示留言内容字数。如：“<I><font color=gray>(87字)</font></I>”</td>" & vbCrLf
    Response.Write "      <td><label>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowContentLen' value='1'>" & vbCrLf
    Response.Write "是</label>" & vbCrLf
    Response.Write "        <label>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowContentLen' value='0' checked>" & vbCrLf
    Response.Write "否</label></td>" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr> " & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>显示留言时间：</strong></td>" & vbCrLf
    Response.Write "      <td><span>" & vbCrLf
    Response.Write "        <SELECT name='ShowTime' ID='ShowTime'>" & vbCrLf
    Response.Write "          <option value='0' selected>不显示</option>" & vbCrLf
    Response.Write "          <option value='2'>短日期</option>" & vbCrLf
    Response.Write "          <option value='3'>时间</option>" & vbCrLf
    Response.Write "          <option value='1'>短日期+时间</option>" & vbCrLf
    Response.Write "          <option value='4'>格式化后的时间</option>" & vbCrLf
    Response.Write "        </SELECT>" & vbCrLf
    Response.Write "      </span></td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><span>是否显示用户名</span>：</strong><br>" & vbCrLf
    Response.Write "选择是，则显示用户名</td>" & vbCrLf
    Response.Write "      <td><label>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowUserName' value='1'>" & vbCrLf
    Response.Write "是</label>" & vbCrLf
    Response.Write "        <label>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowUserName' value='0' checked>" & vbCrLf
    Response.Write "否</label></td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td colspan='3'>显示样式举例：<IMG src='../GuestBook/Images/common2.gif' border='0'>[知心姐姐] 羡慕啊，你的站点真酷！<I><font color=gray>(251字) － 冰雪空气，</font><font color=green>2005-12-12 01:28:27</font></I> </td>" & vbCrLf
    Response.Write "    </tr> " & vbCrLf
    Response.Write "   " & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='3' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoCreateCode'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value='生成嵌入代码'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "" & vbCrLf
End Sub
'=================================================
'过程名：DoCreateCode()
'作  用：生成留言首页嵌入代码结果
'参  数：无
'KindId         KindId=0表示调用所有类别的最新留言，KindId为不同的值对应不同类别，KindId=10000只显示精华留言
'OnlyTitle      为0显示主题留言和回复,为1只显示主题留言，不显示回复
'Num            显示数量，列表显示多少条留言主题
'Titlelen       留言长度，留言标题的长度，显示多少个字
'Order          如果为0 按操作排序 1 则按留言时间排序
'
'ShowPic        主题图片标志 0 不显示 1 符号，2 图片（样式一）
'ShowKindName   是否显示留言类别    为0不显示,为1显示
'ShowContentLen 是否显示留言内容字数 0 不显示 1 显示
'ShowTime       显示时间 0 不显示 1 短日期+长时间 2 短日期 3 时间 4 格式化后的时间
'ShowUserName   是否显示用户名 0 不显示 1 显示
'=================================================
Sub DoCreateCode()
    If AdminPurview = 2 And AdminPurview_GuestBook > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    Dim Titlelen, Num, Order, KindID, OnlyTitle, ShowKindName, ShowContentLen, ShowUserName, ShowTime, ShowPic
    
    KindID = PE_CLng(Trim(Request("KindID")))
    OnlyTitle = PE_CLng(Trim(Request("OnlyTitle")))
    Num = PE_CLng(Trim(Request("Num")))
    Titlelen = PE_CLng(Trim(Request("Titlelen")))
    Order = PE_CLng(Trim(Request("Order")))
    
    ShowPic = PE_CLng(Trim(Request("ShowPic")))
    ShowKindName = PE_CLng(Trim(Request("ShowKindName")))
    ShowContentLen = PE_CLng(Trim(Request("ShowContentLen")))
    ShowUserName = PE_CLng(Trim(Request("ShowUserName")))
    ShowTime = PE_CLng(Trim(Request("ShowTime")))
    
    
    Response.Write " <br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
    Response.Write "   <tr class='title'>  " & vbCrLf
    Response.Write "      <td height='22' colspan='2'> " & vbCrLf
    Response.Write "        <div align='center'><strong>嵌入代码生成结果</strong></div>" & vbCrLf
    Response.Write "      </td>    " & vbCrLf
    Response.Write "   </tr>       " & vbCrLf
    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='300' class='tdbg'>请将此代码拷贝到您需要输出最新留言列表的页面：" & vbCrLf
    Response.Write "      </td>      " & vbCrLf
    Response.Write "     <td class='tdbg'><textarea name='strCode' cols='58' rows='6' id='strCode'>" & vbCrLf
    Response.Write "  <script src='{$InstallDir}guestbook/newguest.asp?KindId=" & KindID & "&OnlyTitle=" & OnlyTitle & "&num=" & Num & "&Titlelen=" & Titlelen & "&Order=" & Order & "&ShowPic=" & ShowPic & "&ShowKindName=" & ShowKindName & "&ShowContentLen=" & ShowContentLen & "&ShowUserName=" & ShowUserName & "&ShowTime=" & ShowTime & "'></script>"
    Response.Write "     </textarea>" & vbCrLf
    Response.Write "     <br></td>    </tr>  </table>" & vbCrLf
End Sub


Function IsRadioChecked(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function
'=================================================
'函数名：ShowTip()
'作  用：显示浮动信息
'参  数：无
'=================================================
Sub ShowTip()
    Response.Write "<div id=toolTipLayer style='position: absolute; visibility: hidden'></div>" & vbCrLf
    Response.Write "<SCRIPT language=JavaScript>" & vbCrLf
    Response.Write "var ns4 = document.layers;" & vbCrLf
    Response.Write "var ns6 = document.getElementById && !document.all;" & vbCrLf
    Response.Write "var ie4 = document.all;" & vbCrLf
    Response.Write "offsetX = 0;" & vbCrLf
    Response.Write "offsetY = 20;" & vbCrLf
    Response.Write "var toolTipSTYLE='';" & vbCrLf
    Response.Write "function initToolTips()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(ns4||ns6||ie4)" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    if(ns4) toolTipSTYLE = document.toolTipLayer;" & vbCrLf
    Response.Write "    else if(ns6) toolTipSTYLE = document.getElementById('toolTipLayer').style;" & vbCrLf
    Response.Write "    else if(ie4) toolTipSTYLE = document.all.toolTipLayer.style;" & vbCrLf
    Response.Write "    if(ns4) document.captureEvents(Event.MOUSEMOVE);" & vbCrLf
    Response.Write "    else" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      toolTipSTYLE.visibility = 'visible';" & vbCrLf
    Response.Write "      toolTipSTYLE.display = 'none';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.onmousemove = moveToMouseLoc;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function toolTip(msg, fg, bg)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(toolTip.arguments.length < 1)" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    if(ns4) toolTipSTYLE.visibility = 'hidden';" & vbCrLf
    Response.Write "    else toolTipSTYLE.display = 'none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    if(!fg) fg = '#333333';" & vbCrLf
    Response.Write "    if(!bg) bg = '#FFFFFF';" & vbCrLf
    Response.Write "    var content = '<table border=""0"" cellspacing=""0"" cellpadding=""1"" bgcolor=""' + fg + '""><td>' + '<table border=""0"" cellspacing=""0"" cellpadding=""1"" bgcolor=""' + bg + '""><td align=""left"" nowrap style=""line-height: 120%""><font color=""' + fg + '"">' + msg + '&nbsp\;</font></td></table></td></table>';" & vbCrLf
    Response.Write "    if(ns4)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      toolTipSTYLE.document.write(content);" & vbCrLf
    Response.Write "      toolTipSTYLE.document.close();" & vbCrLf
    Response.Write "      toolTipSTYLE.visibility = 'visible';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ns6)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      document.getElementById('toolTipLayer').innerHTML = content;" & vbCrLf
    Response.Write "      toolTipSTYLE.display='block'" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ie4)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      document.all('toolTipLayer').innerHTML=content;" & vbCrLf
    Response.Write "      toolTipSTYLE.display='block'" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function moveToMouseLoc(e)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(ns4||ns6)" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    x = e.pageX;" & vbCrLf
    Response.Write "    y = e.pageY;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    x = event.x + document.body.scrollLeft;" & vbCrLf
    Response.Write "    y = event.y + document.body.scrollTop;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  toolTipSTYLE.left = x + offsetX;" & vbCrLf
    Response.Write "  toolTipSTYLE.top = y + offsetY;" & vbCrLf
    Response.Write "  return true;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "initToolTips();" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

'=================================================
'函数名：CheckKindPurview()
'作  用：检查用户类别权限
'参  数：Num----权限内容
'        KindID  ----类别ID
'=================================================
Function CheckKindPurview(ByVal Num, ByVal KindID)
    Dim arrGuestBook, arrNum, arrKindID, KindPurview
    KindPurview = False
    arrNum = PE_CLng(Trim(Num))
    arrKindID = PE_CLng(Trim(KindID))
    CheckKindPurview = False
    arrGuestBook = Split(arrClass_GuestBook, "|||")
    If arrNum > UBound(arrGuestBook) Then
        KindPurview = False
    Else
        If InStr(arrGuestBook(arrNum), arrKindID) > 0 Then
            KindPurview = True
        End If
    End If
    CheckKindPurview = KindPurview
End Function
%>
