<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub Execute()
    strFileName = "User_Friend.asp"

    Select Case Action
    Case "AddFriend"
        Call AddFriend
    Case "SaveNewFriend"
        Call SaveNewFriend
    Case "DelFriend"
        Call DelFriend
    Case "Move"
        Call Move
    Case "ManageGroup"
        Call ManageGroup
    Case "CreateNewGroup"
        Call CreateNewGroup
    Case "SaveNewGroup"
        Call SaveNewGroup
    Case "ModifyGroup"
        Call ModifyGroup
    Case "SaveModifyGroup"
        Call SaveModifyGroup
    Case "DelGroup"
        Call DelGroup
    Case Else
        Call main
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub


Sub Move()
    Dim FriendID, GroupID
    FriendID = Request.Form("FriendID")
    GroupID = Request.Form("GroupID")
    If IsValidID(FriendID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>指定的用户错误或未指定用户！</li>"
        Exit Sub
    End If
    If GroupID = "" Or IsNull(GroupID) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>未指定用户组！</li>"
        Exit Sub
    Else
        GroupID = PE_CLng(GroupID)
    End If
    Conn.Execute ("Update PE_Friend set GroupID=" & GroupID & " where UserName='" & UserName & "' and ID in (" & FriendID & ")")
    Call WriteSuccessMsg("<li>批量移动成功。</li>", ComeUrl)
End Sub

Sub DelFriend()
    Dim FriendID
    FriendID = Request.Form("FriendID")
    If IsValidID(FriendID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>指定的用户ID错误！</li>"
        Exit Sub
    End If

    Conn.Execute ("delete from PE_Friend where UserName='" & UserName & "' and ID in (" & FriendID & ")")

    Call WriteSuccessMsg("<li>删除用户成功。</li>", ComeUrl)

End Sub

Sub main()
    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If
    Dim GroupID, strJS
    Dim sqlFriend, rsFriend, sqlGroup, rsGroup, i, GetFriendGroup
    GroupID = Trim(Request("GroupID"))
    If GroupID <> "" Or IsNull(GroupID) Then
        GroupID = PE_CLng(GroupID)
    End If
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function unselectall(){" & vbCrLf
    Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.Name != 'chkAll'&&e.disabled==false)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "function ConfirmDel(){" & vbCrLf
    Response.Write " if(document.myform.Action.value=='DelFriend'){" & vbCrLf
    Response.Write "     if(confirm('确定要删除选中的用户吗？'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " if(document.myform.Action.value=='Move'){" & vbCrLf
    Response.Write "     if(document.myform.GroupID.value==''){" & vbCrLf
    Response.Write "         alert('请选择移动到的组别！');" & vbCrLf
    Response.Write "         document.myform.GroupID.focus();" & vbCrLf
    Response.Write "         return false;}" & vbCrLf
    Response.Write "     if(confirm('确定要批量移动吗？'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "          return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
    
    sqlGroup = "select UserFriendGroup from PE_User where UserName='" & UserName & "'"
    Set rsGroup = Conn.Execute(sqlGroup)
    If rsGroup.BOF And rsGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsGroup(0) = "" Or IsNull(rsGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        Else
            GetFriendGroup = Split(rsGroup(0), "$")
        End If

        If UBound(GetFriendGroup) < 1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了默认组！</li>"
            Exit Sub
        End If
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>"
    For i = UBound(GetFriendGroup) To 0 Step -1
        'Response.Write "<option value='" & i & "'>" & GetFriendGroup(i) & "</option>"
        'Response.Write "GroupID="&GroupID
        'Response.Write "aai="&i
        If i = GroupID Then
            Response.Write "<a href='User_Friend.asp?GroupID=" & i & "'><font color='red'>" & GetFriendGroup(i) & "</font></a>"
        Else
            Response.Write "<a href='User_Friend.asp?GroupID=" & i & "'>" & GetFriendGroup(i) & "</a>"
        End If
        Response.Write " | "
    Next
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_Friend.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='30' width='30' align='center'><strong>选中</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>姓名</strong></td>"
    Response.Write "            <td width='100' align='center'><strong>组别</strong></td>"
    Response.Write "            <td width='150' align='center' ><strong>邮件</strong></td>"
    Response.Write "            <td width='160' align='center' ><strong>主页</strong></td>"
    Response.Write "            <td width='70' align='center' ><strong>QQ</strong></td>"
    Response.Write "            <td align='center' ><strong>操作</strong></td>"
    Response.Write "          </tr>"

    '"select D.ID, D.EquipmentID, D.UserName, D.lessonMonth, D.lessonDay, D.lessonNumber, D.lessonYear, D.UserClass, D.UserType, D.RegisterTime,D.Used, F.ClassroomID, F.EquipmentName, C.ClassroomName from PE_UsedDetail D left join ( PE_Equipment F left join  PE_Classroom C on F.ClassroomID = C.ID ) on D.EquipmentID = F.ID where 1=1"

    sqlFriend = "select F.ID,F.FriendName,F.AddTime,F.GroupID,U.Email,C.QQ,C.Homepage from PE_Friend F left join ( PE_User U left join PE_Contacter C on U.ContacterID = C.ContacterID ) on F.FriendName=U.UserName where F.UserName='" & UserName & "'"
    If GroupID <> "" Or IsNull(GroupID) Then
        sqlFriend = sqlFriend & " and F.GroupID=" & PE_CLng(GroupID)
    End If
    sqlFriend = sqlFriend & " order by F.AddTime desc"
    Set rsFriend = Server.CreateObject("adodb.recordset")
    rsFriend.open sqlFriend, Conn, 1, 1
    If rsFriend.BOF And rsFriend.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>尚未添加任何成员！<br><br></td></tr>"
    Else
        totalPut = rsFriend.RecordCount
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
                rsFriend.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim FriendNum
        FriendNum = 0
        Do While Not rsFriend.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'><input name='FriendID' type='checkbox' onclick='unselectall()' id='MessageID' value='" & rsFriend("ID") & "'></td>"
            Response.Write "      <td width='80' align='center'>" & rsFriend("FriendName") & "</td>"
            Response.Write "      <td width='100' align='center'>"
            'response.write UBound(GetFriendGroup) & "<br>" &  PE_CLng(rsFriend("GroupID"))
            'response.end
            If UBound(GetFriendGroup) < PE_CLng(rsFriend("GroupID")) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>数据库信息错误！</li>"
                Exit Sub
            Else
                Response.Write GetFriendGroup(PE_CLng(rsFriend("GroupID")))
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='150' align='center'>" & rsFriend("Email") & "</td>"
            Response.Write "    <td width='160' align='center'>"
            If rsFriend("Homepage") = "" Or IsNull(rsFriend("Homepage")) Then
                Response.Write "未填"
            Else
                Response.Write rsFriend("Homepage")
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='70' align='center'>"
            If rsFriend("QQ") = "" Or IsNull(rsFriend("QQ")) Then
                Response.Write "未填"
            Else
                Response.Write rsFriend("QQ")
            End If
            Response.Write "</td>"
            Response.Write "    <td align='center'>"
            Response.Write "<a href='User_Message.asp?Action=New&inceptUser=" & rsFriend("FriendName") & "'>发短消息</a>"
            Response.Write "</td>"
            Response.Write "</tr>"

            FriendNum = FriendNum + 1
            If FriendNum >= MaxPerPage Then Exit Do
            rsFriend.MoveNext
        Loop
    End If
    rsFriend.Close
    Set rsFriend = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中本页显示的所有用户</td><td>"
    Response.Write "<input name='submit1' type='submit' value='删除选定的用户' onClick=""document.myform.Action.value='DelFriend'"" >"
    Response.Write "      &nbsp;&nbsp;&nbsp;<select name='GroupID'>" & vbCrLf
    Response.Write "<option value=''>将选定的用户移动到...</option>"
    For i = UBound(GetFriendGroup) To 0 Step -1
        Response.Write "<option value='" & i & "'>" & GetFriendGroup(i) & "</option>"
    Next
    Set rsGroup = Nothing
    Response.Write "      </select>" & vbCrLf
    Response.Write "&nbsp;<input name='submit1' type='submit' value='移动' onClick=""document.myform.Action.value='Move'"" >"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"

    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个用户", True)
End Sub

Sub SaveNewFriend()
    Dim FriendName, GroupID, rsFriendName, rsFriend, sqlFriend, rsFriendExist, i
    FriendName = ReplaceBadChar(Request.Form("FriendName"))
    GroupID = Request.Form("GroupID")
    If FriendName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>成员用户名不能为空！</li>"
        Exit Sub
    End If

    If GroupID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>成员组ID不能为空！</li>"
        Exit Sub
    Else
        GroupID = PE_CLng(GroupID)
    End If
        
    FriendName = Split(FriendName, ",")
    Dim strTemp
    For i = 0 To UBound(FriendName)
        If strTemp = "" Then
            strTemp = FriendName(i)
        Else
            If FoundInArr(strTemp, FriendName(i), ",") = False And FriendName(i) <> UserName Then
                strTemp = strTemp & "," & FriendName(i)
            End If
        End If
    Next
    FriendName = Split(strTemp, ",")
    Set rsFriend = Server.CreateObject("adodb.recordset")
    sqlFriend = "select * from PE_Friend"
    rsFriend.open sqlFriend, Conn, 1, 3
    For i = 0 To UBound(FriendName)
        If i >= 5 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>最多只能发送给6个用户，您的名单5位以后的请重新添加！</li>"
            Exit For
        End If
        Set rsFriendName = Conn.Execute("select UserName From PE_User Where UserName='" & FriendName(i) & "'")
        If rsFriendName.BOF And rsFriendName.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>添加的用户名不存在！</li>"
            Exit Sub
        End If
        Set rsFriendExist = Conn.Execute("select UserName From PE_Friend Where FriendName='" & FriendName(i) & "' and UserName='" & UserName & "'")
        If Not (rsFriendExist.BOF And rsFriendExist.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>禁止重复添加用户！</li>"
            Exit Sub
        End If

        rsFriend.addnew
        rsFriend("UserName") = UserName
        rsFriend("FriendName") = FriendName(i)
        rsFriend("AddTime") = Now()
        rsFriend("GroupID") = GroupID
        rsFriend.Update
    Next
    Set rsFriend = Nothing
    Call WriteSuccessMsg("添加成功！", "User_Friend.asp")
End Sub

Sub AddFriend()
    Dim sqlGroup, rsGroup, GetFriendGroup, i, strHTML
    strHTML = "<script language=javascript>" & vbCrLf
    strHTML = strHTML & "function CheckSubmit(){" & vbCrLf
    strHTML = strHTML & "  if(document.form1.FriendName.value==''){" & vbCrLf
    strHTML = strHTML & "      alert('成员用户名不能为空！');" & vbCrLf
    strHTML = strHTML & "   document.form1.FriendName.focus();" & vbCrLf
    strHTML = strHTML & "      return false;" & vbCrLf
    strHTML = strHTML & "    }" & vbCrLf
    strHTML = strHTML & "}" & vbCrLf
    strHTML = strHTML & "</script>" & vbCrLf
    strHTML = strHTML & "<form method='post' action='User_Friend.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    strHTML = strHTML & " <br> <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    strHTML = strHTML & "    <tr class='title'>" & vbCrLf
    strHTML = strHTML & "      <td height='22' colspan='2'><div align='center'>添 加 成 员</div></td>" & vbCrLf
    strHTML = strHTML & "    </tr>" & vbCrLf
    strHTML = strHTML & "    <tr class='tdbg'>" & vbCrLf
    strHTML = strHTML & "      <td width='25%' class='tdbg5' align='right'>成员用户名：</td>" & vbCrLf
    strHTML = strHTML & "      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='FriendName' type='text' id='FriendName' size='25' maxlength='30'>&nbsp;&nbsp;<font color='#FF0000'>*</font></td>" & vbCrLf
    strHTML = strHTML & "    </tr>" & vbCrLf
    strHTML = strHTML & "    <tr class='tdbg'>" & vbCrLf
    strHTML = strHTML & "      <td width='25%' class='tdbg5' align='right'>成 员 组：</td>" & vbCrLf
    strHTML = strHTML & "      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select name='GroupID'>" & vbCrLf

    sqlGroup = "select UserFriendGroup from PE_User where UserName='" & UserName & "'"
    Set rsGroup = Conn.Execute(sqlGroup)
    If rsGroup.BOF And rsGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsGroup(0) = "" Or IsNull(rsGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        Else
            GetFriendGroup = Split(rsGroup(0), "$")
        End If
        If UBound(GetFriendGroup) < 1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了默认组！</li>"
            Exit Sub
        End If
    End If
    For i = UBound(GetFriendGroup) To 0 Step -1
        If i = UBound(GetFriendGroup) Then
            strHTML = strHTML & "<option value='" & i & "' selected>" & GetFriendGroup(i) & "</option>"
        Else
            strHTML = strHTML & "<option value='" & i & "'>" & GetFriendGroup(i) & "</option>"
        End If
    Next
    Set rsGroup = Nothing
    strHTML = strHTML & "      </select>&nbsp;&nbsp;<font color='#FF0000'>*</font></td>" & vbCrLf
    strHTML = strHTML & "    </tr>" & vbCrLf
    strHTML = strHTML & "            <tr class='tdbg'>" & vbCrLf
    strHTML = strHTML & "                <td align='center'  colspan='2'>" & vbCrLf
    strHTML = strHTML & "                    <input type='hidden' name='Action' value='SaveNewFriend'>" & vbCrLf
    strHTML = strHTML & "                    <input type='submit' value='添加成员'>" & vbCrLf
    strHTML = strHTML & "                    <input type='button' name='cancel' value=' 取 消 ' onClick=""JavaScript:window.location.href='User_Friend.asp'"">" & vbCrLf
    strHTML = strHTML & "                </td>" & vbCrLf
    strHTML = strHTML & "            </tr>" & vbCrLf
    strHTML = strHTML & "  </table></form>" & vbCrLf
    strHTML = strHTML & "    <br>" & vbCrLf
    strHTML = strHTML & "    <b>&nbsp;&nbsp;注：</b><br>" & vbCrLf
    strHTML = strHTML & "    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1、可以用英文状态下的逗号将用户名隔开实现添加多个用户，最多<b>5</b>个用户。<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2、已经添加过的成员，不允许重复添加。" & vbCrLf
    Response.Write strHTML
End Sub

Sub CreateNewGroup()
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "  if(document.form1.GroupName.value==''){" & vbCrLf
    Response.Write "      alert('新创建的组名称不能为空！');" & vbCrLf
    Response.Write "   document.form1.GroupName.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form method='post' action='User_Friend.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "  <br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='2'><div align='center'>创 建 新 组</div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='25%' class='tdbg5' align='right'>新组名称：</td>" & vbCrLf
    Response.Write "      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='GroupName' type='text' id='GroupName' size='20' maxlength='20'>&nbsp;&nbsp;<font color='#FF0000'>*</font>&nbsp;不超过6个汉字</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "            <tr class='tdbg'>" & vbCrLf
    Response.Write "                <td align='center'  colspan='2'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='Action' value='SaveNewGroup'>" & vbCrLf
    Response.Write "                    <input type='submit' value='添加成员组'>" & vbCrLf
    Response.Write "                    <input type='button' name='cancel' value=' 取 消 ' onClick=""JavaScript:window.location.href='User_Friend.asp'"">" & vbCrLf
    Response.Write "                </td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "  </table></form>" & vbCrLf
    Response.Write "    <br>" & vbCrLf
    Response.Write "    <b>&nbsp;&nbsp;注：</b><br>" & vbCrLf
    Response.Write "    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;网站限制创建8个分组。" & vbCrLf
End Sub

Sub SaveNewGroup()
    Dim rsUserFriendGroup, GetFriendGroup, GroupName
    GroupName = ReplaceBadChar(Request("GroupName"))
    If GroupName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>新创建的组名称不能为空！</li>"
        Exit Sub
    End If
    If GetStrLen(GroupName) > 12 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>新创建的组名称不能超过6个汉字！</li>"
        Exit Sub
    End If
    Set rsUserFriendGroup = Conn.Execute("select UserFriendGroup from PE_User where UserName='" & UserName & "'")
    If rsUserFriendGroup.BOF And rsUserFriendGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsUserFriendGroup(0) = "" Or IsNull(rsUserFriendGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        End If
        If UBound(Split(rsUserFriendGroup(0), "$")) < 1 Or UBound(Split(rsUserFriendGroup(0), "$")) > 7 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组或添加组超过8个了！</li>"
            Exit Sub
        Else
            GetFriendGroup = rsUserFriendGroup(0) & "$" & GroupName
        End If
    End If
    Set rsUserFriendGroup = Nothing
    Conn.Execute ("update PE_User set UserFriendGroup= '" & GetFriendGroup & "' where UserName='" & UserName & "'")
    Response.Redirect "User_Friend.asp?Action=ManageGroup"
End Sub

Sub ManageGroup()
    Dim rsUserFriendGroup, GetFriendGroup, j, i
    Set rsUserFriendGroup = Conn.Execute("select UserFriendGroup from PE_User where UserName='" & UserName & "'")
    If rsUserFriendGroup.BOF And rsUserFriendGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsUserFriendGroup(0) = "" Or IsNull(rsUserFriendGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        Else
            GetFriendGroup = Split(rsUserFriendGroup(0), "$")
        End If
        If UBound(GetFriendGroup) < 1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        End If
    End If
    Set rsUserFriendGroup = Nothing
    Response.Write "    <br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' height='22' class='title'>" & vbCrLf
    Response.Write "    <td width='60'>ID</td>" & vbCrLf
    Response.Write "    <td width='200'>成员组名</td>" & vbCrLf
    Response.Write "    <td width='80'>成员数量</td>" & vbCrLf
    Response.Write "    <td>操 作</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    j = 0
'Response.Write "aaa="&Conn.Execute("select count(*) from PE_Friend where UserName='" & UserName &"' and GroupID=1")(0)
'response.end
    For i = UBound(GetFriendGroup) To 0 Step -1
        j = j + 1
        Response.Write "     <tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'"">" & vbCrLf
        Response.Write "    <td width='60'>" & j & "</td>" & vbCrLf
        Response.Write "    <td width='200'>" & GetFriendGroup(i) & "</td>" & vbCrLf
        Response.Write "    <td width='80'>" & vbCrLf
        Response.Write Conn.Execute("select count(*) from PE_Friend where UserName='" & UserName & "' and GroupID=" & i & "")(0)
        Response.Write "    </td>" & vbCrLf
        If i <> 0 Then
            Response.Write "    <td><a href='User_Friend.asp?Action=ModifyGroup&GroupID=" & i & "'>修改</a>" & vbCrLf
        Else
            Response.Write "    <td><font color='#CCCCCC'>修改</font>" & vbCrLf
        End If
        If i = 0 Or i = 1 Then
            Response.Write " | <font color='#CCCCCC'>删除</font> | " & vbCrLf
        Else
            Response.Write " | <a href='User_Friend.asp?Action=DelGroup&GroupID=" & i & "' onclick=""return confirm('删除该分组后，该分组中的好友也将删除，确定要删除此组吗？');"">删除</a> | " & vbCrLf
        End If
        Response.Write "<a href='User_Friend.asp?GroupID=" & i & "'>列出名单</a></td>    </tr>" & vbCrLf
    Next
    Response.Write "    </table>" & vbCrLf
    Response.Write "    <br>" & vbCrLf
    Response.Write "    <b>&nbsp;&nbsp;注：</b><br>" & vbCrLf
    Response.Write "    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;默认组黑名单，拒收所有来自黑名单的短信。" & vbCrLf
End Sub


Sub ModifyGroup()
    Dim GroupID, rsUserFriendGroup, GetFriendGroup
    GroupID = Request("GroupID")
    If GroupID = "" Or IsNull(GroupID) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>成员组ID不能为空！</li>"
        Exit Sub
    Else
        GroupID = PE_CLng(GroupID)
    End If
    Set rsUserFriendGroup = Conn.Execute("select UserFriendGroup from PE_User where UserName='" & UserName & "'")
    If rsUserFriendGroup.BOF And rsUserFriendGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsUserFriendGroup(0) = "" Or IsNull(rsUserFriendGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        Else
            GetFriendGroup = Split(rsUserFriendGroup(0), "$")
        End If
        If UBound(GetFriendGroup) < 1 Or UBound(GetFriendGroup) < GroupID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了默认组！</li>"
            Exit Sub
        End If
    End If
    Set rsUserFriendGroup = Nothing

    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "  if(document.form1.GroupName.value==''){" & vbCrLf
    Response.Write "      alert('组名称不能为空！');" & vbCrLf
    Response.Write "   document.form1.GroupName.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form method='post' action='User_Friend.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "  <br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='2'><div align='center'>修 改 成 员 组</div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='25%' class='tdbg5' align='right'>名 称：</td>" & vbCrLf
    Response.Write "      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='GroupName' type='text' id='GroupName' value='" & GetFriendGroup(GroupID) & "' size='20' maxlength='20'>&nbsp;&nbsp;<font color='#FF0000'>*</font>&nbsp;不超过6个汉字</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "            <tr class='tdbg'>" & vbCrLf
    Response.Write "                <td align='center'  colspan='2'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='Action' value='SaveModifyGroup'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='GroupID' value='" & GroupID & "'>" & vbCrLf
    Response.Write "                    <input type='submit' value=' 修 改 '>" & vbCrLf
    Response.Write "                    <input type='button' name='cancel' value=' 取 消 ' onClick=""JavaScript:window.location.href='User_Friend.asp?Action=ManageGroup'"">" & vbCrLf
    Response.Write "                </td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "  </table></form>" & vbCrLf
End Sub

Sub DelGroup()
    Dim GroupID, rsUserFriendGroup, GroupName, GetFriendGroup
    GroupID = Request("GroupID")
    If GroupID = "" Or IsNull(GroupID) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>成员组ID不能为空！</li>"
        Exit Sub
    Else
        GroupID = PE_CLng(GroupID)
    End If
    Set rsUserFriendGroup = Conn.Execute("select UserFriendGroup from PE_User where UserName='" & UserName & "'")
    If rsUserFriendGroup.BOF And rsUserFriendGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsUserFriendGroup(0) = "" Or IsNull(rsUserFriendGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        Else
            GetFriendGroup = Split(rsUserFriendGroup(0), "$")
        End If
        If UBound(GetFriendGroup) < 1 Or UBound(GetFriendGroup) < GroupID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了默认组！</li>"
            Exit Sub
        End If
    End If
    If InStr(rsUserFriendGroup(0), "$" & GetFriendGroup(GroupID)) + Len("$" & GetFriendGroup(GroupID)) - 1 = Len(rsUserFriendGroup(0)) Then
        GroupName = Left(rsUserFriendGroup(0), InStr(rsUserFriendGroup(0), "$" & GetFriendGroup(GroupID)) - 1)
    Else
        Dim RightLength
        RightLength = Len(rsUserFriendGroup(0)) - (InStr(rsUserFriendGroup(0), "$" & GetFriendGroup(GroupID)) + Len("$" & GetFriendGroup(GroupID)) - 1)
        GroupName = Left(rsUserFriendGroup(0), InStr(rsUserFriendGroup(0), "$" & GetFriendGroup(GroupID)) - 1) & Right(rsUserFriendGroup(0), RightLength)
    End If
    Set rsUserFriendGroup = Nothing
    Conn.Execute ("update PE_User set UserFriendGroup= '" & GroupName & "' where UserName='" & UserName & "'")
    Conn.Execute ("Delete from PE_Friend Where GroupID=" & GroupID & " and UserName='" & UserName & "'")
    Call WriteSuccessMsg("删除组成功！", "User_Friend.asp")

End Sub



Sub SaveModifyGroup()
    Dim rsUserFriendGroup, GetFriendGroup, GroupName, GroupID, i
    Dim strTemp
    strTemp = ""
    GroupID = Request("GroupID")
    If GroupID = "" Or IsNull(GroupID) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>成员组ID不能为空！</li>"
        Exit Sub
    Else
        GroupID = PE_CLng(GroupID)
    End If
    If GroupID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>默认黑名单组禁止修改！</li>"
        Exit Sub
    End If
    GroupName = ReplaceBadChar(Request("GroupName"))
    If GroupName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>组名称不能为空！</li>"
        Exit Sub
    End If
    If GetStrLen(GroupName) > 12 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>组名称不能超过6个汉字！</li>"
        Exit Sub
    End If
    Set rsUserFriendGroup = Conn.Execute("select UserFriendGroup from PE_User where UserName='" & UserName & "'")
    If rsUserFriendGroup.BOF And rsUserFriendGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户未登录或用户名错误！</li>"
        Exit Sub
    Else
        If rsUserFriendGroup(0) = "" Or IsNull(rsUserFriendGroup(0)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了网站默认组！</li>"
            Exit Sub
        Else
            GetFriendGroup = Split(rsUserFriendGroup(0), "$")
        End If
        If UBound(GetFriendGroup) < 1 Or UBound(GetFriendGroup) < GroupID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库信息错误或删除了默认组！</li>"
            Exit Sub
        Else
            For i = 0 To UBound(GetFriendGroup)
                If i = GroupID Then
                    strTemp = strTemp & "$" & GroupName
                Else
                    If strTemp = "" Then
                        strTemp = GetFriendGroup(i)
                    Else
                        strTemp = strTemp & "$" & GetFriendGroup(i)
                    End If
                End If
            Next
        End If
    End If
    Set rsUserFriendGroup = Nothing
    Conn.Execute ("update PE_User set UserFriendGroup= '" & strTemp & "' where UserName='" & UserName & "'")
    Call CloseConn
    Response.Redirect "User_Friend.asp?Action=ManageGroup"
End Sub

%>
