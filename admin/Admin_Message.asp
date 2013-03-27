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
Const PurviewLevel_Others = "Message"   '其他权限

Dim MessageID


Response.Write "<html><head><title>短消息管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("短 消 息 管 理", 10046)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Message.asp'>短消息管理首页</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_Message.asp?Action=Send'>发布网站消息</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_Message.asp?Action=BatchDel'>批量删除操作</a>&nbsp;|&nbsp;"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

MessageID = Trim(Request("MessageID"))
If IsValidID(MessageID) = False Then
    MessageID = ""
End If

strFileName = "Admin_Message.asp?Action=" & Action & "&Field=" & strField & "&keyword=" & Keyword

Select Case Action
Case "Send"
    Call Send
Case "Save"
    Call Save
Case "Read"
    Call Read
Case "BatchDel"
    Call BatchDel
Case "DelUserMessage"
    Call DelUserMessage
Case "DelChkMessage"
    Call DelChkMessage
Case "Del"
    Call Del
Case Else
    Call main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsMessage, sqlMessage
    
    Call ShowJS_Main("短消息")
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Message.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td width='30' align='center'><strong>选中</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    Response.Write "            <td width='100' align='center' ><strong>发件人</strong></td>"
    Response.Write "            <td width='100' align='center' ><strong>收件人</strong></td>"
    Response.Write "            <td align='center' ><strong>短消息主题</strong></td>"
    Response.Write "            <td width='140' align='center' ><strong>日期</strong></td>"
    Response.Write "            <td width='70' align='center' ><strong>大小</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>已读</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>操作</strong></td>"
    Response.Write "          </tr>"

    sqlMessage = "Select * From PE_Message where 1=1"
    If Keyword <> "" Then
        Select Case strField
        Case "Title"
            sqlMessage = sqlMessage & " and Title like '%" & Keyword & "%' "
        Case "Content"
            sqlMessage = sqlMessage & " and Content like '%" & Keyword & "%' "
        Case "Incept"
            sqlMessage = sqlMessage & " and Incept='" & Keyword & "' "
        Case "Sender"
            sqlMessage = sqlMessage & " and Sender='" & Keyword & "' "
        Case Else
            sqlMessage = sqlMessage & " and Title like '%" & Keyword & "%' "
        End Select
    End If
    sqlMessage = sqlMessage & " order by ID desc"

    Set rsMessage = Server.CreateObject("adodb.recordset")
    rsMessage.Open sqlMessage, Conn, 1, 1
    If rsMessage.BOF And rsMessage.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何短消息！<br><br></td></tr>"
    Else
        totalPut = rsMessage.RecordCount
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
                rsMessage.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim MessageNum
        MessageNum = 0

        Do While Not rsMessage.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'><input name='MessageID' type='checkbox' onclick='unselectall()' id='MessageID' value='" & rsMessage("ID") & "'></td>"
            Response.Write "        <td width='25' align='center'>" & rsMessage("ID") & "</td>"
            Response.Write "        <td width='100' align='center' >" & rsMessage("Sender") & "</td>"
            Response.Write "        <td width='100' align='center' >" & rsMessage("Incept") & "</td>"
            Response.Write "        <td>"
            Response.Write "<a href='Admin_Message.asp?Action=Read&MessageID=" & rsMessage("ID") & "'>"
            If rsMessage("Flag") = 1 Then
                Response.Write PE_HTMLEncode(rsMessage("Title"))
            Else
                Response.Write "<font color=blue>" & PE_HTMLEncode(rsMessage("Title")) & "</font>"
            End If
            Response.Write "</a></td>"
            Response.Write "      <td width='140' align='center'>" & rsMessage("SendTime") & "</td>"
            Response.Write "      <td width='70' align='center'>" & Len(rsMessage("Content")) & "Byte</td>"
            Response.Write "    <td width='40' align='center'>"
            If rsMessage("Flag") = 1 Then
                Response.Write "<font color=green><b>√</b></font>"
            Else
                Response.Write "<font color=red><b>×</b></font>"
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='60' align='center'>"
            Response.Write "<a href='Admin_Message.asp?Action=Del&MessageID=" & rsMessage("ID") & "' onclick=""return confirm('确定要删除此短消息吗？');"">删除</a>"
            Response.Write "</td>"
            Response.Write "</tr>"

            MessageNum = MessageNum + 1
            If MessageNum >= MaxPerPage Then Exit Do
            rsMessage.MoveNext
        Loop
    End If
    rsMessage.Close
    Set rsMessage = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='130' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中所有的短消息</td><td>"
    Response.Write "<input type='submit' value='删除选定的短消息' name='submit' onClick=""document.myform.Action.value='Del'"">&nbsp;&nbsp;"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条短消息", True)
    End If
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='90' align='right'><strong>短消息搜索：</strong></td>"
    Response.Write "   <td>" & GetMessageSearch() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub ShowJS_Send()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function SelectUser(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=UserList&DefaultValue='+document.myform.InceptUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        document.myform.InceptUser.value=arr;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.Sender.value==''){" & vbCrLf
    Response.Write "     alert('消息发送人不能为空！');" & vbCrLf
    Response.Write "     document.myform.Sender.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "     alert('短消息标题不能为空！');" & vbCrLf
    Response.Write "     document.myform.Title.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "     alert('短消息内容不能为空！');" & vbCrLf
    Response.Write "     document.myform.Content.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Send()
    Call ShowJS_Send
    Dim UserType, UserName
    UserType = PE_CLng(Trim(Request("UserType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table><br>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Message.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>发 布 网 站 消 息</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>接收方选择：</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptUserType' value='0'"
    If UserType = 0 Then Response.Write " checked"
    Response.Write "> 所有会员</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptUserType' value='1'"
    If UserType = 1 Then Response.Write " checked"
    Response.Write "> 指定会员组</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptUserType' value='2'"
    If UserType = 2 Then Response.Write " checked"
    Response.Write "> 指定用户名</td><td><input type='text' name='InceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=【<a href='#' onclick=""SelectUser();""><font color='green'>会员列表</font></a>】</font>"
    Response.Write "<br>多个用户名间请用<font color='#0000FF'>英文的逗号</font>分隔</td></tr></table>"
    
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>短消息标题：</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Title' size='66' id='Title' value=''>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>短消息内容：</td>"
    Response.Write "      <td>"
    Response.Write "        <textarea name='Content' id='Content' style='display:none'></textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=-3&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='480' height='280'></iframe>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>消息发送人：</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Sender' size='30' id='Sender' value='" & SiteName & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Save'>"
    Response.Write "        <input type='submit' name='Submit' value=' 发 布 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Save()
    Dim rs, sql
    Dim InceptUserType, inceptUser, Sender, GroupID, Title, Content
    Dim rsMessage, sqlMessage

    InceptUserType = PE_CLng(Trim(Request("InceptUserType")))
    Sender = Trim(Request("Sender"))
    inceptUser = ReplaceBadChar(Trim(Request("InceptUser")))
    GroupID = Trim(Request("GroupID"))
    Title = Trim(Request("Title"))
    Content = Trim(Request("Content"))

    Select Case InceptUserType
    Case 1
        If IsValidID(GroupID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定会员组！</li>"
        End If
    Case 2
        If inceptUser = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定接收会员！</li>"
        End If
    End Select
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>短消息主题不能为空！</li>"
    End If
    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>短消息内容不能为空！</li>"
    End If
    If Sender = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>短消息发送人不能为空！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    Sender = ReplaceBadChar(Sender)
    Title = ReplaceBadChar(Title)
    Set rsMessage = Server.CreateObject("adodb.recordset")
    sqlMessage = "select top 1 * from PE_Message"
    rsMessage.Open sqlMessage, Conn, 1, 3

    Select Case InceptUserType
    Case 0  '所有会员
        sql = "select UserName from PE_User order by UserID desc"
    Case 1  '指定会员组
        sql = "select UserName from PE_User where GroupID in (" & GroupID & ") order by UserID desc"
    Case 2  '指定会员
        inceptUser = Replace(inceptUser, ",", "','")
        sql = "select UserName from PE_User where UserName in ('" & inceptUser & "') order by UserID desc"
    End Select
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>未找到任何会员！</li>"
    Else
        Do While Not rs.EOF
            rsMessage.addnew
            rsMessage("Incept") = rs(0)
            rsMessage("Sender") = Sender
            rsMessage("Title") = Title
            rsMessage("Content") = Content
            rsMessage("SendTime") = Now()
            rsMessage("Flag") = 0
            rsMessage("IsSend") = 1
            rsMessage.Update
            Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg+1 where UserName='" & rs(0) & "'")
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    rsMessage.Close
    Set rsMessage = Nothing

    If FoundErr = True Then
        Exit Sub
    Else
        Call WriteSuccessMsg("<li><b>恭喜您，发送短信息成功。</b>", ComeUrl)
    End If
End Sub

Sub BatchDel()
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <form method='POST' name='myform' action='Admin_Message.asp' target='_self'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='DelUserMessage'>"
    Response.Write "    <tr class='topbg'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>批 量 删 除 操 作</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' height='40'><strong>批量删除会员短消息：</strong><br>可以用英文状态下的逗号将用户名隔开实现多会员同时删除</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Sender' size='32' id='Sender' value=''>&nbsp;&nbsp;"
    Response.Write "        <input name='DelUserMessage' type='submit'  id='DelUserMessage' value=' 提 交 ' onClick=""document.myform.Action.value='DelUserMessage';document.myform.target='_self';"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' height='40'><strong>批量删除指定日期范围内的短消息：</strong><br>默认为删除已读信息</td>"
    Response.Write "      <td>"
    Response.Write "        <select name='DelDate' size=1>"
    Response.Write "          <option value=1>一天前</option>"
    Response.Write "          <option value=3>三天前</option>"
    Response.Write "          <option value=7 selected>一个星期前</option>"
    Response.Write "          <option value=30>一个月前</option>"
    Response.Write "          <option value=60>两个月前</option>"
    Response.Write "          <option value=180>半年前</option>"
    Response.Write "          <option value=''>所有信息</option>"
    Response.Write "        </select>&nbsp;&nbsp;"
    Response.Write "        <input type='checkbox' name='Flag' value='0'> 包括未读信息&nbsp;&nbsp;"
    Response.Write "        <input name='DelChkMessage' type='submit'  id='DelChkMessage' value=' 提 交 ' onClick=""document.myform.Action.value='DelChkMessage';document.myform.target='_self';"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
End Sub

Sub DelUserMessage()
    Dim Sender, i, trs, tsql, Num
    Sender = Trim(Request("Sender"))
    If Sender = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入要批量删除的用户名！</li>"
        Exit Sub
    End If
    Sender = ReplaceBadChar(Sender)
    Sender = Split(Sender, ",")
    For i = 0 To UBound(Sender)
        tsql = "select incept from PE_Message where Sender='" & Sender(i) & "' and flag=0 and IsSend=1"
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open tsql, Conn, 1, 1
        Num = trs.RecordCount
        If Not trs.EOF Then
            Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg-" & Num & " where UserName='" & trs(0) & "'")
        End If
        Set trs = Nothing
        Conn.Execute ("delete from PE_Message where Sender='" & Sender(i) & "'")
    Next
    Call WriteSuccessMsg("<li><b>批量删除短信息成功。</b>", ComeUrl)
End Sub

Sub DelChkMessage()
    Dim PE_DatePart_D, strFlag, DelDate, trs, tsql
    If SystemDatabaseType = "SQL" Then
        PE_DatePart_D = "d"
    Else
        PE_DatePart_D = "'d'"
    End If
    If Trim(Request("Flag")) = "0" Then
        strFlag = ""
    Else
        strFlag = " and flag=1"
    End If
    DelDate = Trim(Request("DelDate"))
    If DelDate = "" Or Not IsNumeric(DelDate) Then
        If Trim(Request("Flag")) = "0" Then
            tsql = "select incept from PE_Message where id>0 " & strFlag & "and flag=0 and IsSend=1"
            Set trs = Server.CreateObject("adodb.recordset")
            trs.Open tsql, Conn, 1, 1
            Do While Not trs.EOF
                Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg-1 where UserName= '" & trs("incept") & "'")
                trs.MoveNext
            Loop
            Set trs = Nothing
        End If
        Conn.Execute ("delete from PE_Message where id>0 " & strFlag)
    Else
        If Trim(Request("Flag")) = "0" Then
            tsql = "select incept from PE_Message where datediff(" & PE_DatePart_D & ",sendtime," & PE_Now & ")>" & CLng(DelDate) & strFlag & " and flag=0 and IsSend=1"
            Set trs = Server.CreateObject("adodb.recordset")
            trs.Open tsql, Conn, 1, 1
            Do While Not trs.EOF
                Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg-1 where UserName= '" & trs("incept") & "'")
                trs.MoveNext
            Loop
            Set trs = Nothing
        End If
    Conn.Execute ("delete from PE_Message where datediff(" & PE_DatePart_D & ",sendtime," & PE_Now & ")>" & CLng(DelDate) & strFlag)
    End If
    Call WriteSuccessMsg("<li><b>批量删除短信息成功。</b>", ComeUrl)
End Sub

Sub Read()
    Dim rs
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>指定的短消息ID错误！</li>"
        Exit Sub
    End If
    MessageID = PE_CLng(MessageID)
    Set rs = Conn.Execute("select * from PE_Message where ID=" & MessageID)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的短消息</li>"
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>会 员 短 消 息</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td><b>发 件 人：</b>" & rs("Sender") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>收 件 人：</b>" & rs("Incept") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>消息时间：</b>" & rs("SendTime") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>消息主题：</b>" & PE_HTMLEncode(rs("Title")) & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>消息内容：</b></tr>"
    Response.Write "  <tr class='tdbg'><td>" & FilterBadTag(rs("Content"), rs("Sender")) & "</td></tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub Del()
    Dim sqlDel, rsDel, tsql, trs
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>指定的短消息ID错误！</li>"
        Exit Sub
    End If
    tsql = "select incept from PE_Message where ID in (" & MessageID & ") and flag=0 and IsSend=1"
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open tsql, Conn, 1, 1

    Do While Not trs.EOF
        Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg-1 where UserName= '" & trs("incept") & "'")
        trs.MoveNext
    Loop
    Set trs = Nothing

    Conn.Execute ("delete from PE_Message where ID in (" & MessageID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Function GetManagePath()
    Dim strPath
    strPath = "您现在的位置：短消息管理&nbsp;&gt;&gt;&nbsp;"
    If Action = "Add" Then
        strPath = strPath & "发布网站消息"
    ElseIf Action = "BatchDel" Then
        strPath = strPath & "批量删除操作"
    Else
        If Keyword = "" Then
            If Action = "Read" Then
                strPath = strPath & "阅读短消息"
            ElseIf Action = "Send" Then
                strPath = strPath & "发布短消息"
            Else
                strPath = strPath & "所有短消息"
            End If
        Else
            Select Case strField
                Case "Title"
                    strPath = strPath & "主题中含有 <font color=red>" & Keyword & "</font> "
                Case "Content"
                    strPath = strPath & "内容中含有 <font color=red>" & Keyword & "</font> "
                Case "Incept"
                    strPath = strPath & "收件人为 <font color=red>" & Keyword & "</font> "
                Case "Sender"
                    strPath = strPath & "发件人为 <font color=red>" & Keyword & "</font> "
                Case Else
                    strPath = strPath & "主题中含有 <font color=red>" & Keyword & "</font> "
            End Select
            strPath = strPath & "的短消息"
        End If
    End If
    GetManagePath = strPath
End Function

Function GetMessageSearch()
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='Admin_Message.asp'>"
    strForm = strForm & "<tr><td height='28' align='center'>"
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='Title' selected>短消息主题</option>"
    strForm = strForm & "<option value='Content'>短消息内容</option>"
    strForm = strForm & "<option value='Incept'>收件人</option>"
    strForm = strForm & "<option value='Sender'>发件人</option>"
    strForm = strForm & "</select>"
    strForm = strForm & " <input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    strForm = strForm & " <input type='submit' name='Submit'  value='搜索'>"
    strForm = strForm & "</td></tr></form></table>"
    GetMessageSearch = strForm
End Function
%>
