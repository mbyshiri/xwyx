<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Message"   '����Ȩ��

Dim MessageID


Response.Write "<html><head><title>����Ϣ����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� �� Ϣ �� ��", 10046)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Message.asp'>����Ϣ������ҳ</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_Message.asp?Action=Send'>������վ��Ϣ</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_Message.asp?Action=BatchDel'>����ɾ������</a>&nbsp;|&nbsp;"
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
    
    Call ShowJS_Main("����Ϣ")
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Message.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    Response.Write "            <td width='100' align='center' ><strong>������</strong></td>"
    Response.Write "            <td width='100' align='center' ><strong>�ռ���</strong></td>"
    Response.Write "            <td align='center' ><strong>����Ϣ����</strong></td>"
    Response.Write "            <td width='140' align='center' ><strong>����</strong></td>"
    Response.Write "            <td width='70' align='center' ><strong>��С</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>�Ѷ�</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>����</strong></td>"
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
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>û���κζ���Ϣ��<br><br></td></tr>"
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
                Response.Write "<font color=green><b>��</b></font>"
            Else
                Response.Write "<font color=red><b>��</b></font>"
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='60' align='center'>"
            Response.Write "<a href='Admin_Message.asp?Action=Del&MessageID=" & rsMessage("ID") & "' onclick=""return confirm('ȷ��Ҫɾ���˶���Ϣ��');"">ɾ��</a>"
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
    Response.Write "    <td width='130' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�����еĶ���Ϣ</td><td>"
    Response.Write "<input type='submit' value='ɾ��ѡ���Ķ���Ϣ' name='submit' onClick=""document.myform.Action.value='Del'"">&nbsp;&nbsp;"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������Ϣ", True)
    End If
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='90' align='right'><strong>����Ϣ������</strong></td>"
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
    Response.Write "     alert('��Ϣ�����˲���Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.Sender.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "     alert('����Ϣ���ⲻ��Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.Title.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "     alert('����Ϣ���ݲ���Ϊ�գ�');" & vbCrLf
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
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� վ �� Ϣ</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>���շ�ѡ��</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptUserType' value='0'"
    If UserType = 0 Then Response.Write " checked"
    Response.Write "> ���л�Ա</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptUserType' value='1'"
    If UserType = 1 Then Response.Write " checked"
    Response.Write "> ָ����Ա��</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptUserType' value='2'"
    If UserType = 2 Then Response.Write " checked"
    Response.Write "> ָ���û���</td><td><input type='text' name='InceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=��<a href='#' onclick=""SelectUser();""><font color='green'>��Ա�б�</font></a>��</font>"
    Response.Write "<br>����û���������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr></table>"
    
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>����Ϣ���⣺</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Title' size='66' id='Title' value=''>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>����Ϣ���ݣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <textarea name='Content' id='Content' style='display:none'></textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=-3&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='480' height='280'></iframe>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>��Ϣ�����ˣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Sender' size='30' id='Sender' value='" & SiteName & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Save'>"
    Response.Write "        <input type='submit' name='Submit' value=' �� �� '>"
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
            ErrMsg = ErrMsg & "<li>��ָ����Ա�飡</li>"
        End If
    Case 2
        If inceptUser = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ�����ջ�Ա��</li>"
        End If
    End Select
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ⲻ��Ϊ�գ�</li>"
    End If
    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ݲ���Ϊ�գ�</li>"
    End If
    If Sender = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ�����˲���Ϊ�գ�</li>"
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
    Case 0  '���л�Ա
        sql = "select UserName from PE_User order by UserID desc"
    Case 1  'ָ����Ա��
        sql = "select UserName from PE_User where GroupID in (" & GroupID & ") order by UserID desc"
    Case 2  'ָ����Ա
        inceptUser = Replace(inceptUser, ",", "','")
        sql = "select UserName from PE_User where UserName in ('" & inceptUser & "') order by UserID desc"
    End Select
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>δ�ҵ��κλ�Ա��</li>"
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
        Call WriteSuccessMsg("<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b>", ComeUrl)
    End If
End Sub

Sub BatchDel()
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <form method='POST' name='myform' action='Admin_Message.asp' target='_self'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='DelUserMessage'>"
    Response.Write "    <tr class='topbg'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� ɾ �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' height='40'><strong>����ɾ����Ա����Ϣ��</strong><br>������Ӣ��״̬�µĶ��Ž��û�������ʵ�ֶ��Աͬʱɾ��</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Sender' size='32' id='Sender' value=''>&nbsp;&nbsp;"
    Response.Write "        <input name='DelUserMessage' type='submit'  id='DelUserMessage' value=' �� �� ' onClick=""document.myform.Action.value='DelUserMessage';document.myform.target='_self';"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' height='40'><strong>����ɾ��ָ�����ڷ�Χ�ڵĶ���Ϣ��</strong><br>Ĭ��Ϊɾ���Ѷ���Ϣ</td>"
    Response.Write "      <td>"
    Response.Write "        <select name='DelDate' size=1>"
    Response.Write "          <option value=1>һ��ǰ</option>"
    Response.Write "          <option value=3>����ǰ</option>"
    Response.Write "          <option value=7 selected>һ������ǰ</option>"
    Response.Write "          <option value=30>һ����ǰ</option>"
    Response.Write "          <option value=60>������ǰ</option>"
    Response.Write "          <option value=180>����ǰ</option>"
    Response.Write "          <option value=''>������Ϣ</option>"
    Response.Write "        </select>&nbsp;&nbsp;"
    Response.Write "        <input type='checkbox' name='Flag' value='0'> ����δ����Ϣ&nbsp;&nbsp;"
    Response.Write "        <input name='DelChkMessage' type='submit'  id='DelChkMessage' value=' �� �� ' onClick=""document.myform.Action.value='DelChkMessage';document.myform.target='_self';"" style='cursor:hand;'>"
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
        ErrMsg = ErrMsg & "<li>������Ҫ����ɾ�����û�����</li>"
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
    Call WriteSuccessMsg("<li><b>����ɾ������Ϣ�ɹ���</b>", ComeUrl)
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
    Call WriteSuccessMsg("<li><b>����ɾ������Ϣ�ɹ���</b>", ComeUrl)
End Sub

Sub Read()
    Dim rs
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
        Exit Sub
    End If
    MessageID = PE_CLng(MessageID)
    Set rs = Conn.Execute("select * from PE_Message where ID=" & MessageID)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ķ���Ϣ</li>"
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
    Response.Write "    <td height='22' align='center'><strong>�� Ա �� �� Ϣ</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td><b>�� �� �ˣ�</b>" & rs("Sender") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>�� �� �ˣ�</b>" & rs("Incept") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>��Ϣʱ�䣺</b>" & rs("SendTime") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>��Ϣ���⣺</b>" & PE_HTMLEncode(rs("Title")) & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>��Ϣ���ݣ�</b></tr>"
    Response.Write "  <tr class='tdbg'><td>" & FilterBadTag(rs("Content"), rs("Sender")) & "</td></tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub Del()
    Dim sqlDel, rsDel, tsql, trs
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
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
    strPath = "�����ڵ�λ�ã�����Ϣ����&nbsp;&gt;&gt;&nbsp;"
    If Action = "Add" Then
        strPath = strPath & "������վ��Ϣ"
    ElseIf Action = "BatchDel" Then
        strPath = strPath & "����ɾ������"
    Else
        If Keyword = "" Then
            If Action = "Read" Then
                strPath = strPath & "�Ķ�����Ϣ"
            ElseIf Action = "Send" Then
                strPath = strPath & "��������Ϣ"
            Else
                strPath = strPath & "���ж���Ϣ"
            End If
        Else
            Select Case strField
                Case "Title"
                    strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> "
                Case "Content"
                    strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> "
                Case "Incept"
                    strPath = strPath & "�ռ���Ϊ <font color=red>" & Keyword & "</font> "
                Case "Sender"
                    strPath = strPath & "������Ϊ <font color=red>" & Keyword & "</font> "
                Case Else
                    strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> "
            End Select
            strPath = strPath & "�Ķ���Ϣ"
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
    strForm = strForm & "<option value='Title' selected>����Ϣ����</option>"
    strForm = strForm & "<option value='Content'>����Ϣ����</option>"
    strForm = strForm & "<option value='Incept'>�ռ���</option>"
    strForm = strForm & "<option value='Sender'>������</option>"
    strForm = strForm & "</select>"
    strForm = strForm & " <input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    strForm = strForm & " <input type='submit' name='Submit'  value='����'>"
    strForm = strForm & "</td></tr></form></table>"
    GetMessageSearch = strForm
End Function
%>
