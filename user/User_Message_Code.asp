<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const MaxMessageNum = 100           '�û������ն���Ϣ������������ϵͳ���Զ�ɾ��
Const MaxTitleLength = 50           '����Ϣ������󳤶�
Const MaxContentLength = 1000       '����Ϣ������󳤶�

Dim rs, sql, Passed
Dim MessageID, ManageType, BoxName, ActionName, MessageCount

Sub Execute()
    MessageID = Trim(Request("MessageID"))
    ManageType = Trim(Request("ManageType"))
    Passed = Trim(Request("Passed"))
    If Passed = "" Then
        Passed = Session("Passed")
    End If
    If Passed = "" Then
        Passed = "All"
    End If
    Session("Passed") = Passed


    Select Case ManageType
    Case "Inbox"
        BoxName = "�ռ���"
    Case "Outbox"
        BoxName = "�ݸ���"
    Case "IsSend"
        BoxName = "�ѷ���"
    Case "Recycle"
        BoxName = "�ϼ���"
    Case Else
        BoxName = "�ռ���"
        ManageType = "Inbox"
    End Select

    If Action = "" Then Action = "Manage"

    FileName = "User_Message.asp?Action=" & Action & "&ManageType=" & ManageType
    strFileName = FileName & "&Field=" & strField & "&keyword=" & Keyword

    Call DelOutMessage
    Select Case Action
    Case "New", "Edit", "Re", "Fw"
        ActionName = "д����"
        Call SendMessage
    Case "SendMessage", "SaveMessage"
        ActionName = "���Ͷ���"
        Call SaveMessage
    Case "SendEdit", "SaveEdit"
        ActionName = "�����������"
        Call SaveEdit
    Case "ReadInbox", "ReadOther"
        ActionName = "�Ķ�����Ϣ"
        Call Read
    Case "Del", "Delete"
        ActionName = "ɾ������"
        Call Del
    Case "Clear"
        ActionName = "����ռ���"
        Call Clear
    Case "Manage"
        Call main
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub main()
    Call ShowJS_Main("����Ϣ")
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetPath() & "</td>"
    Response.Write "    <td width='100'>" & GetSpace() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_Message.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>ѡ��</strong></td>"
    'Response.write "            <td width='25' align='center'><strong>ID</strong></td>"
    If ManageType = "Inbox" Or ManageType = "Recycle" Then
        Response.Write "            <td width='120' align='center' ><strong>������</strong></td>"
    Else
        Response.Write "            <td width='120' align='center' ><strong>�ռ���</strong></td>"
    End If
    Response.Write "            <td align='center' ><strong>����Ϣ����</strong></td>"
    Response.Write "            <td width='140' align='center' ><strong>����</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>��С</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>�Ѷ�</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>����</strong></td>"
    Response.Write "          </tr>"

    sql = "Select * From PE_Message"
    Select Case ManageType
    Case "Inbox"
        sql = sql & " where IsSend = 1 and DelR = 0 and Incept = '" & UserName & "'"
    Case "Outbox"
        sql = sql & " where Sender = '" & UserName & "' and IsSend = 0 and delS = 0"
    Case "IsSend"
        sql = sql & " where Sender = '" & UserName & "' and IsSend = 1 and delS = 0"
    Case "Recycle"
        sql = sql & " where ((Sender = '" & UserName & "' and delS = 1) or (Incept = '" & UserName & "' and DelR = 1))"
    Case Else
        sql = sql & " where IsSend = 1 and DelR = 0 and Incept = '" & UserName & "'"
    End Select
    If Keyword <> "" Then
        Select Case strField
        Case "Title"
            sql = sql & " and Title like '%" & Keyword & "%' "
        Case "Content"
            sql = sql & " and Content like '%" & Keyword & "%' "
        Case Else
            sql = sql & " and Title like '%" & Keyword & "%' "
        End Select
    End If
    If Passed = "True" And ManageType = "Inbox" And Action = "Manage" Then
        sql = sql & " and flag =" & PE_True & ""
    ElseIf Passed = "False" And ManageType = "Inbox" And Action = "Manage" Then
        sql = sql & " and flag =" & PE_False & ""
    End If
    Select Case ManageType
    Case "Inbox"
        sql = sql & " order by Flag,ID desc"
    Case "Outbox", "IsSend", "Recycle"
        sql = sql & " order by ID desc"
    Case Else
        sql = sql & " order by Flag,ID desc"
    End Select

    Dim rsMessage
    Set rsMessage = Server.CreateObject("ADODB.Recordset")
    rsMessage.Open sql, Conn, 1, 1
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
            'Response.write "        <td width='25' align='center'>" & rsMessage("ID") & "</td>"
            If ManageType = "Inbox" Or ManageType = "Recycle" Then
                Response.Write "        <td width='120' align='center' >" & rsMessage("Sender") & "</td>"
            Else
                Response.Write "        <td width='120' align='center' >" & rsMessage("Incept") & "</td>"
            End If
            Response.Write "        <td>"
            Select Case ManageType
            Case "Inbox"
                Response.Write "<a href='User_Message.asp?Action=ReadInbox&MessageID=" & rsMessage("ID") & "'>"
            Case "Outbox"
                Response.Write "<a href='User_Message.asp?Action=Edit&MessageID=" & rsMessage("ID") & "'>"
            Case Else
                Response.Write "<a href='User_Message.asp?Action=ReadOther&MessageID=" & rsMessage("ID") & "'>"
            End Select
            If rsMessage("Flag") = 1 Then
                Response.Write PE_HTMLEncode(rsMessage("Title"))
            Else
                Response.Write "<font color=blue>" & PE_HTMLEncode(rsMessage("Title")) & "</font>"
            End If
            Response.Write "</a></td>"
            Response.Write "      <td width='140' align='center'>" & rsMessage("SendTime") & "</td>"
            Response.Write "      <td width='80' align='center'>" & Len(rsMessage("Content")) & "Byte</td>"
            Response.Write "    <td width='40' align='center'>"
            If rsMessage("Flag") = 1 Then
                Response.Write "<font color=green><b>��</b></font>"
            Else
                Response.Write "<font color=red><b>��</b></font>"
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='80' align='center'>"
            If ManageType = "Recycle" Then
                Response.Write "<a href='User_Message.asp?Action=Del&ManageType=" & ManageType & "&MessageID=" & rsMessage("ID") & "' onclick=""return confirm('ȷ��Ҫɾ���˶���Ϣ��ɾ������Ϣ�����ɻָ���');"">ɾ��</a>"
            Else
                Response.Write "<a href='User_Message.asp?Action=Del&ManageType=" & ManageType & "&MessageID=" & rsMessage("ID") & "' onclick=""return confirm('ȷ��Ҫɾ���˶���Ϣ��');"">ɾ��</a>"
            End If
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
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�б�ҳ��ʾ�����ж���Ϣ</td><td>"
    Response.Write "<input name='submit1' type='submit' value='ɾ��ѡ���Ķ���Ϣ' onClick=""document.myform.Action.value='Del'"" >"
    Response.Write "&nbsp;&nbsp;<input name='submit1' type='submit' value='���" & BoxName & "' onClick=""document.myform.Action.value='Clear'"" >"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������Ϣ", True)

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>����Ϣ������</strong></td>"
    Response.Write "   <td>" & GetSearchForm() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub


Sub SendMessage()
    If MaxSendNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ�����û�з��Ͷ���Ϣ��Ȩ�ޣ�"
        Exit Sub
    End If
    If MessageID <> "" And IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
        Exit Sub
    End If
    Response.Cookies("SendMessage") = "No"

    Dim inceptUser, Sender, SendTime, Title, Content, i
    Dim chatloglist
    MessageID = PE_CLng(MessageID)
    inceptUser = Request("inceptUser")
    Select Case Action
    Case "Edit"
        sql = "Select * from PE_Message where Sender='" & UserName & "' and IsSend=0 and ID=" & MessageID
    Case "Re"
        sql = "SELECT * from PE_Message where Incept='" & UserName & "' and ID=" & MessageID
    Case "Fw"
        sql = "SELECT * from PE_Message where (Incept='" & UserName & "' or Sender='" & UserName & "') and ID=" & MessageID
    End Select

    If MessageID <> "" And IsNumeric(MessageID) And sql <> "" Then
        Set rs = Conn.Execute(sql)
        If Not (rs.BOF And rs.EOF) Then
            Sender = rs("Sender")
            SendTime = rs("SendTime")
            Select Case Action
            Case "Re"
                inceptUser = rs("Sender")
                Title = "Re: " & rs("Title")
                Content = Content & "======�� " & SendTime & " ��������д����======" & "<br>"
                Content = Content & rs("Content") & "<br>"
                Content = Content & "================================================" & "<br>"
            Case "Fw"
                Title = "Fw: " & rs("Title")
                Content = Content & "============== ������ת����Ϣ ==============" & "<br>"
                Content = Content & "ԭ�����ˣ�" & Sender & " " & "<br>"
                Content = Content & "ԭ�������ݣ�" & "<br>"
                Content = Content & rs("Content") & "<br>"
                Content = Content & "============================================" & "<br>"
            Case "Edit"
                inceptUser = rs("Incept")
                Title = rs("Title")
                Content = rs("Content")
            End Select
            Content = Server.HTMLEncode(Content)
        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������</li>"
            Set rs = Nothing
            Exit Sub
        End If
        Set rs = Nothing
    End If

    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function SelectFromFriend(){" & vbCrLf
    Response.Write "var str1=document.myform.InceptUser.value;" & vbCrLf
    Response.Write "var str2=document.myform.FriendList.value;" & vbCrLf
    Response.Write "if (document.myform.FriendList.value!=''){" & vbCrLf
    Response.Write "   if (str1==''){" & vbCrLf
    Response.Write "       document.myform.InceptUser.value=str2;" & vbCrLf
    Response.Write "   }" & vbCrLf
    Response.Write "   else{" & vbCrLf
    Response.Write "       if (checkFriend(str1,str2))" & vbCrLf
    Response.Write "       {" & vbCrLf
    Response.Write "       document.myform.InceptUser.value=str1+','+str2;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "   }" & vbCrLf
    Response.Write "   }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function checkFriend(friendlist,thisfriend){" & vbCrLf
    Response.Write "   if(friendlist==thisfriend)" & vbCrLf
    Response.Write "       {" & vbCrLf
    Response.Write "       return false;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "   else" & vbCrLf
    Response.Write "       {" & vbCrLf
    Response.Write "       var str=friendlist.split("","");" & vbCrLf
    Response.Write "       for(i=0;i<str.length;i++)" & vbCrLf
    Response.Write "           {" & vbCrLf
    Response.Write "           if(str[i]==thisfriend)" & vbCrLf
    Response.Write "               return false;   " & vbCrLf
    Response.Write "           }" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "}" & vbCrLf
    'Response.Write "function SelectUser(){" & vbCrLf
    'Response.Write "    var arr=showModalDialog('User_SourceList.asp?TypeSelect=UserList&DefaultValue='+document.myform.InceptUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    'Response.Write "    if (arr != null){" & vbCrLf
    'Response.Write "        document.myform.InceptUser.value=arr;" & vbCrLf
    'Response.Write "    }" & vbCrLf
    'Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.InceptUser.value==''){" & vbCrLf
    Response.Write "     alert('�ռ��˲���Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.InceptUser.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "     alert('����Ϣ���ⲻ��Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.Title.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "       document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "       document.myform.Content.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "     alert('����Ϣ���ݲ���Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.Content.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Message.asp' target='_self'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center' colspan='2'><strong>׫ д �� �� Ϣ</strong></td>"
    Response.Write "  </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='20%' align='right'>�ռ��ˣ�</td>"
    Response.Write "      <td width='80%'>"
    Response.Write "        <input type='text' name='InceptUser' size='52' id='InceptUser' value='" & inceptUser & "'>"
    Response.Write "      <select name='FriendList' onchange=""SelectFromFriend();"">"
    Response.Write "      <option value=''>��ѡ��...</option>"
    Response.Write GetFriendListOption
    Response.Write "      </select>"
    'Response.Write "       ��<a href='#' onclick=""SelectUser();"">��Ա�б�</a>��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>���⣺</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Title' size='66' id='Title' value='" & Title & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>���ݣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <textarea name='Content' id='Content' style='display:none'>" & Content & "</textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='485' height='280' ></iframe>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Send'>"
    If Action = "Edit" Then
        Response.Write "        <input name='Send' type='submit'  id='Send' value=' �� �� ' onClick=""document.myform.Action.value='SendEdit';document.myform.target='_self';"" style='cursor:hand;'>&nbsp; "
        Response.Write "        <input name='Save' type='submit'  id='Save' value=' �� �� ' onClick=""document.myform.Action.value='SaveEdit';document.myform.target='_self';"" style='cursor:hand;'>"
        Response.Write "   <input name='MessageID' type='hidden' id='MessageID' value='" & MessageID & "'>"
    Else
        Response.Write "        <input name='Send' type='submit'  id='Send' value=' �� �� ' onClick=""document.myform.Action.value='SendMessage';document.myform.target='_self';"" style='cursor:hand;'>&nbsp; "
        Response.Write "        <input name='Save' type='submit'  id='Save' value=' �� �� ' onClick=""document.myform.Action.value='SaveMessage';document.myform.target='_self';"" style='cursor:hand;'>"
    End If
    Response.Write "        <input type='reset' name='Clear' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td  colspan='2'>1��������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������<b>" & MaxSendNum & "</b>���û���<br>2�� �������<b>" & MaxTitleLength & "</b>���ַ����������<b>" & MaxContentLength & "</b>���ַ�</td>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub SaveMessage()
    If MaxSendNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ�����û�з��Ͷ���Ϣ��Ȩ�ޣ�"
        Exit Sub
    End If
    
    Dim rsMessage, sqlMessage, incept, Title, Content
    incept = Trim(Request("InceptUser"))
    Title = Trim(Request("Title"))
    
    For i = 1 To Request.Form("Content").Count
        Content = Content & FilterJS(Request.Form("Content")(i))
    Next
    
    'Content = Trim(Request("Content"))
    If Request.Cookies("SendMessage") = "Yes" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�벻Ҫ����������ͬ�Ķ���Ϣ��</li>"
        Exit Sub
    End If
    If incept = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ռ��˲���Ϊ�գ�</li>"
    End If
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ⲻ��Ϊ�գ�</li>"
    ElseIf Len(Title) > MaxTitleLength Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���������ӦС��" & MaxTitleLength & "����</li>"
    End If
    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ݲ���Ϊ�գ�</li>"
    ElseIf Len(Content) > MaxContentLength Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ݹ�����ӦС��" & MaxContentLength & "����</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    incept = ReplaceBadChar(incept)
    Title = ReplaceBadChar(Title)
    
   ' For i = 1 To Request.Form("Content").Count
   '     Content = Content & FilterJS(Request.Form("Content")(i))
   ' Next
    'Content = ReplaceBadUrl(Content)
    Set rsMessage = Server.CreateObject("adodb.recordset")
    sqlMessage = "select top 1 * from PE_Message"
    rsMessage.Open sqlMessage, Conn, 1, 3
    If Action = "SaveMessage" Then
        rsMessage.addnew
        rsMessage("Incept") = incept
        rsMessage("Sender") = UserName
        rsMessage("Title") = Title
        rsMessage("Content") = Content
        rsMessage("SendTime") = Now()
        rsMessage("Flag") = 0
        rsMessage("IsSend") = 0
        rsMessage.Update
        Call WriteSuccessMsg("<li><b>��ϲ�����������Ϣ�ɹ���</b><br>����Ϣ���������Ĳݸ����С�", ComeUrl)
    Else
        incept = Split(incept, ",")
        Dim strTemp, i
        For i = 0 To UBound(incept)
            If strTemp = "" Then
                strTemp = incept(i)
            Else
                If FoundInArr(strTemp, incept(i), ",") = False And incept(i) <> UserName Then
                    strTemp = strTemp & "," & incept(i)
                End If
            End If
        Next
        incept = Split(strTemp, ",")
        For i = 0 To UBound(incept)
            If i >= MaxSendNum Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���ֻ�ܷ��͸�" & MaxSendNum & "���û�����������" & MaxSendNum & "λ�Ժ�������·��ͣ�</li>"
                Exit For
            End If
            Set rs = Conn.Execute("select UserName from PE_User where UserName='" & Replace(incept(i), "'", "") & "'")
            If rs.BOF And rs.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�޴��û�--" & incept(i) & "�������ռ����Ƿ���д��ȷ��</li>"
                Set rs = Nothing
                rsMessage.Close
                Set rsMessage = Nothing
                Exit Sub
            End If
            Set rs = Nothing
            If CheckBlackFriend(incept(i)) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���<font color='red'>" & incept(i) & "</font>�����˺�����������<font color='red'>" & incept(i) & "</font>���������˺���������˶��ŷ��ͱ���ֹ��</li>"
                Exit Sub
            End If
            rsMessage.addnew
            rsMessage("Incept") = incept(i)
            rsMessage("Sender") = UserName
            rsMessage("Title") = Title
            rsMessage("Content") = Content
            rsMessage("SendTime") = Now()
            rsMessage("Flag") = 0
            rsMessage("IsSend") = 1
            rsMessage.Update
            '�����û�����Ϣ����
            Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg+1 where UserName='" & incept(i) & "'")
        Next
        Call WriteSuccessMsg("<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���Ͷ���Ϣͬʱ�����������ѷ�����Ϣ�С�", ComeUrl)
    End If
    rsMessage.Close
    Set rsMessage = Nothing
    Response.Cookies("SendMessage") = "Yes"
End Sub

Sub SaveEdit()
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
        Exit Sub
    End If
    If MaxSendNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ�����û�з��Ͷ���Ϣ��Ȩ�ޣ�"
        Exit Sub
    End If

    Dim rsMessage, sqlMessage, incept, Title, Content
    incept = Trim(Request("Incept"))
    Title = Trim(Request("Title"))
    
    'Content = Trim(Request("Content"))
    For i = 1 To Request.Form("Content").Count
        Content = Content & FilterJS(Request.Form("Content")(i))
    Next
    
    If incept = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ռ��˲���Ϊ�գ�</li>"
    End If
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ⲻ��Ϊ�գ�</li>"
    ElseIf Len(Title) > MaxTitleLength Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���������ӦС��" & MaxTitleLength & "����</li>"
    End If
    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ݲ���Ϊ�գ�</li>"
    ElseIf Len(Content) > MaxContentLength Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ϣ���ݹ�����ӦС��" & MaxContentLength & "����</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    incept = ReplaceBadChar(incept)
    Title = ReplaceBadChar(Title)
    Content = ReplaceBadUrl(FilterJS(Content))
    If Action = "SaveEdit" Then
        Set rsMessage = Server.CreateObject("adodb.recordset")
        sqlMessage = "select * from PE_Message where ID=" & PE_CLng(MessageID) & " and Sender='" & UserName & "'"
        rsMessage.Open sqlMessage, Conn, 1, 3
        If Not (rsMessage.BOF And rsMessage.EOF) Then
            rsMessage("Incept") = incept
            rsMessage("Title") = Title
            rsMessage("Content") = Content
            rsMessage("SendTime") = Now()
            rsMessage("Flag") = 0
            rsMessage("IsSend") = 0
            rsMessage.Update
        End If
        rsMessage.Close
        Set rsMessage = Nothing
        Call WriteSuccessMsg("<li><b>��ϲ�����������Ϣ�ɹ���</b><br>����Ϣ���������Ĳݸ����С�", ComeUrl)
    Else
        Conn.Execute ("delete from PE_Message where ID=" & PE_CLng(MessageID) & " and Sender='" & UserName & "'")
        incept = Split(incept, ",")
        Dim strTemp
        For i = 0 To UBound(incept)
            If strTemp = "" Then
                strTemp = incept(i)
            Else
                If FoundInArr(strTemp, incept(i), ",") = False And incept(i) <> UserName Then
                    strTemp = strTemp & "," & incept(i)
                End If
            End If
        Next
        incept = Split(strTemp, ",")
        Set rsMessage = Server.CreateObject("adodb.recordset")
        sqlMessage = "select top 1 * from PE_Message"
        rsMessage.Open sqlMessage, Conn, 1, 3
        For i = 0 To UBound(incept)
            If i >= MaxSendNum Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���ֻ�ܷ��͸�" & MaxSendNum & "���û�����������" & MaxSendNum & "λ�Ժ�������·��ͣ�</li>"
                Exit For
            End If
            Set rs = Conn.Execute("select UserName from PE_User where UserName='" & Replace(incept(i), "'", "") & "'")
            If rs.BOF And rs.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�޴��û�--" & incept(i) & "�������ռ����Ƿ���д��ȷ��</li>"
                Set rs = Nothing
                rsMessage.Close
                Set rsMessage = Nothing
                Exit Sub
            End If
            Set rs = Nothing
            rsMessage.addnew
            rsMessage("Incept") = incept(i)
            rsMessage("Sender") = UserName
            rsMessage("Title") = Title
            rsMessage("Content") = Content
            rsMessage("SendTime") = Now()
            rsMessage("Flag") = 0
            rsMessage("IsSend") = 1
            rsMessage.Update
            '�����û�����Ϣ����
            Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg+1 where UserName='" & incept(i) & "'")
        Next
        rsMessage.Close
        Set rsMessage = Nothing
        Call WriteSuccessMsg("<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���Ͷ���Ϣͬʱ�����������ѷ�����Ϣ�С�", ComeUrl)
    End If
End Sub

Sub Read()
    Dim NextID, NextSender
    
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
        Exit Sub
    End If
    MessageID = PE_CLng(MessageID)
    
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_Message where (Incept='" & UserName & "' or Sender='" & UserName & "') and ID=" & MessageID
    rs.Open sql, Conn, 1, 3
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ķ���Ϣ</li>"
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetPath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>�� �� �� �� Ϣ</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='center'>"
    Response.Write "      <a href='User_Message.asp?Action=Delete&MessageID=" & rs("ID") & "'><img src='images/m_delete.gif' border=0 alt='ɾ����Ϣ'></a> &nbsp; "
    Response.Write "      <a href='User_Message.asp?Action=New'><img src='images/m_to.gif' border=0 alt='������Ϣ'></a> &nbsp;"
    Response.Write "      <a href='User_Message.asp?Action=Re&touser={$sender}&MessageID=" & rs("ID") & "'><img src='images/m_re.gif' border=0 alt='�ظ���Ϣ'></a>&nbsp;"
    Response.Write "      <a href='User_Message.asp?Action=Fw&MessageID=" & rs("ID") & "'><img src='images/m_fw.gif' border=0 alt='ת����Ϣ'></a>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td><b>�� �� �ˣ�</b>" & rs("Sender") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>����ʱ�䣺</b>" & rs("SendTime") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>��Ϣ���⣺</b>" & PE_HTMLEncode(rs("Title")) & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td>" & FilterBadTag(rs("Content"), rs("Sender")) & "</td></tr>"

    If UserName <> rs("Sender") Then
        If rs("Flag") = 0 Then
            rs("Flag") = 1
            rs.Update
            Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg-1 where UserName='" & UserName & "'")
        End If
    End If
    rs.Close
    Set rs = Nothing

    Set rs = Conn.Execute("select ID,Sender from PE_Message where Incept='" & UserName & "' and Flag=0 and IsSend=1 order by SendTime")
    If Not (rs.BOF And rs.EOF) Then
        NextID = rs(0)
        NextSender = rs(1)
    End If
    Set rs = Nothing

    If Action = "ReadInbox" And NextID <> "" Then
        Response.Write "  <tr class='tdbg'><td align='right'>"
        Response.Write "   <a href=User_Message.asp?Action=ReadInbox&MessageID=" & NextID & ">[��ȡ��һ����Ϣ]</a>"
        Response.Write "  </td></tr>"
    End If
    Response.Write "</table>"
End Sub

Sub Del()
    If IsValidID(MessageID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
        Exit Sub
    End If
    If Action = "Delete" Then
        Conn.Execute ("delete from PE_Message where Incept='" & UserName & "' and DelR=1 and ID in (" & MessageID & ")")
        Conn.Execute ("delete from PE_Message where Sender='" & UserName & "' and DelS=1 and IsSend=0 and ID in (" & MessageID & ")")
        Conn.Execute ("update PE_Message set DelS=2 where Sender='" & UserName & "' and DelS=1 and IsSend=1 and ID in (" & MessageID & ")")
        Conn.Execute ("update PE_Message set DelR=1 where Incept='" & UserName & "' and ID in (" & MessageID & ")")
        Conn.Execute ("update PE_Message set DelS=1 where Sender='" & UserName & "' and ID in (" & MessageID & ")")
    Else
        Select Case ManageType
        Case "Inbox"
            Conn.Execute ("update PE_Message set DelR=1 where Incept='" & UserName & "' and ID in (" & MessageID & ")")
        Case "Outbox"
            Conn.Execute ("update PE_Message set DelS=1 where Sender='" & UserName & "' and IsSend=0 and ID in (" & MessageID & ")")
        Case "IsSend"
            Conn.Execute ("update PE_Message set DelS=1 where Sender='" & UserName & "' and IsSend=1 and ID in (" & MessageID & ")")
        Case "Recycle"
            Conn.Execute ("delete from PE_Message where Incept='" & UserName & "' and DelR=1 and ID in (" & MessageID & ")")
            Conn.Execute ("delete from PE_Message where Sender='" & UserName & "' and DelS=1 and IsSend=0 and ID in (" & MessageID & ")")
            Conn.Execute ("update PE_Message set DelS=2 where Sender='" & UserName & "' and DelS=1 and IsSend=1 and ID in (" & MessageID & ")")
        End Select
    End If
    Update_User_Message (UserName)
    If Action = "Delete" Or ManageType = "Recycle" Then
        Call WriteSuccessMsg("<li>ɾ������Ϣ�ɹ���</li>", ComeUrl)
    Else
        Call WriteSuccessMsg("<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��</li>", ComeUrl)
    End If
End Sub

Sub Clear()
    Select Case ManageType
    Case "Inbox"
        Conn.Execute ("update PE_Message set DelR=1 where Incept='" & UserName & "' and DelR=0")
    Case "Outbox"
        Conn.Execute ("update PE_Message Set DelS=1 where Sender='" & UserName & "' and DelS=0 and IsSend=0")
    Case "IsSend"
        Conn.Execute ("update PE_Message Set DelS=1 where Sender='" & UserName & "' and DelS=0 and IsSend=1")
    Case "Recycle"
        Conn.Execute ("delete from PE_Message where Incept='" & UserName & "' and DelR=1")
        Conn.Execute ("delete from PE_Message where Sender='" & UserName & "' and DelS=1 and IsSend=0")
        Conn.Execute ("update PE_Message set DelS=2 where Sender='" & UserName & "' and DelS=1 and IsSend=1")
    Case Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ô���</li>"
        Exit Sub
    End Select
    Update_User_Message (UserName)
    If ManageType = "Recycle" Then
        Call WriteSuccessMsg("<li>ɾ������Ϣ�ɹ���</li>", ComeUrl)
    Else
        Call WriteSuccessMsg("<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��</li>", ComeUrl)
    End If
End Sub

Sub Update_User_Message(incept)
    Dim trs
    Set trs = Conn.Execute("select Count(Id) from PE_Message where incept='" & incept & "'and flag=0 and DelR=0")
    If trs(0) = 0 Then
        Conn.Execute ("update PE_User set UnReadMsg=0 where UserName='" & incept & "'")
    Else
        Conn.Execute ("update PE_User set UnReadMsg=" & trs(0) & " where UserName='" & incept & "'")
    End If
End Sub
Sub DelOutMessage()
    Dim OutNum
    MessageCount = 0
    Set rs = Conn.Execute("select count(ID) From PE_Message where Incept='" & UserName & "'")
    MessageCount = rs(0)
    If MessageCount > MaxMessageNum Then
        OutNum = MessageCount - MaxMessageNum
        Set rs = Conn.Execute("select top " & OutNum & " ID From PE_Message where Incept='" & UserName & "' order by ID Asc,DelR Desc")
        While Not rs.EOF
            Conn.Execute ("delete from PE_Message where ID=" & rs(0))
            rs.MoveNext
        Wend
        MessageCount = MaxMessageNum
    End If
    rs.Close
    Set rs = Nothing
End Sub

Function GetSpace()
    Dim tmpSpace, SpacePercent, strSpace
    If MaxMessageNum > 0 Then
        strSpace = strSpace & "�ռ�ʹ�ã� "
        If FormatNumber(MessageCount / MaxMessageNum * 100, 0, -1) < 50 Then
            strSpace = strSpace & "<font color='green'>" & FormatPercent(MessageCount / MaxMessageNum, 0, -1) & "</font>"
        ElseIf FormatNumber(MessageCount / MaxMessageNum * 100, 0, -1) < 80 Then
            strSpace = strSpace & "<font color='blue'>" & FormatPercent(MessageCount / MaxMessageNum, 0, -1) & "</font>"
        Else
            strSpace = strSpace & "<font color='red'>" & FormatPercent(MessageCount / MaxMessageNum, 0, -1) & "</font>"
        End If
    End If
    GetSpace = strSpace
End Function

Function GetPath()
    Dim strPath
    strPath = "����Ϣ����"
    If Action = "Manage" Then
        strPath = strPath & "&nbsp;&gt;&gt;&nbsp;" & BoxName & "&nbsp;&gt;&gt;&nbsp;"
        If Keyword = "" Then
            If ManageType = "Inbox" And Action = "Manage" And Passed = "False" Then
                strPath = strPath & "δ�Ķ��Ķ���Ϣ"
            ElseIf ManageType = "Inbox" And Action = "Manage" And Passed = "True" Then
                strPath = strPath & "���Ķ��Ķ���Ϣ"
            Else
                strPath = strPath & "���ж���Ϣ"
            End If
        Else
            Select Case strField
                Case "Title"
                    strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �Ķ���Ϣ"
                Case "Content"
                    strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �Ķ���Ϣ"
                Case Else
                    strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �Ķ���Ϣ"
            End Select
        End If
    Else
        strPath = strPath & "&nbsp;&gt;&gt;&nbsp;" & ActionName
    End If
    GetPath = strPath
End Function

Function GetFriendListOption()
    Dim FriendListOption, arraytemp, strTemp, i
    strTemp = ""
    Set FriendListOption = Conn.Execute("select top 20 FriendName from PE_Friend where UserName='" & UserName & "' and GroupID<>0 order by AddTime desc")
    If Not FriendListOption.EOF Then
        arraytemp = FriendListOption.GetRows(-1)
        FriendListOption.Close
    End If
    Set FriendListOption = Nothing
    If IsArray(arraytemp) Then
        For i = 0 To UBound(arraytemp, 2)
            strTemp = strTemp & "<option value='" & arraytemp(0, i) & "'>" & arraytemp(0, i) & ""
        Next
    End If
    GetFriendListOption = strTemp
End Function
Function GetSearchForm()
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    strForm = strForm & "<tr><td height='28' align='center'>"
    strForm = strForm & " <select name='ManageType'>"
    strForm = strForm & "<option value='Inbox' "
    If ManageType = "Inbox" Then strForm = strForm & "selected"
    strForm = strForm & ">�ռ���</option>"
    strForm = strForm & "<option value='Outbox' "
    If ManageType = "Outbox" Then strForm = strForm & "selected"
    strForm = strForm & ">�ݸ���</option>"
    strForm = strForm & "<option value='IsSend' "
    If ManageType = "IsSend" Then strForm = strForm & "selected"
    strForm = strForm & ">�ѷ���</option>"
    strForm = strForm & "<option value='Recycle' "
    If ManageType = "Recycle" Then strForm = strForm & "selected"
    strForm = strForm & ">�ϼ���</option>"
    strForm = strForm & "</select>"
    strForm = strForm & " <select name='Field' size='1'>"
    strForm = strForm & "<option value='Title' selected>����Ϣ����</option>"
    strForm = strForm & "<option value='Content'>����Ϣ����</option>"
    strForm = strForm & "</select>"
    strForm = strForm & " <input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    strForm = strForm & "<input type='submit' name='Submit'  value='����'>"
    strForm = strForm & "</td></tr></form></table>"
    GetSearchForm = strForm
End Function

Function CheckBlackFriend(inceputName)
    Dim strFriend, strBlack
    CheckBlackFriend = False
    Set strFriend = Conn.Execute("select FriendName from PE_Friend where (UserName='" & UserName & "' or UserName='" & inceputName & "') and GroupID=0")
    If Not strFriend.EOF Then
        strBlack = strFriend.GetString(, , ",", "", "")
        If InStr(strBlack, inceputName) Or InStr(strBlack, UserName) Then CheckBlackFriend = True
    End If
    strFriend.Close
    Set strFriend = Nothing
End Function
%>
