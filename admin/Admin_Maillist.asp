<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "MailList"   '����Ȩ��

Response.Write "<html><head><title>�ʼ��б����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� �� �� �� �� ��", 10047)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Maillist.asp'>�����ʼ��б�</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_Maillist.asp?Action=Export'>�����ʼ��б�</a>&nbsp;|&nbsp;"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Action = Trim(Request("Action"))
Select Case Action
Case "Send"
    Call SendMaillist
Case "Export"
    Call ExportMail
Case "DoExport"
    Call DoExportMail
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim reSend
    Dim UserType, UserName
    UserType = PE_CLng(Trim(Request("UserType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function SelectUser(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=UserList&DefaultValue='+document.myform.inceptUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        document.myform.inceptUser.value=arr;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.subject.value==''){" & vbCrLf
    Response.Write "     alert('�ʼ����ⲻ��Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.subject.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "     alert('�ʼ����ݲ���Ϊ�գ�');" & vbCrLf
    Response.Write "     editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.SendperPage.value==''){" & vbCrLf
    Response.Write "     alert('������������Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.SendperPage.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_Maillist.asp'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b> �� �� �� ��</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>�ռ���ѡ��</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0'"
    If UserType = 0 Then Response.Write " checked"
    Response.Write "> ���л�Ա</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1'"
    If UserType = 1 Then Response.Write " checked"
    Response.Write "> ָ����Ա��</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2'"
    If UserType = 2 Then Response.Write " checked"
    Response.Write "> ָ���û���</td><td><input type='text' name='inceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=��<a href='#' onclick=""SelectUser();""><font color='green'>��Ա�б�</font></a>��</font>"
    Response.Write "����û���������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='3'"
    If UserType = 3 Then Response.Write " checked"
    Response.Write "> ָ����ԱEmail</td><td><input type='text' name='InceptEmail' size='40'>"
    Response.Write "���Email������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='15%' align='right'>�ʼ����⣺</td>"
    Response.Write "    <td width='85%'>"
    Response.Write "      <input type=text name=subject size=64>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right'>�ʼ����ݣ�</td>"
    Response.Write "    <td>"
    Response.Write "      <textarea name='Content' id='Content' style='display:none'></textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=-2&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='480' height='280' ></iframe>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='15%' align='right'>�����ˣ�</td>"
    Response.Write "    <td width='85%'>"
    Response.Write "      <input type='text' name='sendername' size='64' value='" & SiteName & "'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='15%' align='right'>������Email��</td>"
    Response.Write "    <td width='85%'>"
    Response.Write "      <input type='text' name='senderemail' size='64' value='" & WebmasterEmail & "'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='15%' align='right'>ÿ�η���������</td>"
    Response.Write "    <td width='85%'>"
    Response.Write "      <input type='text' name='SendperPage' size='5' value='100'>���ʼ�"
    Response.Write "    </td>"
    Response.Write "  </tr>"
        
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right'>�ʼ����ȼ���</td>"
    Response.Write "    <td>"
    Response.Write "      <input type='radio' name='Priority' value='1'>"
    Response.Write "      ��"
    Response.Write "      <input type='radio' name='Priority' value='3' checked>"
    Response.Write "      ��ͨ"
    Response.Write "      <input type='radio' name='Priority' value='5'>"
    Response.Write "      ��"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan=2 align=center>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='Send'>"
    Response.Write "      <input name='SendCount' type='hidden' id='SendCount' value='1'>"
    Response.Write "      <input name='Submit' type='submit' id='Submit' value=' �� �� ' "
    Response.Write "      >&nbsp;"
    Response.Write "      <input  name='Reset' type='reset' id='Reset' value=' �� �� '>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</form>"
    Response.Write "</table>"
End Sub

Sub SendMaillist()
    Dim sql, rs
    Dim totalsend, SendperPage, sendMsg, sendCount, endCount
    Dim Sendername, Senderemail, Subject, Content, Priority, InceptType, GroupID, inceptUser, InceptEmail, i, j, k
    i = 0
    j = 0
    k = 0
    sendMsg = ""
    sendCount = PE_CLng(Request("SendCount"))
    If sendCount < 1 Then sendCount = 1
    Sendername = Trim(Request("sendername"))
    Senderemail = Trim(Request("senderemail"))
    Subject = Trim(Request("Subject"))
    '�����ʼ����ݵı༭��
    Content = Trim(Request("Content"))
    Priority = Trim(Request("Priority"))
    SendperPage = PE_CLng(Request("SendperPage"))
    If Sendername = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����˲���Ϊ�գ�</li>"
    End If
    If Senderemail = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Email����Ϊ�գ�</li>"
    End If
    If Subject = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ʼ����ⲻ��Ϊ�գ�</li>"
    End If
    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ʼ����ݲ���Ϊ�գ�</li>"
    End If
    If Priority = "" Then
        Priority = 3
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    InceptType = CLng(Request("inceptType"))
    sql = "select UserName,Email from PE_User "
    If InceptType = 0 Then
        sql = sql & " where 1=1"
    ElseIf InceptType = 1 Then
        GroupID = Trim(Request("GroupID"))
        If IsValidID(GroupID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����Ա�飡</li>"
            Exit Sub
        End If
        If InStr(GroupID, ",") > 0 Then
            sql = sql & " where GroupID in (" & GroupID & ")"
        Else
            sql = sql & " where GroupID=" & GroupID
        End If
    ElseIf InceptType = 2 Then
        inceptUser = Replace(ReplaceBadChar(Request("InceptUser")), ",", "','")
        If inceptUser = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ�������˵��û�����</li>"
            Exit Sub
        End If
        sql = sql & " where UserName in ('" & inceptUser & "')"
    ElseIf InceptType = 3 Then
        InceptEmail = Replace(ReplaceBadChar(Request("InceptEmail")), ",", "','")
        If InceptEmail = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ�������˵����䣡</li>"
            Exit Sub
        End If
        sql = sql & " where Email in ('" & InceptEmail & "')"
    End If
    
    Dim PE_Mail
    Set PE_Mail = New SendMail
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ʱû�л�Աע�ᣡ</li>"
    Else
        sendMsg = sendMsg & "<li>���ڷ����У���ȴ�</li>"
        totalsend = rs.RecordCount
        endCount = sendCount + SendperPage - 1
        'Response.write "start:" &sendCount &"<br>"
        If endCount >= totalsend Then
            endCount = totalsend
        End If
        'Response.write "end:" &endCount &"<br>"
        'Response.write Content
        If Not rs.EOF Then
            If sendCount > 1 And sendCount <= endCount Then
                rs.Move sendCount - 1
            End If
        End If
        Do While Not rs.EOF
            If IsValidEmail(rs("Email")) = True Then
                ErrMsg = PE_Mail.Send(rs("Email"), rs("UserName"), Subject, Content, Sendername, Senderemail, Priority)
                If ErrMsg = "" Then
                    i = i + 1
                    sendMsg = sendMsg & "<li>�ɹ��� " & rs("UserName") & " �����ʼ���</li>"
                Else
                    j = j + 1
                    sendMsg = sendMsg & "<li><font color='red'>�� " & rs("UserName") & " �����ʼ�ʧ�ܣ�ʧ��ԭ��" & ErrMsg & "</font></li>"
                End If
            Else
                k = k + 1
            End If
            sendCount = sendCount + 1
            If sendCount > endCount Then Exit Do
            rs.MoveNext
        Loop
        sendMsg = sendMsg & "<li>�ɹ������ʼ���" & i & "��</li>"
        If j > 0 Then sendMsg = sendMsg & "<li>�����ʼ�ʧ�ܣ�" & j & "��<li>"
        If k > 0 Then sendMsg = sendMsg & "<li>δ�����ʼ���" & j & "�⣨�ʼ���ַ����<li>"
        If sendCount > totalsend Then
            Response.Write sendMsg
        Else
            If sendCount <= totalsend Then
                endCount = sendCount + SendperPage - 1
                If endCount >= totalsend Then
                    endCount = totalsend
                End If
                Response.Write "<div align='left'><form name=""sendmail"" method='post' action=""Admin_maillist.asp?Action=Send"">" & vbCrLf
                Response.Write "��" & totalsend & "���ʼ�,���͵�" & sendCount & "������" & endCount & "���ʼ�" & vbCrLf
                Response.write "<input type='hidden' name='sendername' value='"&Sendername&"'>" &vbCrLf
                Response.write "<input type='hidden' name='senderemail' value='"&Senderemail&"'>" &vbCrLf
                Response.write "<input type='hidden' name='Subject' value='"&Subject&"'>" &vbCrLf
                Response.write "<input type='hidden' name='Content' value='"&Content&"'>" &vbCrLf
                Response.write "<input type='hidden' name='Priority' value='"&Priority&"'>" &vbCrLf
                Response.write "<input type='hidden' name='SendperPage' value='"&SendperPage&"'>" &vbCrLf
                Response.write "<input type='hidden' name='inceptType' value='"&inceptType&"'>" &vbCrLf
                Response.write "<input type='hidden' name='GroupID' value='"&GroupID&"'>" &vbCrLf
                Response.write "<input type='hidden' name='InceptUser' value='"&InceptUser&"'>" &vbCrLf
                Response.write "<input type='hidden' name='InceptEmail' value='"&InceptEmail&"'>" &vbCrLf
                Response.write "<input type='hidden' name='SendCount' value='"&sendCount&"'>" &vbCrLf
                Response.Write "<input type='submit' name='submit' value='��������'>" & vbCrLf
                Response.Write "</form></div>"
            End If
            Response.Write sendMsg
        End If
    End If
    rs.Close
    Set rs = Nothing
    Set PE_Mail = Nothing
End Sub
Sub ShowJS_SendMail()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "     alert('�ʼ����ݲ���Ϊ�գ�');" & vbCrLf
    Response.Write "     editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub ExportMail()
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "<form method='post' action='Admin_Maillist.asp?Action=DoExport'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b> �ʼ��б��������������ݿ�</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='24%' height='80' align='right'>�����ʼ��б����ݿ⣺</td>"
    Response.Write "    <td width='76%' height='80'>"
    Response.Write "      <input name='ExportType' type='hidden' id='ExportType' value='1'>"
    Response.Write "      &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp;"
    Response.Write "      <select name='GroupID' id='GroupID'>" & GetUserGroup_Option & "</select>"
    Response.Write "      &nbsp;<font color=blue>��</font>&nbsp;"
    Response.Write "      <input name='ExportFileName' type='text' id='ExportFileName' value='maillist.mdb' size='30' maxlength='200'>"
    Response.Write "      <input type='submit' name='Submit' value='��ʼ'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</form>"
    Response.Write "</table>"
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "<form method='post' action='Admin_Maillist.asp?Action=DoExport'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>�ʼ��б������������ı�</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='24%' height='80' align='right'>�����ʼ��б��ı���</td>"
    Response.Write "    <td width='76%' height='80'>"
    Response.Write "      <input name='ExportType' type='hidden' id='ExportType' value='2'>"
    Response.Write "      &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp;"
    Response.Write "      <select name='GroupID' id='GroupID'>" & GetUserGroup_Option & "</select>"
    Response.Write "      </select>"
    Response.Write "      &nbsp;<font color=blue>��</font>&nbsp;"
    Response.Write "      <input name='ExportFileName' type='text' id='ExportFileName' value='maillist.txt' size='30' maxlength='200'>"
    Response.Write "      <input type='submit' name='Submit2' value='��ʼ' "
    If ObjInstalled_FSO = False Then Response.Write " disabled"
    Response.Write ">"
    If ObjInstalled_FSO = False Then
        Response.Write "      <font color=red>��ķ�������֧�� FSO! ����ʹ�ô˹��ܡ�</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</form>"
    Response.Write "</table>"
End Sub

Sub DoExportMail()
    Dim sql, rs
    Dim ExportType, GroupID, ExportFileName, strResult, i
    ExportType = PE_CLng(Trim(Request("ExportType")))
    GroupID = PE_CLng(Trim(Request("GroupID")))
    ExportFileName = Trim(Request("ExportFileName"))  
    If ExportFileName = "" Then
        FoundErr = True
        If ExportType = 1 Then
            ErrMsg = ErrMsg & "<li>������Ҫ���������ݿ��ļ�����</li>"
        Else
            ErrMsg = ErrMsg & "<li>������Ҫ�������ı��ļ�����</li>"
        End If
    Else
        ExportFileName = Replace(Replace(ExportFileName, "'", ""), Chr(34), "")
    End If
    
    Set rs = Server.CreateObject("adodb.recordset")
    If GroupID = 0 Then
        sql = "select Email from PE_User where Email like '%@%'"
    Else
        sql = "select Email from PE_User where Email like '%@%' and GroupID=" & GroupID & ""
    End If
    rs.Open sql, Conn, 1, 1

    i = 0
    Select Case ExportType
    Case 1
        Dim tconn, tconnstr
        Set tconn = Server.CreateObject("ADODB.Connection")
        tconnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(ExportFileName)
        tconn.Open tconnstr
        Do While Not rs.EOF
            tconn.Execute ("insert into [user] (useremail) values ('" & rs(0) & "')")
            rs.MoveNext
            i = i + 1
        Loop
        tconn.Close
        Set tconn = Nothing
        strResult = "�����ɹ��������� " & i & " ����ԱEmail��ַ�����ݿ� " & ExportFileName & "��<a href=" & ExportFileName & ">������ｫ���ݿ����ػر���</a>"
    Case 2
        Dim filepath, writefile
    
        Application.Lock
        filepath = Server.MapPath("" & ExportFileName & "")
        Set writefile = fso.CreateTextFile(filepath, True)
        Do While Not rs.EOF
            writefile.WriteLine rs(0)
            rs.MoveNext
            i = i + 1
        Loop
        writefile.Close
        Application.UnLock
        strResult = "�����ɹ��������� " & i & " ����ԱEmail��ַ��" & ExportFileName & "�ļ���<a href=" & ExportFileName & ">������ｫ�ļ����ػر���</a>"
    End Select
    rs.Close
    Set rs = Nothing

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' align=center><b>�ʼ��б���������������Ϣ</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='100' align='center'>" & strResult & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub



Function GetUserGroup_Option()
    Dim strGroup, rsGroup
    strGroup = "<option value='0'>ȫ����Ա</option>"
    Set rsGroup = Conn.Execute("select GroupID,GroupName from PE_UserGroup order by GroupType asc,GroupID asc")
    Do While Not rsGroup.EOF
        strGroup = strGroup & "<option value='" & rsGroup(0) & "'>" & rsGroup(1) & "</option>"
        rsGroup.MoveNext
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
    GetUserGroup_Option = strGroup
End Function
%>