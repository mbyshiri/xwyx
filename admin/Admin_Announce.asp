<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_ContentEx.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Announce"   '����Ȩ��

Dim ItemName, ID


strFileName = "Admin_Announce.asp?Action=" & Action
ItemName = "����"
ID = Trim(Request("ID"))
If IsValidID(ID) = False Then
    ID = ""
End If

Response.Write "<html><head><title>�������</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� վ �� �� �� ��", 10023)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Announce.asp'>���������ҳ</a>&nbsp;|&nbsp;<a href='Admin_Announce.asp?Action=Add'>����¹���</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveAnnounce
Case "SetNew", "CancelNew", "SetShowType", "Move", "Del"
    Call SetProperty
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "����������ʧ�ܣ�ʧ��ԭ��" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim rs, sql
    Call ShowJS_Main(ItemName)
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetChannelList(ChannelID) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>"
    Call ShowManagePath(ChannelID)
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Announce.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "  <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td width='30' height='22' align='center'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='30' height='22' align='center'><strong>ID</strong></td>"
    Response.Write "    <td height='22' align='center'><strong>�� ��</strong></td>"
    Response.Write "    <td width='60' height='22' align='center'><strong>���¹���</strong></td>"
    Response.Write "    <td width='60' height='22' align='center'><strong>��ʾ��ʽ</strong></td>"
    Response.Write "    <td width='60' height='22' align='center'><strong>������</strong></td>"
    Response.Write "    <td width='120' height='22' align='center'><strong>����ʱ��</strong></td>"
    Response.Write "    <td width='60' height='22' align='center'><strong>��Ч��</strong></td>"
    Response.Write "    <td width='150' height='22' align='center'><strong>����</strong></td>"
    Response.Write "  </tr>"

    sql = "select * from PE_Announce"
    If ChannelID >= -1 Then
        sql = sql & " where ChannelID=" & ChannelID
    End If
    sql = sql & " order by IsSelected desc,ID desc"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>û���κι��棡<br><br></td></tr>"
    Else
        Do While Not rs.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "      <td width='30' align='center'><input name='ID' type='checkbox' onclick='unselectall()' value='" & rs("ID") & "'></td>"
            Response.Write "      <td width='30' align='center'>" & rs("ID") & "</td>"
            Response.Write "      <td><a href='Admin_Announce.asp?Action=Modify&ID=" & rs("ID") & "'"
            Response.Write "      title='�������ݣ�" & Left(nohtml(rs("content")), 200) & "'>" & rs("Title") & "</a></td>"
            Response.Write "      <td width='60' align='center'>"
            If rs("IsSelected") = True Then
                Response.Write "<font color=green>��</font>"
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='60' align='center'>"
            If rs("ShowType") = 0 Then
                Response.Write "ȫ��"
            ElseIf rs("ShowType") = 1 Then
                Response.Write "����"
            ElseIf rs("ShowType") = 2 Then
                Response.Write "����"
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='60' align='center'>" & rs("Author") & "</td>"
            Response.Write "      <td width='120' align='center'>" & rs("DateAndTime") & "</td>"
            Response.Write "      <td width='60' align='center'>"
            If rs("OutTime") > 0 Then
                Response.Write rs("OutTime") & "��"
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='150' align='center'>"
            Response.Write "      <a href='Admin_Announce.asp?Action=Modify&ID=" & rs("ID") & "'>�޸�</a>&nbsp;"
            Response.Write "      <a href='Admin_Announce.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('ȷ��Ҫɾ���˹�����');"">ɾ��</a>&nbsp;"
            If rs("IsSelected") = False Then
                Response.Write "      <a href='Admin_Announce.asp?Action=SetNew&ID=" & rs("ID") & "'>��Ϊ����</a>"
            Else
                Response.Write "      <a href='Admin_Announce.asp?Action=CancelNew&ID=" & rs("ID") & "'>ȡ������</a>"
            End If
            Response.Write "      </td>"
            Response.Write "    </tr>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='130' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�����еĹ���</td><td>"
    Response.Write "<input type='submit' value='ɾ��ѡ���Ĺ���' name='submit' onClick=""document.myform.Action.value='Del'"">&nbsp;&nbsp;"
    Response.Write "<input type='submit' value='����ѡ��������ʾ��ʽ' name='submit' onClick=""document.myform.Action.value='SetShowType'"">"
    Response.Write "<select name='ShowType'>"
    Response.Write "  <option value='0'>ȫ��</option>"
    Response.Write "  <option value='1'>����</option>"
    Response.Write "  <option value='2'>����</option>"
    Response.Write "</select>&nbsp;&nbsp;"
    Response.Write "<input type='submit' value='��ѡ���Ĺ����ƶ��� ->' name='submit' onClick=""document.myform.Action.value='Move'"">"
    Response.Write "<select name='ChannelID' id='ChannelID'>" & GetChannel_Option(0) & "</select>"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    Response.Write "<br><b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;ֻ�н�������Ϊ���¹����Ż���ǰ̨��ʾ"
    Response.Write "<br><br>"
End Sub

Sub ShowJS_AddModify()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "     alert('������ⲻ��Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.Title.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "     alert('�������ݲ���Ϊ�գ�');" & vbCrLf
    Response.Write "     editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Add()
    Call ShowJS_AddModify
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Announce.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='20%' align='right'>����Ƶ����</td>"
    Response.Write "      <td width='80%'>"
    Response.Write "        <select name='ChannelID' id='ChannelID'>" & GetChannel_Option(0) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>���⣺</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Title' size='66' id='Title' value=''>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>���ݣ�</td>"
    Response.Write "      <td>"
    Response.Write "       <textarea name='Content' id='Content' style='display:none' ></textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=-1&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='480' height='280' ></iframe>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>�����ˣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='Author' type='text' id='Author' value='" & AdminName & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>����ʱ�䣺</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='DateAndTime' type='text' id='DateAndTime' value='" & Now() & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>��Ч�ڣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='OutTime' type='text' id='OutTime' value='1' size='10' maxlength='20'> �죨Ϊ0ʱ����ʾ��Զ��Ч��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>��ʾ���ͣ�</td>"
    Response.Write "      <td>" & GetShowType_Option(1) & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>&nbsp;</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='IsSelected' type='checkbox' id='IsSelected' value='yes' checked>"
    Response.Write "        ��Ϊ���¹���</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input type='submit' name='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim sql, rs
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹ���ID��</li>"
        Exit Sub
    Else
        ID = PE_CLng(ID)
    End If
    sql = "select * from PE_Announce where ID=" & ID
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ĺ��棡</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Call ShowJS_AddModify
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Announce.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    Response.Write "      <td height='22' colspan='2'><strong>�� �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>����Ƶ����</td>"
    Response.Write "      <td>"
    Response.Write "        <select name='ChannelID' id='ChannelID'>" & GetChannel_Option(rs("ChannelID")) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='20%' align='right'>���⣺</td>"
    Response.Write "      <td width='80%'>"
    Response.Write "        <input type='text' name='Title' size='66' id='Title' value='" & rs("Title") & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>���ݣ�</td>"
    Response.Write "      <td>"
    Response.Write "       <textarea name='Content' id='Content' style='display:none' >" & Server.HTMLEncode(rs("Content")) & "</textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='480' height='280' ></iframe>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>�����ˣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='Author' type='text' id='Author' value='" & rs("Author") & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>����ʱ�䣺</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='DateAndTime' type='text' id='DateAndTime' value='" & rs("DateAndTime") & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>��Ч�ڣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='OutTime' type='text' id='OutTime' value='" & rs("OutTime") & "' size='10' maxlength='20'> �죨Ϊ0ʱ����ʾ��Զ��Ч��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>��ʾ���ͣ�</td>"
    Response.Write "      <td>" & GetShowType_Option(rs("ShowType")) & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>&nbsp;</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='IsSelected' type='checkbox' id='IsSelected' value='yes' "
    If rs("IsSelected") = True Then Response.Write " checked"
    Response.Write "        >��Ϊ���¹���</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input type='submit' name='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rs.Close
    Set rs = Nothing
End Sub

Sub SaveAnnounce()
    Dim Title, Content, Author, DateAndTime, ShowType, IsSelected, OutTime
    Dim rs, sql
    Title = Trim(Request("Title"))
    Content = Trim(Request("Content"))
    Author = Trim(Request("Author"))
    DateAndTime = PE_CDate(Trim(Request("DateAndTime")))
    ShowType = PE_CLng(Request("ShowType"))
    IsSelected = Trim(Request("IsSelected"))
    OutTime = PE_CLng(Request("OutTime"))
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ⲻ��Ϊ�գ�</li>"
    End If
    If Len(Title) > 250 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������������ӦС��250����</li>"
    End If
    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ݲ���Ϊ�գ�</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    Title = PE_HTMLEncode(Title)
    Author = PE_HTMLEncode(Author)
    If ShowType = "" Then
        ShowType = 0
    Else
        ShowType = PE_CLng(ShowType)
    End If
    If IsSelected = "yes" Then
        IsSelected = True
    Else
        IsSelected = False
    End If
    Set rs = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        sql = "select top 1 * from PE_Announce"
        rs.Open sql, Conn, 1, 3
        rs.addnew
    ElseIf Action = "SaveModify" Then
        If ID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ������ID</li>"
            Exit Sub
        Else
            sql = "select * from PE_Announce where ID=" & PE_CLng(ID)
            rs.Open sql, Conn, 1, 3
            If rs.BOF And rs.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ĺ��棡</li>"
                rs.Close
                Set rs = Nothing
                Exit Sub
            End If
        End If
    End If

    rs("ChannelID") = ChannelID
    rs("Title") = Title
    rs("Content") = Content
    rs("Author") = Author
    rs("DateAndTime") = DateAndTime
    rs("ShowType") = ShowType
    rs("IsSelected") = IsSelected
    rs("OutTime") = OutTime
    rs.Update
    rs.Close
    Set rs = Nothing
    Call ClearSiteCache(0)
    Call WriteEntry(2, AdminName, "���湫��ɹ���" & Title)

    Call CloseConn
    Response.Redirect "admin_announce.asp?ChannelID=" & ChannelID
End Sub

Sub SetProperty()
    Dim sqlProperty, rsProperty
    Dim ShowType, MoveChannelID
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If InStr(ID, ",") > 0 Then
        sqlProperty = "select * from PE_Announce where ID in (" & ID & ")"
    Else
        sqlProperty = "select * from PE_Announce where ID=" & ID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetNew"
            rsProperty("IsSelected") = True
        Case "CancelNew"
            rsProperty("IsSelected") = False
        Case "SetShowType"
            ShowType = Trim(Request("ShowType"))
            If ShowType = "" Then
                ShowType = 0
            Else
                ShowType = PE_CLng(ShowType)
            End If
            rsProperty("ShowType") = ShowType
        Case "Move"
            MoveChannelID = PE_CLng(Trim(Request("ChannelID")))
            rsProperty("ChannelID") = MoveChannelID
        Case "Del"
            rsProperty.Delete
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing
    
    Call ClearSiteCache(0)
    Call WriteEntry(2, AdminName, "���ù������Գɹ���" & ID)
    Call CloseConn
    Response.Redirect ComeUrl
End Sub


Function GetShowType_Option(ShowType)
    Dim strShowType
    strShowType = "<input type='radio' name='ShowType' value='0'"
    If ShowType = 0 Then
        strShowType = strShowType & " checked"
    End If
    strShowType = strShowType & ">" & "ȫ��&nbsp;&nbsp;"
    strShowType = strShowType & "<input type='radio' name='ShowType' value='1'"
    If ShowType = 1 Then
        strShowType = strShowType & " checked"
    End If
    strShowType = strShowType & ">" & "����&nbsp;&nbsp;"
    strShowType = strShowType & "<input type='radio' name='ShowType' value='2'"
    If ShowType = 2 Then
        strShowType = strShowType & " checked"
    End If
    strShowType = strShowType & ">" & "����&nbsp;&nbsp;"
    GetShowType_Option = strShowType
End Function
%>
