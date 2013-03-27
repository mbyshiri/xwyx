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
Const PurviewLevel_Others = "UserGroup"   '����Ȩ��

Dim rsUserGroup, GroupSetting

GroupID = PE_CLng(Trim(Request("GroupID")))

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>��Ա�����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� Ա �� �� ��", 10042)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'>��������</td>" & vbCrLf
Response.Write "    <td height='30'><a href='?'>��Ա�������ҳ</a>&nbsp;|&nbsp;<a href='?Action=Add'>������Ա��</a> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveGroup
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
    Dim strSql, i
    strSql = "SELECT GroupID,GroupName,GroupIntro,GroupType,GroupSetting,arrClass_Browse,arrClass_View,arrClass_Input FROM PE_UserGroup ORDER by GroupType asc,GroupID asc"
    Set rsUserGroup = Server.CreateObject("adodb.recordset")
    rsUserGroup.Open strSql, Conn, 1, 1
    If rsUserGroup.BOF And rsUserGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������ݿ���û���ҵ��κλ�Ա�����ϣ��������ݿ��Ѿ��𻵣����Ĭ�����ݿ��е���PE_UserGroup��</li>"
        rsUserGroup.Close
        Set rsUserGroup = Nothing
        Exit Sub
    End If
    
    totalPut = rsUserGroup.recordcount
    CurrentPage = Trim(Request("page"))
    If CurrentPage = "" Then
        CurrentPage = 1
    Else
        CurrentPage = PE_CLng(CurrentPage)
    End If
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 20
    
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
            rsUserGroup.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' height='22' class='title'>" & vbCrLf
    Response.Write "    <td width='35'>ID</td>" & vbCrLf
    Response.Write "    <td width='120'>��Ա����</td>" & vbCrLf
    Response.Write "    <td>��Ա����</td>" & vbCrLf
    Response.Write "    <td width='120'>��Ա������</td>" & vbCrLf
    Response.Write "    <td width='60'>��Ա����</td>" & vbCrLf
    Response.Write "    <td width='150'>�� ��</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    Dim UserGroupNum
    UserGroupNum = 0
    Do While Not rsUserGroup.EOF
        Response.Write "     <tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'"">" & vbCrLf
        Response.Write "    <td width='35'>" & rsUserGroup("GroupID") & "</td>" & vbCrLf
        Response.Write "    <td width='120'>" & rsUserGroup("GroupName") & "</td>" & vbCrLf
        Response.Write "    <td align='left'>" & rsUserGroup("GroupIntro") & "</td>" & vbCrLf
        Response.Write "    <td width='120'>" & GetGroupType(rsUserGroup("GroupType")) & "</td>" & vbCrLf
        Response.Write "    <td width='60'>" & GetGroupNum(rsUserGroup("GroupID")) & "</td>" & vbCrLf
        Response.Write "    <td width='150'><a href='Admin_UserGroup.asp?Action=Modify&GroupID=" & rsUserGroup("GroupID") & "'>�޸�</a>"
        If rsUserGroup("GroupType") > 2  and rsUserGroup("GroupType") <> 5 Then
            Response.Write " | <a href='?Action=Del&GroupID=" & rsUserGroup("GroupID") & "' onclick=""return confirm('ȷʵҪɾ���˻�Ա����');"">ɾ��</a>" & vbCrLf
        Else
            Response.Write " | <font color='#CCCCCC'>ɾ��</font>"
        End If
        Response.Write " | <a href='Admin_User.asp?SearchType=11&GroupID=" & rsUserGroup("GroupID") & "'>�г���Ա</a></td>"
        Response.Write "    </tr>" & vbCrLf
        rsUserGroup.MoveNext
        UserGroupNum = UserGroupNum + 1
        If UserGroupNum >= MaxPerPage Then Exit Do
    Loop
    rsUserGroup.Close
    Set rsUserGroup = Nothing

    Response.Write "</table>" & vbCrLf
    Response.Write ShowPage("Admin_UserGroup.asp", totalPut, MaxPerPage, CurrentPage, True, True, "����Ա��", True)
End Sub

Function GetGroupType(GroupType)
    Select Case GroupType
    Case 0
        GetGroupType = "�ȴ��ʼ���֤��Ա"
    Case 1
        GetGroupType = "�ȴ�����Ա��˻�Ա"
    Case 2
        GetGroupType = "Ĭ�ϻ�Ա��"
    Case 3
        GetGroupType = "ע���Ա"
    Case 4
        GetGroupType = "�� �� ��"
    Case 5
        GetGroupType = "����Ͷ��"        
    Case Else
        GetGroupType = "δ֪��Ա��"
    End Select
End Function

Sub ShowJS_Check()
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function GetClassPurview(){" & vbCrLf
    Dim rsChannel, ChannelDir
    If PE_Clng(GroupID)<>-1 Then
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType<>4 And ModuleType<>5 And ModuleType<>7 And ModuleType<>8 And Disabled=" & PE_False & " ORDER BY OrderID")
    Else
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType = 1 And Disabled=" & PE_False & " ORDER BY OrderID")
    End If    
    
    Do While Not rsChannel.EOF
        ChannelDir = rsChannel(0)
        Response.Write "if(document.form1." & ChannelDir & "purview[1].checked==true){" & vbCrLf
        Response.Write "  document.form1.arrClass_Browse_" & ChannelDir & ".value='';" & vbCrLf
        Response.Write "  document.form1.arrClass_View_" & ChannelDir & ".value='';" & vbCrLf
        Response.Write "  document.form1.arrClass_Input_" & ChannelDir & ".value='';" & vbCrLf
        Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Browse.length;i++){" & vbCrLf
        Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Browse[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_Browse[i].checked==true){" & vbCrLf
        Response.Write "             if (document.form1.arrClass_Browse_" & ChannelDir & ".value=='')" & vbCrLf
        Response.Write "                 document.form1.arrClass_Browse_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Browse[i].value;" & vbCrLf
        Response.Write "             else" & vbCrLf
        Response.Write "                 document.form1.arrClass_Browse_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Browse[i].value;" & vbCrLf
        Response.Write "         }" & vbCrLf
        Response.Write "     }" & vbCrLf
        Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_View.length;i++){" & vbCrLf
        Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_View[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_View[i].checked==true){" & vbCrLf
        Response.Write "             if (document.form1.arrClass_View_" & ChannelDir & ".value=='')" & vbCrLf
        Response.Write "                 document.form1.arrClass_View_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
        Response.Write "             else" & vbCrLf
        Response.Write "                 document.form1.arrClass_View_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
        Response.Write "         }" & vbCrLf
        Response.Write "     }" & vbCrLf
        Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Input.length;i++){" & vbCrLf
        Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Input[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_Input[i].checked==true){" & vbCrLf
        Response.Write "             if (document.form1.arrClass_Input_" & ChannelDir & ".value=='')" & vbCrLf
        Response.Write "                 document.form1.arrClass_Input_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
        Response.Write "             else" & vbCrLf
        Response.Write "                 document.form1.arrClass_Input_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
        Response.Write "         }" & vbCrLf
        Response.Write "     }" & vbCrLf
        Response.Write "  }" & vbCrLf
        rsChannel.MoveNext
    Loop
    Set rsChannel = Nothing
    Response.Write "}" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "  if(document.form1.GroupName.value==''){" & vbCrLf
    Response.Write "      alert('��Ա�����Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "   document.form1.GroupName.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    GetClassPurview();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub
Sub Add()
    Call ShowJS_Check
    Response.Write "<form method='post' action='Admin_UserGroup.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='3'><div align='center'>�� �� �� Ա ��</div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>��Ա�����ƣ�</td>" & vbCrLf
    Response.Write "      <td><input name='GroupName' type='text' id='GroupName' size='20' maxlength='20'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>��Ա��˵����</td>" & vbCrLf
    Response.Write "      <td><input name='GroupIntro' type='text' id='GroupIntro' size='50' maxlength='200'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>�� �� �ͣ�</td>" & vbCrLf
    Response.Write "      <td><select name='GroupType' id='GroupType'>" & vbCrLf
    Response.Write "                            <option value='3'>ע���Ա</option>" & vbCrLf
    Response.Write "                            <option value='4'>�� �� ��</option>" & vbCrLf
    Response.Write "                        </select><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>����Ȩ�ޣ�</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting1' type='checkbox' value='1'>�ڷ�����Ϣ��Ҫ��˵�Ƶ���������Ա������Ϣ����Ҫ���<br>" & vbCrLf
    Response.Write "<input name='GroupSetting2' type='checkbox' value='1'>�����޸ĺ�ɾ������˵ģ��Լ��ģ���Ϣ<br>" & vbCrLf
    Response.Write "<input name='GroupSetting21' type='checkbox' value='1'>������Ϣʱ�������ñ���ǰ׺<br>" & vbCrLf
    Response.Write "<input name='GroupSetting22' type='checkbox' value='1'>������Ϣʱ���������Ƿ���ʾ��������<br>" & vbCrLf
    Response.Write "<input name='GroupSetting23' type='checkbox' value='1'>������Ϣʱ��������ת������<br>" & vbCrLf
    Response.Write "<input name='GroupSetting24' type='checkbox' value='1'>������ϢʱHTML�༭��Ϊ�߼�ģʽ��Ĭ��Ϊ���ģʽ��<br>" & vbCrLf
    Response.Write "ÿ����෢��<input name='GroupSetting3' type='text' value='10' size='6' maxlength='6' style='text-align: center;'>����Ϣ����������������Ϊ<b>0</b>����<br>"
    Response.Write "������Ϣʱ��ȡ����Ϊ��Ŀ���õ�<input name='GroupSetting4' type='text' value='1' size='5' maxlength='5' style='text-align: center;'>��<br>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>����Ȩ�ޣ�</td>" & vbCrLf
    Response.Write "         <td><input name='GroupSetting5' type='checkbox' value='1'>�ڽ�ֹ�������۵���Ŀ��Ȼ�ɷ�������<br><input name='GroupSetting6' type='checkbox' value='1'>��������Ҫ��˵���Ŀ�﷢�����۲���Ҫ���</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>����ϢȨ�ޣ�</td>" & vbCrLf
    Response.Write "      <td> ÿ������ͬʱ��<input name='GroupSetting7' type='text' value='1' size='4' maxlength='4' style='text-align: center;'>�˷��Ͷ���Ϣ�����Ϊ0���������Ͷ���Ϣ��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>�ղؼ�Ȩ�ޣ�</td>" & vbCrLf
    Response.Write "      <td>��Ա�ղؼ���������¼<input name='GroupSetting8' type='text' value='500' size='5' maxlength='5' style='text-align: center;'>����Ϣ�����Ϊ0����û���ղ�Ȩ�ޣ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>�ϴ��ļ�Ȩ�ޣ�</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting9' type='checkbox' value='1' checked>�����ڿ����ϴ���Ƶ���ϴ��ļ�<br>��������ϴ�<input name='GroupSetting10' type='text' value='1024' size='5' style='text-align: center;'>K���ļ�����������ֵ����ĳһƵ��������ʱ����Ƶ������Ϊ׼����</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>�̳�Ȩ�ޣ�</td>" & vbCrLf
    Response.Write "      <td>����ʱ�������ܵ��ۿ��ʣ�<input name='GroupSetting11' type='text' value='100' size='5' maxlength='5' style='text-align: center;'> %<br>"
    Response.Write "        <input name='GroupSetting12' type='checkbox' value='1' checked>�Ƿ���������������Żݣ���ָ����Ա�۵���Ʒ��Ч��<br>"
    Response.Write "        ����͸֧������ȣ�<input name='GroupSetting13' type='text' value='0' size='6' maxsize='6' style='text-align: center;'> Ԫ�����<br>"
    Response.Write "        <input name='GroupSetting30' type='checkbox' value='1'>�Ƿ����������Ʒ<br>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>�Ʒѷ�ʽ��</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting14' type='radio' value='0' checked>ֻ�ж�" & PointName & "����" & PointName & "ʱ����ʹ��Ч���Ѿ����ڣ��Կ��Բ鿴�շ����ݣ�" & PointName & "����󣬼�ʹ��Ч��û�е��ڣ�Ҳ���ܲ鿴�շ����ݡ�<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting14' type='radio' value='1'>ֻ�ж���Ч�ڣ�ֻҪ����Ч���ڣ�" & PointName & "������Կ��Բ鿴�շ����ݣ����ں󣬼�ʹ��Ա��" & PointName & "Ҳ���ܲ鿴�շ����ݡ�<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting14' type='radio' value='2'>ͬʱ�ж�" & PointName & "����Ч�ڣ�" & PointName & "�������Ч�ڵ��ں󣬾Ͳ��ɲ鿴�շ����ݡ�<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting14' type='radio' value='3'>ͬʱ�ж�" & PointName & "����Ч�ڣ�" & PointName & "���겢����Ч�ڵ��ں󣬲Ų��ܲ鿴�շ����ݡ�" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>��" & PointName & "��ʽ��</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting15' type='radio' value='0' checked>��Ч���ڣ��鿴�շ����ݲ���" & PointName & "����Ҳ������¼��<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting15' type='radio' value='1'>��Ч���ڣ��鿴�շ����ݲ���" & PointName & "����������¼��<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting15' type='radio' value='2'>��Ч���ڣ��鿴�շ�����Ҳ��" & PointName & "����<br>" & vbCrLf
    Response.Write "��Ч���ڣ��ܹ����Կ�<input name='GroupSetting16' type='text' value='0' size='10' maxlength='10' style='text-align: center;'> ���շ���Ϣ�����Ϊ0�������ƣ�<br>" & vbCrLf
    Response.Write "��Ч���ڣ�ÿ�������Կ�<input name='GroupSetting17' type='text' value='100' size='10' maxlength='10' style='text-align: center;'> ���շ���Ϣ�����Ϊ0�������ƣ�" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>������ֵ��</td>" & vbCrLf
    Response.Write "         <td><input name='GroupSetting18' type='checkbox' value='1' checked>���������һ�" & PointName & "<br><input name='GroupSetting19' type='checkbox' value='1' checked>���������һ���Ч��<br><input name='GroupSetting20' type='checkbox' value='1' checked>����" & PointName & "���͸�����</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>�ۺϿռ䣺</td>" & vbCrLf
    Response.Write "         <td><input name='GroupSetting25' type='checkbox'>���þۺϿռ�<br>" & vbCrLf
    Response.Write "         <input name='GroupSetting26' type='checkbox'>����ʱ�������Ա���<br>" & vbCrLf
    Response.Write " �ۺϿռ�����Ϊ:<input name='GroupSetting27' type='text' value='10' size='4' maxlength='10' style='text-align: center;'>M<br>" & vbCrLf
    Response.Write "         <input name='GroupSetting28' type='checkbox'>�û�������������Ƥ��" & vbCrLf
    Response.Write "    </td></tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td colspan='3'>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='10' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td colspan='2' align='center'>Ƶ �� Ȩ �� �� ϸ �� ��</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
 
    Dim rsChannel
    Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType<>4 And ModuleType<>5 And ModuleType<>7 And ModuleType<>8 AND Disabled=" & PE_False & " ORDER BY OrderID")
    Do While Not rsChannel.EOF
        Response.Write "          <tr valign='top'>" & vbCrLf
        Response.Write "           <td><fieldset>" & vbCrLf
        Response.Write "   <legend>�˻�Ա���ڡ�<font color='red'>" & rsChannel("ChannelName") & "</font>��Ƶ����Ȩ�ޣ�</legend>" & vbCrLf
        Response.Write "    <table width='100%' cellspacing='1'>" & vbCrLf
        Response.Write "        <tr class='tdbg'>" & vbCrLf
        Response.Write "                <td width='50%'><input type='radio' name='" & rsChannel("ChannelDir") & "purview' checked onClick=table" & rsChannel("ChannelID") & ".style.display='none'>���κ�Ȩ��(������Ŀ����)"
        Response.Write "&nbsp;&nbsp;<input type='radio' name='" & rsChannel("ChannelDir") & "purview' onClick=table" & rsChannel("ChannelID") & ".style.display='block'>���û�Ա���ڸ�Ƶ����Ȩ��</td>" & vbCrLf
        Response.Write "             <td></td>" & vbCrLf
        Response.Write "        <tr class='tdbg' id='table" & rsChannel("ChannelID") & "' style='display:none'>" & vbCrLf
        Response.Write "         <td width='50%'>" & vbCrLf
        Response.Write "         <iframe id='frm" & rsChannel("ChannelDir") & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Group&ChannelID=" & rsChannel("ChannelID") & "'></iframe>" & vbCrLf
        Response.Write "         <input name='arrClass_Browse_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Browse_" & rsChannel("ChannelDir") & "' value=''>" & vbCrLf
        Response.Write "         <input name='arrClass_View_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_View_" & rsChannel("ChannelDir") & "' value=''>" & vbCrLf
        Response.Write "         <input name='arrClass_Input_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Input_" & rsChannel("ChannelDir") & "' value=''></td>" & vbCrLf
        Response.Write "         <td width='50%'><font color='#0000FF'>ע��</font><br>1����ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ��Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ������Ȩ�ޣ�����������Ŀ��ָ�������Ȩ�ޡ�<br>2����ɫ����ѡ�е���Ŀ��˵������ĿΪ������Ŀ����Ա���ڴ���Ŀӵ������Ͳ鿴Ȩ�ޡ�<br><br><font color='red'>Ȩ�޺��壺</font><br>�������ָ�����������Ŀ����Ϣ�б�<br>�鿴����ָ���Բ鿴����Ŀ�е���Ϣ������<br>��������ָ�����ڴ���Ŀ�з�����Ϣ</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "   </table>" & vbCrLf
        Response.Write "   </fieldset></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "            <tr>" & vbCrLf
    Response.Write "                <td align='center'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='Action' value='SaveAdd'>" & vbCrLf
    Response.Write "                    <input type='submit' value='��ӻ�Ա��'>" & vbCrLf
    Response.Write "                    <input type='button' name='cancel' value=' ȡ �� ' onClick=""JavaScript:window.location.href='Admin_UserGroup.asp'"">" & vbCrLf
    Response.Write "                </td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
 
End Sub

Sub Modify()
    GroupID = PE_CLng(Trim(Request.QueryString("GroupID")))
    If GroupID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��GroupID</li>"
        Exit Sub
    End If
        
    Set rsUserGroup = Conn.Execute("SELECT GroupID,GroupName,GroupIntro,GroupType,GroupSetting,arrClass_Browse,arrClass_View,arrClass_Input FROM PE_UserGroup WHERE GroupID=" & GroupID & "")
    If rsUserGroup.BOF And rsUserGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ļ�Ա��</li>"
        rsUserGroup.Close
        Set rsUserGroup = Nothing
        Exit Sub
    End If
    
    '��ֹ��Ա�ֹ��޸����ݿ⵼����������ȱ�ٵĴ���
    GroupSetting = rsUserGroup("GroupSetting") & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
    '���
    GroupSetting = Split(GroupSetting, ",")
    Dim i
    For i = 0 To UBound(GroupSetting)
        If GroupSetting(i) = "" Then GroupSetting(i) = 0
    Next
    
    Call ShowJS_Check
    
    Response.Write "<form method='post' action='Admin_UserGroup.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='3'><div align='center'>�� �� �� Ա �� �� ��</div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>��Ա�����ƣ�</td>" & vbCrLf
    Response.Write "      <td><input name='GroupName' type='text' id='GroupName' value='" & rsUserGroup("GroupName") & "' size='20' maxlength='20'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>��Ա��˵����</td>" & vbCrLf
    Response.Write "      <td><input name='GroupIntro' type='text' id='GroupIntro' value='" & rsUserGroup("GroupIntro") & "' size='50' maxlength='200'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>�� �� �ͣ�</td>" & vbCrLf
    Response.Write "      <td><select name='GroupType' id='GroupType'"
    If rsUserGroup("GroupType") < 3 or rsUserGroup("GroupType") = 5  Then Response.Write "disabled"
    Response.Write ">" & vbCrLf
    If rsUserGroup("GroupType") = 0 Then
        Response.Write "        <option value='0' selected>�ȴ��ʼ���֤</option>" & vbCrLf
    End If
    If rsUserGroup("GroupType") = 1 Then
        Response.Write "        <option value='1' selected>�ȴ�����Ա����</option>" & vbCrLf
    End If
    If rsUserGroup("GroupType") = 2 Then
        Response.Write "        <option value='2' selected>Ĭ�ϻ�Ա��</option>" & vbCrLf
    End If
    If rsUserGroup("GroupType") = 5 Then
        Response.Write "        <option value='5' selected>����Ͷ��</option>" & vbCrLf
    End If    
    Response.Write "            <option value='3'"
    If rsUserGroup("GroupType") = 3 Then
        Response.Write " selected"
    End If
    Response.Write ">ע���Ա</option>" & vbCrLf
    Response.Write "            <option value='4'"
    If rsUserGroup("GroupType") = 4 Then
        Response.Write " selected"
    End If
    Response.Write ">�� �� ��</option></select>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    If rsUserGroup("GroupID")<>-1 Then            
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>����Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting1' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(1)), 1) & ">�ڷ�����Ϣ��Ҫ��˵�Ƶ���������Ա������Ϣ����Ҫ���<br>" & vbCrLf
        Response.Write "<input name='GroupSetting2' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(2)), 1) & ">�����޸ĺ�ɾ������˵ģ��Լ��ģ���Ϣ<br>" & vbCrLf
        Response.Write "<input name='GroupSetting21' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(21)), 1) & ">������Ϣʱ�������ñ���ǰ׺<br>" & vbCrLf
        Response.Write "<input name='GroupSetting22' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(22)), 1) & ">������Ϣʱ���������Ƿ���ʾ��������<br>" & vbCrLf
        Response.Write "<input name='GroupSetting23' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(23)), 1) & ">������Ϣʱ��������ת������<br>" & vbCrLf
        Response.Write "<input name='GroupSetting24' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(24)), 1) & ">������ϢʱHTML�༭��Ϊ�߼�ģʽ��Ĭ��Ϊ���ģʽ��<br>" & vbCrLf
        Response.Write "ÿ����෢��<input name='GroupSetting3' type='text' value='" & GroupSetting(3) & "' size='6' maxlength='6' style='text-align: center;'>����Ϣ����������������Ϊ<b>0</b>����<br>"
        Response.Write "������Ϣʱ��ȡ����Ϊ��Ŀ���õ�<input name='GroupSetting4' type='text' value='" & GroupSetting(4) & "' size='5' maxlength='5' style='text-align: center;'>��<br>"
        
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "    <td width='15%' class='tdbg5' align='right'>����Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "         <td><input name='GroupSetting5' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(5)), 1) & ">�ڽ�ֹ�������۵���Ŀ����Ȼ�ɷ�������<br><input name='GroupSetting6' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(6)), 1) & ">��������Ҫ��˵���Ŀ�﷢�����۲���Ҫ���</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>����ϢȨ�ޣ�</td>" & vbCrLf
        Response.Write "      <td>ÿ������ͬʱ��<input name='GroupSetting7' type='text' value='" & GroupSetting(7) & "' size='4' maxlength='4' style='text-align: center;'>�˷��Ͷ���Ϣ�����Ϊ0���������Ͷ���Ϣ��</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>�ղؼ�Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "      <td>��Ա�ղؼ���������¼<input name='GroupSetting8' type='text' value='" & GroupSetting(8) & "' size='5' maxlength='5' style='text-align: center;'>����Ϣ�����Ϊ0����û���ղ�Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>�ϴ��ļ�Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting9' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(9)), 1) & ">�����ڿ����ϴ���Ƶ���ϴ��ļ�<br>��������ϴ�<input name='GroupSetting10' type='text' value='" & GroupSetting(10) & "' size='5' style='text-align: center;'>K���ļ�����������ֵ����ĳһƵ��������ʱ����Ƶ������Ϊ׼����</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>�̳�Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "      <td>����ʱ�������ܵ��ۿ��ʣ�<input name='GroupSetting11' type='text' value='" & GroupSetting(11) & "' size='5' maxlength='5' style='text-align: center;'>%<br>"
        Response.Write "<input name='GroupSetting12' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(12)), 1) & ">�Ƿ���������������Żݣ���ָ����Ա�۵���Ʒ��Ч��<br> ����͸֧������ȣ�<input name='GroupSetting13' type='text' value='" & GroupSetting(13) & "' size='6' maxsize='6' style='text-align: center;'>Ԫ�����" & vbCrLf
        Response.Write "        <br><input name='GroupSetting30' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(30)), 1) & ">�Ƿ����������Ʒ<br>"
        Response.Write "    </td></tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>�Ʒѷ�ʽ��</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting14' type='radio' " & RadioValue(PE_CLng(GroupSetting(14)), 0) & ">ֻ�ж�" & PointName & "����" & PointName & "ʱ����ʹ��Ч���Ѿ����ڣ��Կ��Բ鿴�շ����ݣ�" & PointName & "����󣬼�ʹ��Ч��û�е��ڣ�Ҳ���ܲ鿴�շ����ݡ�<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting14' " & RadioValue(PE_CLng(GroupSetting(14)), 1) & ">ֻ�ж���Ч�ڣ�ֻҪ����Ч���ڣ�" & PointName & "������Կ��Բ鿴�շ����ݣ����ں󣬼�ʹ��Ա��" & PointName & "Ҳ���ܲ鿴�շ����ݡ�<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting14' " & RadioValue(PE_CLng(GroupSetting(14)), 2) & ">ͬʱ�ж�" & PointName & "����Ч�ڣ�" & PointName & "�������Ч�ڵ��ں󣬾Ͳ��ɲ鿴�շ����ݡ�<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting14' " & RadioValue(PE_CLng(GroupSetting(14)), 3) & ">ͬʱ�ж�" & PointName & "����Ч�ڣ�" & PointName & "���겢����Ч�ڵ��ں󣬲Ų��ܲ鿴�շ����ݡ�" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>��" & PointName & "��ʽ��</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting15' type='radio' " & RadioValue(PE_CLng(GroupSetting(15)), 0) & ">��Ч���ڣ��鿴�շ����ݲ���" & PointName & "��Ҳ������¼��<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting15' " & RadioValue(PE_CLng(GroupSetting(15)), 1) & ">��Ч���ڣ��鿴�շ����ݲ���" & PointName & "��������¼��<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting15' " & RadioValue(PE_CLng(GroupSetting(15)), 2) & ">��Ч���ڣ��鿴�շ�����Ҳ��" & PointName & "��<br>" & vbCrLf
        Response.Write "��Ч���ڣ��ܹ����Կ�<input name='GroupSetting16' type='text' value='" & GroupSetting(16) & "' size='10' maxlength='10' style='text-align: center;'> ���շ���Ϣ�����Ϊ0�������ƣ�<br>" & vbCrLf
        Response.Write "��Ч���ڣ�ÿ�������Կ�<input name='GroupSetting17' type='text' value='" & GroupSetting(17) & "' size='10' maxlength='10' style='text-align: center;'> ���շ���Ϣ�����Ϊ0�������ƣ�" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "    <td width='15%' class='tdbg5' align='right'>������ֵ��</td>" & vbCrLf
        Response.Write "         <td><input name='GroupSetting18' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(18)), 1) & ">���������һ�" & PointName & "<br><input name='GroupSetting19' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(19)), 1) & ">���������һ���Ч��<br><input name='GroupSetting20' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(20)), 1) & ">����" & PointName & "���͸�����</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "    <td width='15%' class='tdbg5' align='right'>�ۺϿռ䣺</td>" & vbCrLf
        Response.Write "         <td><input name='GroupSetting25' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(25)), 1) & ">���þۺϿռ�<br>" & vbCrLf
        Response.Write "         <input name='GroupSetting26' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(26)), 1) & ">�ۺϿռ��������<br>" & vbCrLf
        Response.Write " �ۺϿռ��������Ϊ:<input name='GroupSetting27' type='text' value='" & GroupSetting(27) & "' size='4' maxlength='10' style='text-align: center;'>M<br>" & vbCrLf
        Response.Write "         <input name='GroupSetting28' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(28)), 1) & ">�û�������������Ƥ��" & vbCrLf
        Response.Write "    </td></tr>" & vbCrLf
    Else    	
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>����Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting1' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(1)), 1) & ">�ڷ�����Ϣ��Ҫ��˵�Ƶ���������Ա������Ϣ����Ҫ���<br>" & vbCrLf
        Response.Write "<input name='GroupSetting21' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(21)), 1) & ">������Ϣʱ�������ñ���ǰ׺<br>" & vbCrLf
        Response.Write "<input name='GroupSetting22' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(22)), 1) & ">������Ϣʱ���������Ƿ���ʾ��������<br>" & vbCrLf
        Response.Write "<input name='GroupSetting23' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(23)), 1) & ">������Ϣʱ��������ת������<br>" & vbCrLf
        Response.Write "<input name='GroupSetting24' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(24)), 1) & ">������ϢʱHTML�༭��Ϊ�߼�ģʽ��Ĭ��Ϊ���ģʽ��<br>" & vbCrLf
        Response.Write "ÿ����෢��<input name='GroupSetting3' type='text' value='" & GroupSetting(3) & "' size='6' maxlength='6' style='text-align: center;'>����Ϣ����������������Ϊ<b>0</b>����<br>"
        
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>�ϴ��ļ�Ȩ�ޣ�</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting9' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(9)), 1) & ">�����ڿ����ϴ���Ƶ���ϴ��ļ�<br>��������ϴ�<input name='GroupSetting10' type='text' value='" & GroupSetting(10) & "' size='5' style='text-align: center;'>K���ļ�����������ֵ����ĳһƵ��������ʱ����Ƶ������Ϊ׼����</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf    
    End If        
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td colspan='3'>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='10' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td colspan='2' align='center'>Ƶ �� Ȩ �� �� ϸ �� ��</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
     
    Dim rsChannel, arrPurviews, IsNoPurview
    arrPurviews = rsUserGroup("arrClass_Browse") & "," & rsUserGroup("arrClass_View") & "," & rsUserGroup("arrClass_Input")
    If rsUserGroup("GroupID")<>-1 Then            
        Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType<>4 And ModuleType<>5 and ModuleType<>7 and ModuleType<>8 AND Disabled=" & PE_False & " ORDER BY OrderID")
    Else
        Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType=1 AND Disabled=" & PE_False & " ORDER BY OrderID")    
    End If
    Do While Not rsChannel.EOF
        IsNoPurview = FoundInArr(arrPurviews, rsChannel("ChannelDir") & "none", ",")
        Response.Write "          <tr valign='top'>" & vbCrLf
        Response.Write "           <td><fieldset>" & vbCrLf
        Response.Write "   <legend>�˻�Ա���ڡ�<font color='red'>" & rsChannel("ChannelName") & "</font>��Ƶ����Ȩ�ޣ�</legend>" & vbCrLf
        Response.Write "    <table width='100%' cellspacing='1'>" & vbCrLf
        Response.Write "        <tr class='tdbg'>" & vbCrLf
        Response.Write "                <td width='50%'><input type='radio' name='" & rsChannel("ChannelDir") & "purview' onClick=""table" & rsChannel("ChannelID") & ".style.display='none'"""
        If IsNoPurview = True Then Response.Write "checked"
        Response.Write ">���κ�Ȩ��(������Ŀ����)"
        Response.Write "&nbsp;&nbsp;<input type='radio' name='" & rsChannel("ChannelDir") & "purview' onClick=""table" & rsChannel("ChannelID") & ".style.display='block'"""
        If IsNoPurview = False Then Response.Write "checked"
        Response.Write ">���û�Ա�ڸ�Ƶ����Ȩ��</td>" & vbCrLf
        Response.Write "             <td></td>" & vbCrLf
        Response.Write "        <tr class='tdbg' id='table" & rsChannel("ChannelID") & "' style='display:"
        If IsNoPurview = True Then
            Response.Write "none"
        Else
            Response.Write "block"
        End If
        Response.Write "'>" & vbCrLf
        Response.Write "         <td width='50%'>" & vbCrLf
        Response.Write "         <iframe id='frm" & rsChannel("ChannelDir") & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Group&Action=Modify&ChannelID=" & rsChannel("ChannelID") & "&GroupID=" & GroupID & "'></iframe>" & vbCrLf
        Response.Write "         <input name='arrClass_Browse_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Browse_" & rsChannel("ChannelDir") & "' value='" & rsChannel("ChannelDir") & "none'>" & vbCrLf
        Response.Write "         <input name='arrClass_View_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_View_" & rsChannel("ChannelDir") & "' value='" & rsChannel("ChannelDir") & "none'>" & vbCrLf
        Response.Write "         <input name='arrClass_Input_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Input_" & rsChannel("ChannelDir") & "' value='" & rsChannel("ChannelDir") & "none'></td>" & vbCrLf
        Response.Write "         <td width='50%'><font color='#0000FF'>ע��</font><br>1����ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ��Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ������Ȩ�ޣ�����������Ŀ��ָ�������Ȩ�ޡ�<br>2����ɫ����ѡ�е���Ŀ��˵������ĿΪ������Ŀ����Ա���ڴ���Ŀӵ������Ͳ鿴Ȩ�ޡ�<br><br><font color='red'>Ȩ�޺��壺</font><br>�������ָ�����������Ŀ����Ϣ�б�<br>�鿴����ָ���Բ鿴����Ŀ�е���Ϣ������<br>��������ָ�����ڴ���Ŀ�з�����Ϣ</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "   </table>" & vbCrLf
        Response.Write "   </fieldset></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "            <tr>" & vbCrLf
    Response.Write "                <td align='center'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='GroupID' value='" & rsUserGroup("GroupID") & "'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='Action' value='SaveModify'>" & vbCrLf
    Response.Write "                    <input type='submit' value='�����޸Ľ��'>" & vbCrLf
    Response.Write "                    <input type='button' name='cancel' value=' ȡ �� ' onClick=""JavaScript:window.location.href='Admin_UserGroup.asp'"">" & vbCrLf
    Response.Write "                </td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Del()
    GroupID = PE_CLng(Trim(Request("GroupID")))
    If GroupID = 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ɾ��ϵͳĬ�ϵĻ�Ա��</li>"
        Exit Sub
    End If
    Conn.Execute ("update PE_User set GroupID=1 where GroupID=" & GroupID & "")
    Conn.Execute ("delete from PE_UserGroup where GroupID=" & GroupID & " AND GroupType>=2")
    Call main
End Sub

Sub SaveGroup()
    Dim GroupType, strValue, GroupIntro, i
    Dim rsUserGroup, rsChannel, GroupPurview, GroupPurviewChannel
    Dim arrClass_Browse, arrClass_View, arrClass_Input
    GroupID = Trim(Request.Form("GroupID"))
    GroupName = Trim(Request.Form("GroupName"))
    GroupIntro = Trim(Request.Form("GroupIntro"))
    GroupType = Trim(Request.Form("GroupType"))
    FoundErr = False
    If GroupName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Ա�����Ʋ���Ϊ��</li>"
        Exit Sub
    Else
        GroupName = ReplaceBadChar(GroupName)
    End If
    If GroupType = "" Then
        GroupType = 0
    Else
        GroupType = CLng(GroupType)
    End If
    GroupSetting = ""
    For i = 0 To 30
        strValue = Trim(Request.Form("GroupSetting" & i & ""))
        If strValue = "" Or (Not IsNumeric(strValue)) Then
            strValue = "0"
        End If
        If GroupSetting = "" Then
            GroupSetting = strValue
        Else
            GroupSetting = GroupSetting & "," & strValue
        End If
    Next

    arrClass_Browse = ""
    arrClass_View = ""
    arrClass_Input = ""
    
    Dim tBrowse, tView, tInput, ChannelDir
    If PE_Clng(GroupID)<>-1 then
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType<>4 And ModuleType<>5 And Disabled=" & PE_False & " ORDER BY OrderID")
    Else
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType=1 And Disabled=" & PE_False & " ORDER BY OrderID")
    End If    
    Do While Not rsChannel.EOF
        ChannelDir = rsChannel(0)
        tBrowse = ReplaceBadChar(Trim(Request.Form("arrClass_Browse_" & ChannelDir)))
        tView = ReplaceBadChar(Trim(Request.Form("arrClass_View_" & ChannelDir)))
        tInput = ReplaceBadChar(Trim(Request.Form("arrClass_Input_" & ChannelDir)))
        If tBrowse = "" And tView = "" And tInput = "" Then
            If arrClass_Browse = "" Then
                arrClass_Browse = ChannelDir & "none"
            Else
                arrClass_Browse = arrClass_Browse & "," & ChannelDir & "none"
            End If
            If arrClass_View = "" Then
                arrClass_View = ChannelDir & "none"
            Else
                arrClass_View = arrClass_View & "," & ChannelDir & "none"
            End If
            If arrClass_View = "" Then
                arrClass_View = ChannelDir & "none"
            Else
                arrClass_View = arrClass_View & "," & ChannelDir & "none"
            End If
       Else
            If tBrowse <> "" Then
                If arrClass_Browse = "" Then
                    arrClass_Browse = tBrowse
                Else
                    arrClass_Browse = arrClass_Browse & "," & tBrowse
                End If
            End If
            If tView <> "" Then
                If arrClass_View = "" Then
                    arrClass_View = tView
                Else
                    arrClass_View = arrClass_View & "," & tView
                End If
            End If
            If tInput <> "" Then
                If arrClass_Input = "" Then
                    arrClass_Input = tInput
                Else
                    arrClass_Input = arrClass_Input & "," & tInput
                End If
            End If
        End If
        rsChannel.MoveNext
    Loop
    Set rsChannel = Nothing

    Set rsUserGroup = Server.CreateObject("Adodb.Recordset")
    If Action = "SaveAdd" Then
        rsUserGroup.Open "SELECT * FROM PE_UserGroup WHERE GroupName='" & GroupName & "'", Conn, 1, 3
        If Not (rsUserGroup.BOF And rsUserGroup.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ݿ�������ͬ����Ա�飡</li>"
        Else
            rsUserGroup.addnew
            rsUserGroup("GroupID") = PE_CLng(Conn.Execute("select max(GroupID) from PE_UserGroup")(0)) + 1
        End If
    Else
        rsUserGroup.Open "SELECT * FROM PE_UserGroup WHERE GroupID=" & GroupID, Conn, 1, 3
        If rsUserGroup.BOF And rsUserGroup.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ݿ�û�з��ִ˻�Ա�飡</li>"
        End If
    End If
    If FoundErr = True Then
        rsUserGroup.Close
        Set rsUserGroup = Nothing
        Exit Sub
    End If

    rsUserGroup("GroupName") = GroupName
    rsUserGroup("GroupIntro") = GroupIntro
    If GroupType > 0 Then
        rsUserGroup("GroupType") = GroupType
    End If
    rsUserGroup("GroupSetting") = GroupSetting
    rsUserGroup("arrClass_Browse") = arrClass_Browse
    rsUserGroup("arrClass_View") = arrClass_View
    rsUserGroup("arrClass_Input") = arrClass_Input
    rsUserGroup.Update
    rsUserGroup.Close
    Set rsUserGroup = Nothing
    Call main

End Sub

Function GetGroupNum(iGroupID)
    If Not IsNumeric(iGroupID) Then Exit Function
    Dim rsUserGroup
    Set rsUserGroup = Conn.Execute("SELECT Count(UserID) FROM PE_User WHERE GroupID=" & iGroupID & "")
    If IsNull(rsUserGroup(0)) Then
        GetGroupNum = 0
    Else
        GetGroupNum = rsUserGroup(0)
    End If
    Set rsUserGroup = Nothing
End Function
%>

