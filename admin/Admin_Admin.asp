<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 1      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

strFileName = "Admin_Admin.asp"

Response.Write "<html><head><title>����Ա����</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� �� Ա �� ��", 10049)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td height='30'><a href='Admin_Admin.asp'>����Ա������ҳ</a>&nbsp;|&nbsp;<a href='Admin_Admin.asp?Action=Add'>��������Ա</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call AddAdmin
Case "SaveAdd"
    Call SaveAdd
Case "ModifyPwd"
    Call ModifyPwd
Case "ModifyPurview"
    Call ModifyPurview
Case "SaveModifyPwd"
    Call SaveModifyPwd
Case "SaveModifyPurview"
    Call SaveModifyPurview
Case "Del"
    Call DelAdmin
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(1, AdminName, "����Ա����ʧ�ܣ�ʧ��ԭ��" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsAdminList, sqlAdminList
    Dim iCount
    
    Set rsAdminList = Server.CreateObject("Adodb.RecordSet")
    sqlAdminList = "select * from PE_Admin order by id"
    rsAdminList.Open sqlAdminList, Conn, 1, 1
    If rsAdminList.BOF And rsAdminList.EOF Then
        rsAdminList.Close
        Set rsAdminList = Nothing
        Response.Write "û���κι���Ա��"
        Exit Sub
    End If
    
    totalPut = rsAdminList.RecordCount
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
            rsAdminList.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
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
    Response.Write "    if (e.Name != 'chkAll'&&e.disabled!=true)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "</script>" & vbCrLf
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Admin.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�еĹ���Ա��');"">"
    Response.Write "     <td>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title' height='22'>"
    Response.Write "    <td  width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td  width='30' height='22'><strong>���</strong></td>"
    Response.Write "    <td><strong>����Ա��</strong></td>"
    Response.Write "    <td><strong>ǰ̨��Ա��</strong></td>"
    Response.Write "    <td width='70'><strong>Ȩ ��</strong></td>"
    Response.Write "    <td width='55'><strong>���˵�¼</strong></td>"
    Response.Write "    <td width='95'><strong>����¼IP</strong></td>"
    Response.Write "    <td width='115'><strong>����¼ʱ��</strong></td>"
    Response.Write "    <td width='55'><strong>��¼����</strong></td>"
    Response.Write "    <td width='180'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"
    Do While Not rsAdminList.EOF
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td width='30'><input name='ID' type='checkbox' id='ID' value='" & rsAdminList("ID") & "'"
        If rsAdminList("UserName") = AdminName Then Response.Write " disabled"
        Response.Write " onclick='unselectall()'></td>"
        Response.Write "    <td width='30'>" & rsAdminList("ID") & "</td>"
        Response.Write "    <td>"
        If rsAdminList("AdminName") = AdminName Then
            Response.Write "<font color=red><b>" & rsAdminList("AdminName") & "</b></font>"
        Else
            Response.Write rsAdminList("AdminName")
        End If
        Response.Write "</td>"
        Response.Write "    <td>" & rsAdminList("UserName") & "</td>"
        Response.Write "<td width='70'>"
        Select Case rsAdminList("purview")
            Case 1
              Response.Write "<font color=blue>��������Ա</font>"
            Case 2
              Response.Write "��ͨ����Ա"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='55'>"
        If rsAdminList("EnableMultiLogin") = True Then
            Response.Write "<font color='green'>����</font>"
        Else
            Response.Write "<font color='red'>������</font>"
        End If
        Response.Write "</td>"
        Response.Write "<td width='95'>"
        If rsAdminList("LastLoginIP") <> "" Then
            Response.Write rsAdminList("LastLoginIP")
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='115'>"
        If rsAdminList("LastLoginTime") <> "" Then
            Response.Write rsAdminList("LastLoginTime")
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='55'>"
        If Trim(rsAdminList("LoginTimes")) <> "" Then
            Response.Write rsAdminList("LoginTimes")
        Else
            Response.Write "0"
        End If
        Response.Write "</td>"
        Response.Write "<td width='180'>"
        Response.Write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rsAdminList("ID") & "'>�޸����뼰����</a>&nbsp;&nbsp;"
        Response.Write "<a href='Admin_Admin.asp?Action=ModifyPurview&ID=" & rsAdminList("ID") & "'>�޸�Ȩ��</a>&nbsp;&nbsp;"
        Response.Write "</td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsAdminList.MoveNext
    Loop
    rsAdminList.Close
    Set rsAdminList = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ�����й���Ա</td>"
    Response.Write "    <td><input name='Action' type='hidden' id='Action' value='Del'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�еĹ���Ա'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������Ա", True)
End Sub

Sub ShowJS_Check()
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function SelectAll(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.AdminPurview_Others.length;i++){" & vbCrLf
    Response.Write "    var e = form.AdminPurview_Others[i];" & vbCrLf
    Response.Write "    if (e.disabled==false)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAdd(){" & vbCrLf
    Response.Write "  if(document.form1.username.value==''){" & vbCrLf
    Response.Write "      alert('�û�������Ϊ�գ�');" & vbCrLf
    Response.Write "   document.form1.username.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if(document.form1.Password.value==''){" & vbCrLf
    Response.Write "      alert('���벻��Ϊ�գ�');" & vbCrLf
    Response.Write "   document.form1.Password.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value)){" & vbCrLf
    Response.Write "      alert('��ʼ������ȷ�����벻ͬ��');" & vbCrLf
    Response.Write "   document.form1.PwdConfirm.select();" & vbCrLf
    Response.Write "   document.form1.PwdConfirm.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if (document.form1.Purview[1].checked==true){" & vbCrLf
    Response.Write " GetClassPurview();" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function CheckModifyPwd(){" & vbCrLf
    Response.Write "  if(document.form1.Password.value==''){" & vbCrLf
    Response.Write "      alert('���벻��Ϊ�գ�');" & vbCrLf
    Response.Write "   document.form1.Password.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value)){" & vbCrLf
    Response.Write "      alert('��ʼ������ȷ�����벻ͬ��');" & vbCrLf
    Response.Write "   document.form1.PwdConfirm.select();" & vbCrLf
    Response.Write "   document.form1.PwdConfirm.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function CheckModifyPurview(){" & vbCrLf
    Response.Write "  if (document.form1.Purview[1].checked==true){" & vbCrLf
    Response.Write " GetClassPurview();" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function GetClassPurview(){" & vbCrLf
    
    Dim ChannelDir
    Dim sqlChannel, rsChannel
    sqlChannel = "select ChannelDir,ModuleType from PE_Channel where ChannelType<=1 and ModuleType <> 8 "
    sqlChannel = sqlChannel & "And Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        ChannelDir = rsChannel(0)
        If rsChannel(1) = 4 Then
            Response.Write "  document.form1.arrKind_Modify_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrKind_Del_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrKind_Move_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrKind_Check_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrKind_Quintessence_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrKind_SetOnTop_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrKind_Reply_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  if(document.form1.AdminPurview_" & ChannelDir & "[2].checked==true){" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Modify.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Modify[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_Modify_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_Modify_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Modify[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_Modify_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Modify[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Del.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Del[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_Del_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_Del_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Del[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_Del_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Del[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Move.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Move[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_Move_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_Move_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Move[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_Move_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Move[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Check.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Check[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_Check_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_Check_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Check[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_Check_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Check[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Quintessence.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Quintessence[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_Quintessence_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_Quintessence_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Quintessence[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_Quintessence_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Quintessence[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_SetOnTop.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_SetOnTop[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_SetOnTop_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_SetOnTop_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_SetOnTop[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_SetOnTop_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_SetOnTop[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Reply.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Reply[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrKind_Reply_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrKind_Reply_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Reply[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrKind_Reply_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Reply[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
        ElseIf rsChannel(1) = 7 Then
            Response.Write "  document.form1.arrHouseClass_Input_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrHouseClass_Check_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrHouseClass_Manage_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  if(document.form1.AdminPurview_" & ChannelDir & "[1].checked==true){" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_View.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_View[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrHouseClass_View_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_View_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_View_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Input.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Input[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrHouseClass_Input_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_Input_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_Input_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Check.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Check[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrHouseClass_Check_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_Check_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Check[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_Check_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Check[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Manage.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Manage[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrHouseClass_Manage_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_Manage_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Manage[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrHouseClass_Manage_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Manage[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
        Else
            Response.Write "  document.form1.arrClass_Input_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrClass_Check_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  document.form1.arrClass_Manage_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "  if(document.form1.AdminPurview_" & ChannelDir & "[2].checked==true){" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_View.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_View[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrClass_View_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrClass_View_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrClass_View_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Input.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Input[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrClass_Input_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrClass_Input_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrClass_Input_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Check.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Check[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrClass_Check_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrClass_Check_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Check[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrClass_Check_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Check[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
            Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Manage.length;i++){" & vbCrLf
            Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Manage[i].checked==true){" & vbCrLf
            Response.Write "             if (document.form1.arrClass_Manage_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                 document.form1.arrClass_Manage_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Manage[i].value;" & vbCrLf
            Response.Write "             else" & vbCrLf
            Response.Write "                 document.form1.arrClass_Manage_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Manage[i].value;" & vbCrLf
            Response.Write "         }" & vbCrLf
            Response.Write "     }" & vbCrLf
        End If
        Response.Write " }" & vbCrLf
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
    
End Sub

Sub AddAdmin()
    Call ShowJS_Check
    Response.Write "<form method='post' action='Admin_Admin.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� Ա</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>����Ա����</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='AdminName' type='text'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>ǰ̨��Ա����</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='username' type='text'> <a href='Admin_User.asp?Action=AddUser'>����»�Ա</a></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>��ʼ���룺</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input type='password' name='Password' onkeyup='javascript:EvalPwdStrength(document.forms[0],this.value);' onmouseout='javascript:EvalPwdStrength(document.forms[0],this.value);' onblur='javascript:EvalPwdStrength(document.forms[0],this.value);'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>����ǿ�ȣ�</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'>"
    Call ShowPwdStrength
    Response.Write "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>ȷ�����룺</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input type='password' name='PwdConfirm'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%'>&nbsp;</td>"
    Response.Write "      <td width='88%'><input name='EnableMultiLogin' type='checkbox' value='Yes'>�������ͬʱʹ�ô��ʺŵ�¼</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>Ȩ�����ã� </strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='Purview' type='radio' value='1' onClick=""PurviewDetail.style.display='none'"">��������Ա��ӵ������Ȩ�ޡ�ĳЩȨ�ޣ������Ա������վ��Ϣ���á���վѡ�����õȹ���Ȩ�ޣ�ֻ�г�������Ա���С�"
    Response.Write "  <br><input type='radio' name='Purview' value='2' checked  onClick=""PurviewDetail.style.display=''"">��ͨ����Ա����Ҫ��ϸָ��ÿһ�����Ȩ��</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2'><table id='PurviewDetail' width='100%' border='0' cellspacing='10' cellpadding='0' style='display:'>"
    Response.Write "        <tr>"
    Response.Write "          <td colspan='2' align='center'><strong>�� �� Ա Ȩ �� �� ϸ �� ��</strong></td>"
    Response.Write "        </tr>"

    Dim sqlChannel, rsChannel
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"

    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        ChannelID = rsChannel("ChannelID")
        ChannelName = Trim(rsChannel("ChannelName"))
        ChannelShortName = Trim(rsChannel("ChannelShortName"))
        ChannelDir = rsChannel("ChannelDir")

        Response.Write "<tr valign='top'><td>"
        Response.Write "<fieldset><legend>�˹���Ա�ڡ�<font color='red'>" & ChannelName & "</font>��Ƶ����Ȩ�ޣ�</legend><table width='100%'><tr>"
        If rsChannel("ModuleType") = 4 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='';"">Ƶ������Ա��ӵ����Ŀ�ܱ��Ȩ�ޣ������Թ���IP���ú����ι��</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������Թ����������</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ����Ա��ֻӵ�в�������������Ȩ��</td>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""OtherPurview_" & ChannelDir & ".style.display='';"">Ƶ������Ա��ӵ�������˲���Ƹģ��Ĺ���Ȩ�ޣ����Է���ְλ��Ϣ�͹����˲�</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""OtherPurview_" & ChannelDir & ".style.display='none';"">ְλ��Ϣ����Ա��ӵ��ְλ��Ϣ�ķ������޸�Ȩ�ޣ������ܹ����˲ź�ɾ��ְλ�������Ϣ(����˲���Ϣ��һ��ɾ��)</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""OtherPurview_" & ChannelDir & ".style.display='none';"">�˲���Ϣ����Ա��ӵ�������˲���Ϣ�Ĺ���Ȩ�ް���ɾ�����޸�ӦƸ�ߵļ���</td>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""OtherPurview_" & ChannelDir & ".style.display='';table_" & ChannelDir & ".style.display='none'"">Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ��</td>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ��</td>"
        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""OtherPurview_" & ChannelDir & ".style.display='';"">Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ��</td>"
        Else
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='';"">Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ��</td>"
        End If
        Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='4' checked onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">�ڴ�Ƶ�������κι���Ȩ��</td></tr>"
        If rsChannel("ModuleType") = 4 Then
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='60%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td>"
            Response.Write "  <input name='arrKind_Modify_" & ChannelDir & "' type='hidden' id='arrKind_Modify_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Del_" & ChannelDir & "' type='hidden' id='arrKind_Del_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Move_" & ChannelDir & "' type='hidden' id='arrKind_Move_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Check_" & ChannelDir & "' type='hidden' id='arrKind_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Quintessence_" & ChannelDir & "' type='hidden' id='arrKind_Quintessence_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_SetOnTop_" & ChannelDir & "' type='hidden' id='arrKind_SetOnTop_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Reply_" & ChannelDir & "' type='hidden' id='arrKind_Reply_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            'Response.Write "<b>����Ȩ�ޣ�</b>"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'>������"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>ģ�����&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�"
            Response.Write "  <input name='arrHouseClass_View_" & ChannelDir & "' type='hidden' id='arrHouseClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Input_" & ChannelDir & "' type='hidden' id='arrHouseClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Check_" & ChannelDir & "' type='hidden' id='arrHouseClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Manage_" & ChannelDir & "' type='hidden' id='arrHouseClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='ClassConfig_" & ChannelDir & "'>������Ŀ���ù���&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Area_" & ChannelDir & "'>�����������&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>ģ�����&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>ģ�����&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'>�����˵�&nbsp;"
            Response.Write "</td></tr></table>"
            
            Response.Write "<Table>"
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></Table>"

            Response.Write "</fieldset></td></tr>"

        Else
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>ģ�����&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='JsFile_" & ChannelDir & "'>JS�ļ�����&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'>�����˵�&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Keyword_" & ChannelDir & "'>�ؼ��ֹ���&nbsp;"
            If rsChannel("ModuleType") = 5 Then
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Producer_" & ChannelDir & "'>���̹���&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Trademark_" & ChannelDir & "'>Ʒ�ƹ���&nbsp;"
            Else
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Author_" & ChannelDir & "'>���߹���&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Copyfrom_" & ChannelDir & "'>��Դ����&nbsp;"
            End If
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='XML_" & ChannelDir & "'>����XML&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Field_" & ChannelDir & "'>�Զ����ֶ�&nbsp;"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'>������"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    
    
    Response.Write "   <tr><td><fieldset><legend>�˹���Ա��������վ����Ȩ�ޣ�<input name='chkAll' type='checkbox' id='chkAll' value='Yes' onclick='SelectAll(this.form)'>ѡ������Ȩ��</legend>"
    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ModifyPwd' checked>�޸��Լ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Channel'>��վƵ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Collection'>�ɼ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Message'>����Ϣ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='MailList'>�ʼ��б����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD'>��վ������</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='FriendSite'>�������ӹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Announce'>��վ�������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Vote'>��վ�������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Counter'>��վͳ�ƹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Skin'>��վ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template0'>ͨ��ģ�����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Label'>�Զ����ǩ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowPage'>�Զ���ҳ�����</td>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Cache'>��վ�������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='KeyLink'>վ�����ӹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Rtext'>�ַ����˹���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='UserGroup'>��Ա�����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Card'>��ֵ������</td>"

    If FoundInArr(AllModules, "Classroom", ",") Then
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Equipment'>�ҳ��Ǽǹ���</td>"
    End If
    If FoundInArr(AllModules, "Sdms", ",") Then
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='InfoManage'>ѧ����Ϣ����</td>"
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ScoreManage'>ѧ���ɼ�����</td>"
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='TestManage'>���Թ���</td>"
    End If
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>��Ա����Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_View'>�鿴��Ա��Ϣ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_ModifyInfo'>�޸Ļ�Ա��Ϣ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_MofidyPurview'>�޸Ļ�ԱȨ��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Lock'>��ס/������Ա</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Del'>ɾ����Ա</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Update'>����Ϊ�ͻ�</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Money'>��Ա�ʽ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Point'>��Ա" & PointName & "����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Valid'>��Ա��Ч�ڹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ConsumeLog'>��Ա������ϸ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='RechargeLog'>��Ա��Ч����ϸ</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�̳��ճ���������Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_View'>�鿴����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Confirm'>ȷ�϶���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Modify'>�޸Ķ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Del'>ɾ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Payment'>�տ��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Invoice'>����Ʊ</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Deliver'>�������ͣ�ʵ�</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Download'>�������ͣ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SendCard'>�������ͣ��㿨��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_End'>���嶩��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Transfer'>��������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Print'>������ӡ</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Count'>����ͳ��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_OrderItem'>������ϸ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SaleCount'>����ͳ��/����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Payment'>����֧������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bankroll'>�ʽ���ϸ��ѯ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Deliver'>���˻���¼</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Transfer'>����������¼</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PresentProject'>������������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PaymentType'>���ʽ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='DeliverType'>�ͻ���ʽ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bank'>�����ʻ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Refund'>�˿��</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShoppingCart'>���ﳵ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddPayment'>�������֧��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AgentPayment'>���������֧��</td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"


    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�ͻ���ϵ����Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_View'>�鿴�ͻ���Ϣ</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Add'>��ӿͻ�</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyOwn'>�޸������Լ��Ŀͻ���Ϣ</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyAll'>�޸����пͻ���Ϣ</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Del'>ɾ���ͻ�</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_View'>�鿴�����¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Add'>��ӷ����¼</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyOwn'>�޸��Լ���ӵķ����¼</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyAll'>�޸����з����¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Del'>ɾ�������¼</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_View'>�鿴Ͷ�߼�¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Add'>���Ͷ�߼�¼</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyOwn'>�޸��Լ���ӵ�Ͷ�߼�¼</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyAll'>�޸�����Ͷ�߼�¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Del'>ɾ��Ͷ�߼�¼</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_View'>�鿴�طü�¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Add'>��ӻطü�¼</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyOwn'>�޸��Լ���ӵĻطü�¼</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyAll'>�޸����лطü�¼</td>"
    'Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Del'>ɾ��Ͷ�߼�¼</td>"
    Response.Write "          </tr>"

    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�ֻ����Ź���Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToMember'>���͸���Ա</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToContacter'>���͸���ϵ��</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToConsignee'>���͸������е��ջ���</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToOther'>���͸�������</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewMessageLog'>�鿴���ͽ��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SMS_MessageReceive'>�鿴���յ��Ķ���</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�ʾ�������Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewSurvey'>�鿴�ʾ�</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddSurvey'>�����ʾ�</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurvey'>�����ʾ��޸ġ�ɾ����</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowSurveyCountData'>�鿴������</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurveyTemplate'>�ʾ�ģ�����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ImportSurveyQuestion'>�ʾ���Ŀ����</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ExportSurveyQuestion'>�ʾ���Ŀ����</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewListQuestion'>�鿴�ʾ���Ŀ�б�</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"


    Response.Write "      </fieldset></td></tr>"
    Response.Write "  </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Admin.asp';"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
    Call ShowPurviewTips
End Sub

Sub ModifyPwd()
    Dim UserID
    Dim rsAdmin, sqlAdmin
    UserID = Trim(Request("ID"))
    If UserID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
        Exit Sub
    Else
        UserID = PE_CLng(UserID)
    End If
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˹���Ա��</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub

    End If
    Call ShowJS_Check
    Response.Write "<form method='post' action='Admin_Admin.asp' name='form1' onsubmit='return CheckModifyPwd();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� �� Ա �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>����Ա����</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'>" & rsAdmin("AdminName") & "<input name='ID' type='hidden' value='" & rsAdmin("ID") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>ǰ̨��Ա����</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'><input name='UserName' type='text' value='" & rsAdmin("UserName") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td width='40%' class='tdbg'><strong>�� �� �룺</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'><input type='password' name='Password' onkeyup='javascript:EvalPwdStrength(document.forms[0],this.value);' onmouseout='javascript:EvalPwdStrength(document.forms[0],this.value);' onblur='javascript:EvalPwdStrength(document.forms[0],this.value);'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>����ǿ�ȣ�</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'>"
    Call ShowPwdStrength
    Response.Write "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>ȷ�����룺</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'><input type='password' name='PwdConfirm'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='40%'>&nbsp;</td>"
    Response.Write "      <td width='60%'><input name='EnableMultiLogin' type='checkbox' value='Yes'"
    If rsAdmin("EnableMultiLogin") = True Then Response.Write " checked"
    Response.Write ">�������ͬʱʹ�ô��ʺŵ�¼</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModifyPwd'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "        <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Admin.asp'"" style='cursor:hand;'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"

    rsAdmin.Close
    Set rsAdmin = Nothing
End Sub

Sub ModifyPurview()
    Dim UserID
    Dim rsAdmin, sqlAdmin
    Dim PO
    UserID = Trim(Request("ID"))
    If UserID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
        Exit Sub
    Else
        UserID = PE_CLng(UserID)
    End If
    
    
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˹���Ա��</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub
    End If

    PO = rsAdmin("AdminPurview_Others")
    Call ShowJS_Check
    
    Response.Write "<form method='post' action='Admin_Admin.asp' name='form1' onsubmit='javascript:CheckModifyPurview();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� �� Ա Ȩ ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='12%' class='tdbg'><strong>����Ա����</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'>" & rsAdmin("AdminName") & "<input name='ID' type='hidden' value='" & rsAdmin("ID") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='12%' class='tdbg'><strong>ǰ̨��Ա����</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'>" & rsAdmin("UserName") & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%'>&nbsp;</td>"
    Response.Write "      <td width='88%'><input name='EnableMultiLogin' type='checkbox' value='Yes'"
    If rsAdmin("EnableMultiLogin") = True Then Response.Write " checked"
    Response.Write ">�������ͬʱʹ�ô��ʺŵ�¼</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='12%' class='tdbg'><strong>Ȩ�����ã�</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='Purview' type='radio' value='1' onClick=""PurviewDetail.style.display='none'"""
    If rsAdmin("Purview") = 1 Then Response.Write "checked"
    Response.Write ">��������Ա��ӵ������Ȩ�ޡ�ĳЩȨ�ޣ������Ա������վ��Ϣ���á���վѡ�����õȹ���Ȩ�ޣ�ֻ�г�������Ա����<br>"
    Response.Write "<input type='radio' name='Purview' value='2' onClick=""PurviewDetail.style.display=''"""
    If rsAdmin("Purview") = 2 Then Response.Write "checked"
    Response.Write ">��ͨ����Ա����Ҫ��ϸָ��ÿһ�����Ȩ��</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2'><table id='PurviewDetail' width='100%' border='0' cellspacing='10' cellpadding='0'"
    If rsAdmin("Purview") = 1 Then Response.Write "style='display:none'"
    Response.Write "><tr><td colspan='2' align='center'><strong>�� �� Ա Ȩ �� �� ϸ �� ��</strong></td></tr>"

    Dim rsChannel, sqlChannel
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
   
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        ChannelID = rsChannel("ChannelID")
        ChannelName = Trim(rsChannel("ChannelName"))
        ChannelShortName = Trim(rsChannel("ChannelShortName"))
        ChannelDir = rsChannel("ChannelDir")
        AdminPurview_Channel = rsAdmin("AdminPurview_" & ChannelDir)
        If IsNull(AdminPurview_Channel) Then AdminPurview_Channel = 4

        Response.Write "<tr valign='top'><td>"
        Response.Write "<fieldset><legend>�˹���Ա�ڡ�<font color='red'>" & ChannelName & "</font>��Ƶ����Ȩ�ޣ�</legend><table width='100%'><tr>"
        Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='';"""
        If AdminPurview_Channel = 1 Then Response.Write " checked"
        If rsChannel("ModuleType") = 4 Then
            Response.Write ">Ƶ������Ա��ӵ����Ŀ�ܱ��Ȩ�ޣ������Թ���IP���ú����ι��</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������Թ����������</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">��Ŀ����Ա��ֻӵ�в�������������Ȩ��</td>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write ">Ƶ������Ա��ӵ�������˲���Ƹģ��Ĺ���Ȩ�ޣ����Է���ְλ��Ϣ�͹����˲�</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">ְλ��Ϣ����Ա��ӵ��ְλ��Ϣ�ķ������޸�Ȩ�ޣ������ܹ����˲ź�ɾ��ְλ�������Ϣ(����˲���Ϣ��һ��ɾ��)</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">�˲���Ϣ����Ա��ӵ�������˲���Ϣ�Ĺ���Ȩ�ް���ɾ�����޸�ӦƸ�ߵļ���</td>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write ">Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ��</td>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ��</td>"
        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write ">Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ��</td>"
        Else
            Response.Write ">Ƶ������Ա��ӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">��Ŀ�ܱࣺӵ��������Ŀ�Ĺ���Ȩ�ޣ������������Ŀ��ר��</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">��Ŀ����Ա��ֻӵ�в�����Ŀ����Ȩ��</td>"
        End If
        Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='4' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
        If AdminPurview_Channel = 4 Then Response.Write " checked"
        Response.Write ">�ڴ�Ƶ�������κι���Ȩ��</td></tr>"
        If rsChannel("ModuleType") <> 6 Then
            Response.Write "<tr id='table_" & ChannelDir & "'"
            If AdminPurview_Channel = 3 Then
                Response.Write " style='display:'"
            Else
                Response.Write " style='display:none'"
            End If
             Response.Write ">"
        End If
        If rsChannel("ModuleType") = 4 Then
            Response.Write "<td width='60%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "'></iframe></td>"
            Response.Write "<td>"
            Response.Write "  <input name='arrKind_Modify_" & ChannelDir & "' type='hidden' id='arrKind_Modify_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Del_" & ChannelDir & "' type='hidden' id='arrKind_Del_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Move_" & ChannelDir & "' type='hidden' id='arrKind_Move_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Check_" & ChannelDir & "' type='hidden' id='arrKind_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Quintessence_" & ChannelDir & "' type='hidden' id='arrKind_Quintessence_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_SetOnTop_" & ChannelDir & "' type='hidden' id='arrKind_SetOnTop_" & ChannelDir & "'>"
            Response.Write "  <input name='arrKind_Reply_" & ChannelDir & "' type='hidden' id='arrKind_Reply_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            'Response.Write "<b>����Ȩ�ޣ�</b>"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'" & IsOtherChecked(PO, "AD_" & ChannelDir) & ">������"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">ģ�����&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write "<td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�"
            Response.Write "  <input name='arrHouseClass_View_" & ChannelDir & "' type='hidden' id='arrHouseClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Input_" & ChannelDir & "' type='hidden' id='arrHouseClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Check_" & ChannelDir & "' type='hidden' id='arrHouseClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Manage_" & ChannelDir & "' type='hidden' id='arrHouseClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='ClassConfig_" & ChannelDir & "'" & IsOtherChecked(PO, "ClassConfig_" & ChannelDir) & ">������Ŀ���ù���&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Area_" & ChannelDir & "'" & IsOtherChecked(PO, "Area_" & ChannelDir) & ">�����������&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">ģ�����&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"

        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">ģ�����&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'" & IsOtherChecked(PO, "Menu_" & ChannelDir) & ">�����˵�&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "<Table>"
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "''></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></Table>"
            Response.Write "</fieldset></td></tr>"
        Else
            Response.Write "<td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>ע��</font></strong>��ĿȨ�޲��ü̳��ƶȣ�����ĳһ��Ŀӵ��ĳ�����Ȩ�ޣ����ڴ���Ŀ����������Ŀ�ж�ӵ���������Ȩ�ޣ�����������Ŀ��ָ������Ĺ���Ȩ�ޡ�"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            Response.Write "<b>����Ȩ�ޣ�</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">ģ�����&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='JsFile_" & ChannelDir & "'" & IsOtherChecked(PO, "JsFile_" & ChannelDir) & ">JS�ļ�����&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'" & IsOtherChecked(PO, "Menu_" & ChannelDir) & ">�����˵�&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Keyword_" & ChannelDir & "'" & IsOtherChecked(PO, "Keyword_" & ChannelDir) & ">�ؼ��ֹ���&nbsp;"
            If rsChannel("ModuleType") = 5 Then
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Producer_" & ChannelDir & "'" & IsOtherChecked(PO, "Producer_" & ChannelDir) & ">���̹���&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Trademark_" & ChannelDir & "'" & IsOtherChecked(PO, "Trademark_" & ChannelDir) & ">Ʒ�ƹ���&nbsp;"
            Else
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Author_" & ChannelDir & "'" & IsOtherChecked(PO, "Author_" & ChannelDir) & ">���߹���&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Copyfrom_" & ChannelDir & "'" & IsOtherChecked(PO, "Copyfrom_" & ChannelDir) & ">��Դ����&nbsp;"
            End If
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='XML_" & ChannelDir & "'" & IsOtherChecked(PO, "XML_" & ChannelDir) & ">����XML&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Field_" & ChannelDir & "'" & IsOtherChecked(PO, "Field_" & ChannelDir) & ">�Զ����ֶ�&nbsp;"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'" & IsOtherChecked(PO, "AD_" & ChannelDir) & ">������"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    Response.Write "   <tr><td><fieldset><legend>�˹���Ա��������վ����Ȩ�ޣ�<input name='chkAll' type='checkbox' id='chkAll' value='Yes' onclick='SelectAll(this.form)'>ѡ������Ȩ��</legend>"
    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ModifyPwd'" & IsOtherChecked(PO, "ModifyPwd") & ">�޸��Լ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Channel'" & IsOtherChecked(PO, "Channel") & ">��վƵ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Collection'" & IsOtherChecked(PO, "Collection") & ">�ɼ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Message'" & IsOtherChecked(PO, "Message") & ">����Ϣ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='MailList'" & IsOtherChecked(PO, "MailList") & ">�ʼ��б����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD'" & IsOtherChecked(PO, "AD") & ">��վ������</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='FriendSite'" & IsOtherChecked(PO, "FriendSite") & ">�������ӹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Announce'" & IsOtherChecked(PO, "Announce") & ">��վ�������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Vote'" & IsOtherChecked(PO, "Vote") & ">��վ�������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Counter'" & IsOtherChecked(PO, "Counter") & ">��վͳ�ƹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Skin'" & IsOtherChecked(PO, "Skin") & ">��վ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template'" & IsOtherChecked(PO, "Template") & ">ͨ��ģ�����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Label'" & IsOtherChecked(PO, "Label") & ">�Զ����ǩ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowPage'" & IsOtherChecked(PO, "ShowPage") & ">�Զ���ҳ�����</td>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Cache'" & IsOtherChecked(PO, "Cache") & ">��վ�������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='KeyLink'" & IsOtherChecked(PO, "KeyLink") & ">վ�����ӹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Rtext'" & IsOtherChecked(PO, "Rtext") & ">�ַ����˹���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='UserGroup'" & IsOtherChecked(PO, "UserGroup") & ">��Ա�����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Card'" & IsOtherChecked(PO, "Card") & ">��ֵ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Equipment'" & IsOtherChecked(PO, "Equipment") & ">�ҳ��Ǽǹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='InfoManage'" & IsOtherChecked(PO, "InfoManage") & ">ѧ����Ϣ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ScoreManage'" & IsOtherChecked(PO, "ScoreManage") & ">ѧ���ɼ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='TestManage'" & IsOtherChecked(PO, "TestManage") & ">���Թ���</td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>��Ա����Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_View'" & IsOtherChecked(PO, "User_View") & ">�鿴��Ա��Ϣ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_ModifyInfo'" & IsOtherChecked(PO, "User_ModifyInfo") & ">�޸Ļ�Ա��Ϣ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_MofidyPurview'" & IsOtherChecked(PO, "User_MofidyPurview") & ">�޸Ļ�ԱȨ��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Lock'" & IsOtherChecked(PO, "User_Lock") & ">��ס/������Ա</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Del'" & IsOtherChecked(PO, "User_Del") & ">ɾ����Ա</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Update'" & IsOtherChecked(PO, "User_Update") & ">����Ϊ�ͻ�</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Money'" & IsOtherChecked(PO, "User_Money") & ">��Ա�ʽ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Point'" & IsOtherChecked(PO, "User_Point") & ">��Ա" & PointName & "����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Valid'" & IsOtherChecked(PO, "User_Valid") & ">��Ա��Ч�ڹ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ConsumeLog'" & IsOtherChecked(PO, "ConsumeLog") & ">��Ա������ϸ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='RechargeLog'" & IsOtherChecked(PO, "RechargeLog") & ">��Ա��Ч����ϸ</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�̳��ճ���������Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_View'" & IsOtherChecked(PO, "Order_View") & ">�鿴����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Confirm'" & IsOtherChecked(PO, "Order_Confirm") & ">ȷ�϶���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Modify'" & IsOtherChecked(PO, "Order_Modify") & ">�޸Ķ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Del'" & IsOtherChecked(PO, "Order_Del") & ">ɾ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Payment'" & IsOtherChecked(PO, "Order_Payment") & ">�տ��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Invoice'" & IsOtherChecked(PO, "Order_Invoice") & ">����Ʊ</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Deliver'" & IsOtherChecked(PO, "Order_Deliver") & ">�������ͣ�ʵ�</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Download'" & IsOtherChecked(PO, "Order_Download") & ">�������ͣ������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SendCard'" & IsOtherChecked(PO, "Order_SendCard") & ">�������ͣ��㿨��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_End'" & IsOtherChecked(PO, "Order_End") & ">���嶩��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Transfer'" & IsOtherChecked(PO, "Order_Transfer") & ">��������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Print'" & IsOtherChecked(PO, "Order_Print") & ">������ӡ</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Count'" & IsOtherChecked(PO, "Order_Count") & ">����ͳ��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_OrderItem'" & IsOtherChecked(PO, "Order_OrderItem") & ">������ϸ���</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SaleCount'" & IsOtherChecked(PO, "Order_SaleCount") & ">����ͳ��/����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Payment'" & IsOtherChecked(PO, "Payment") & ">����֧������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bankroll'" & IsOtherChecked(PO, "Bankroll") & ">�ʽ���ϸ��ѯ</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Deliver'" & IsOtherChecked(PO, "Deliver") & ">���˻���¼</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Transfer'" & IsOtherChecked(PO, "Transfer") & ">����������¼</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PresentProject'" & IsOtherChecked(PO, "PresentProject") & ">������������</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PaymentType'" & IsOtherChecked(PO, "PaymentType") & ">���ʽ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='DeliverType'" & IsOtherChecked(PO, "DeliverType") & ">�ͻ���ʽ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bank'" & IsOtherChecked(PO, "Bank") & ">�����ʻ�����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Refund'" & IsOtherChecked(PO, "Order_Refund") & ">�˿��</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShoppingCart'" & IsOtherChecked(PO, "ShoppingCart") & ">���ﳵ����</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddPayment'" & IsOtherChecked(PO, "AddPayment") & ">�������֧��</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AgentPayment'" & IsOtherChecked(PO, "AgentPayment") & ">���������֧��</td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"

    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�ͻ���ϵ����Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_View'" & IsOtherChecked(PO, "Client_View") & ">�鿴�ͻ���Ϣ</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Add'" & IsOtherChecked(PO, "Client_Add") & ">��ӿͻ�</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyOwn'" & IsOtherChecked(PO, "Client_ModifyOwn") & ">�޸������Լ��Ŀͻ���Ϣ</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyAll'" & IsOtherChecked(PO, "Client_ModifyAll") & ">�޸����пͻ���Ϣ</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Del'" & IsOtherChecked(PO, "Client_Del") & ">ɾ���ͻ�</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_View'" & IsOtherChecked(PO, "Service_View") & ">�鿴�����¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Add'" & IsOtherChecked(PO, "Service_Add") & ">��ӷ����¼</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyOwn'" & IsOtherChecked(PO, "Service_ModifyOwn") & ">�޸��Լ���ӵķ����¼</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyAll'" & IsOtherChecked(PO, "Service_ModifyAll") & ">�޸����з����¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Del'" & IsOtherChecked(PO, "Service_Del") & ">ɾ�������¼</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_View'" & IsOtherChecked(PO, "Complain_View") & ">�鿴Ͷ�߼�¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Add'" & IsOtherChecked(PO, "Complain_Add") & ">���Ͷ�߼�¼</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyOwn'" & IsOtherChecked(PO, "Complain_ModifyOwn") & ">�޸��Լ���ӵ�Ͷ�߼�¼</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyAll'" & IsOtherChecked(PO, "Complain_ModifyAll") & ">�޸�����Ͷ�߼�¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Del'" & IsOtherChecked(PO, "Complain_Del") & ">ɾ��Ͷ�߼�¼</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_View'" & IsOtherChecked(PO, "Call_View") & ">�鿴�طü�¼</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Add'" & IsOtherChecked(PO, "Call_Add") & ">��ӻطü�¼</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyOwn'" & IsOtherChecked(PO, "Call_ModifyOwn") & ">�޸��Լ���ӵĻطü�¼</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyAll'" & IsOtherChecked(PO, "Call_ModifyAll") & ">�޸����лطü�¼</td>"
    'Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Del'>ɾ��Ͷ�߼�¼</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�ֻ����Ź���Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToMember'" & IsOtherChecked(PO, "SendSMSToMember") & ">���͸���Ա</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToContacter'" & IsOtherChecked(PO, "SendSMSToContacter") & ">���͸���ϵ��</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToConsignee'" & IsOtherChecked(PO, "SendSMSToConsignee") & ">���͸������е��ջ���</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToOther'" & IsOtherChecked(PO, "SendSMSToOther") & ">���͸�������</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewMessageLog'" & IsOtherChecked(PO, "ViewMessageLog") & ">�鿴���ͽ��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SMS_MessageReceive'" & IsOtherChecked(PO, "SMS_MessageReceive") & ">�鿴���յ��Ķ���</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>�ʾ�������Ȩ��</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewSurvey'" & IsOtherChecked(PO, "ViewSurvey") & ">�鿴�ʾ�</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddSurvey'" & IsOtherChecked(PO, "AddSurvey") & ">�����ʾ�</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurvey'" & IsOtherChecked(PO, "ManageSurvey") & ">�����ʾ��޸ġ�ɾ����</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowSurveyCountData'" & IsOtherChecked(PO, "ShowSurveyCountData") & ">�鿴������</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurveyTemplate'" & IsOtherChecked(PO, "ManageSurveyTemplate") & ">�ʾ�ģ�����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ImportSurveyQuestion'" & IsOtherChecked(PO, "ImportSurveyQuestion") & ">�ʾ���Ŀ����</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ExportSurveyQuestion'" & IsOtherChecked(PO, "ExportSurveyQuestion") & ">�ʾ���Ŀ����</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewListQuestion'" & IsOtherChecked(PO, "ViewListQuestion") & ">�鿴�ʾ���Ŀ�б�</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "      </fieldset></td></tr>"
    Response.Write "  </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModifyPurview'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "    <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Admin.asp'"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
    rsAdmin.Close
    Set rsAdmin = Nothing
    Call ShowPurviewTips
    
End Sub

Sub SaveAdd()
    Dim strAdminName, UserName, Password, PwdConfirm, Purview, EnableMultiLogin
    Dim AdminPurview_Channel, AdminPurview_Others, ChannelPurview_Others
    Dim arrClass_View, arrClass_Input, arrClass_Check, arrClass_Manage
    Dim arrHouseClass_View, arrHouseClass_Input, arrHouseClass_Check, arrHouseClass_Manage
    Dim rsAdmin, sqlAdmin
    Dim arrKind_Modify, arrKind_Del, arrKind_Move, arrKind_Check, arrKind_Quintessence, arrKind_SetOnTop, arrKind_Reply, HouseEnable

    '��֤��ȫ��
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ƿ��ύ������!</li>"
    End If
   
    strAdminName = Trim(Request("AdminName"))
    UserName = Trim(Request("UserName"))
    Password = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    Purview = Trim(Request("purview"))
    EnableMultiLogin = Trim(Request("EnableMultiLogin"))
    AdminPurview_Others = ReplaceBadChar(Trim(Request("AdminPurview_Others")))
    ChannelPurview_Others = ReplaceBadChar(Trim(Request("ChannelPurview_Others")))
    If ChannelPurview_Others <> "" Then
        AdminPurview_Others = AdminPurview_Others & "," & ChannelPurview_Others
    End If

    If strAdminName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ա������Ϊ�գ�</li>"
    Else
        If CheckBadChar(strAdminName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ա���к��зǷ��ַ���</li>"
        End If
    End If
    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ǰ̨��Ա������Ϊ�գ�</li>"
    Else
        If CheckBadChar(UserName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ǰ̨��Ա���к��зǷ��ַ���</li>"
        End If
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ʼ���벻��Ϊ�գ�</li>"
    End If
    If PwdConfirm <> Password Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ȷ������������ʼ������ͬ��</li>"
    End If
    If CheckBadChar(Password) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ʼ�����к��зǷ��ַ���</li>"
    End If
    If Purview = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ԱȨ�޲���Ϊ�գ�</li>"
    Else
        Purview = PE_CLng(Purview)
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    
    Dim rsUser
    Set rsUser = Conn.Execute("Select * from PE_User where UserName='" & UserName & "'")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ǰ̨��Ա��</li>"
    End If
    Set rsUser = Nothing
    If FoundErr = True Then Exit Sub
    
    sqlAdmin = "Select * from PE_Admin where AdminName='" & strAdminName & "'"
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If Not (rsAdmin.BOF And rsAdmin.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ˹���Ա��</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub
    End If
    
    rsAdmin.addnew
    rsAdmin("AdminName") = strAdminName
    rsAdmin("UserName") = UserName
    rsAdmin("Password") = MD5(Password, 16)
    rsAdmin("LoginTimes") = 0
    rsAdmin("purview") = Purview
    rsAdmin("AdminPurview_Others") = AdminPurview_Others
    If EnableMultiLogin = "Yes" Then
        rsAdmin("EnableMultiLogin") = True
    Else
        rsAdmin("EnableMultiLogin") = False
    End If
    
    arrClass_View = ""
    arrClass_Input = ""
    arrClass_Check = ""
    arrClass_Manage = ""

    Dim rsChannel, sqlChannel
    HouseEnable = False
    sqlChannel = "select ChannelID,ChannelDir from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        ChannelID = rsChannel(0)
        ChannelDir = rsChannel(1)
        If rsChannel("ChannelID") = 998 Then
            HouseEnable = True
        End If
        AdminPurview_Channel = PE_CLng(Trim(Request("AdminPurview_" & ChannelDir)))
        rsAdmin("AdminPurview_" & ChannelDir) = AdminPurview_Channel
        If AdminPurview_Channel = 3 Then
            If ChannelID = 4 Then
                arrKind_Modify = ReplaceBadChar(Trim(Request("arrKind_Modify_" & ChannelDir)))
                arrKind_Del = ReplaceBadChar(Trim(Request("arrKind_Del_" & ChannelDir)))
                arrKind_Move = ReplaceBadChar(Trim(Request("arrKind_Move_" & ChannelDir)))
                arrKind_Check = ReplaceBadChar(Trim(Request("arrKind_Check_" & ChannelDir)))
                arrKind_Quintessence = ReplaceBadChar(Trim(Request("arrKind_Quintessence_" & ChannelDir)))
                arrKind_SetOnTop = ReplaceBadChar(Trim(Request("arrKind_SetOnTop_" & ChannelDir)))
                arrKind_Reply = ReplaceBadChar(Trim(Request("arrKind_Reply_" & ChannelDir)))
            ElseIf ChannelID = 998 Then
                arrHouseClass_View = ReplaceBadChar(Trim(Request("arrHouseClass_View_" & ChannelDir)))
                arrHouseClass_Input = ReplaceBadChar(Trim(Request("arrHouseClass_Input_" & ChannelDir)))
                arrHouseClass_Check = ReplaceBadChar(Trim(Request("arrHouseClass_Check_" & ChannelDir)))
                arrHouseClass_Manage = ReplaceBadChar(Trim(Request("arrHouseClass_Manage_" & ChannelDir)))
            Else
                If arrClass_View = "" Then
                    arrClass_View = ReplaceBadChar(Trim(Request("arrClass_View_" & ChannelDir)))
                Else
                    arrClass_View = arrClass_View & "," & ReplaceBadChar(Trim(Request("arrClass_View_" & ChannelDir)))
                End If
                If arrClass_Input = "" Then
                    arrClass_Input = ReplaceBadChar(Trim(Request("arrClass_Input_" & ChannelDir)))
                Else
                    arrClass_Input = arrClass_Input & "," & ReplaceBadChar(Trim(Request("arrClass_Input_" & ChannelDir)))
                End If
                If arrClass_Check = "" Then
                    arrClass_Check = ReplaceBadChar(Trim(Request("arrClass_Check_" & ChannelDir)))
                Else
                    arrClass_Check = arrClass_Check & "," & ReplaceBadChar(Trim(Request("arrClass_Check_" & ChannelDir)))
                End If
                If arrClass_Manage = "" Then
                    arrClass_Manage = ReplaceBadChar(Trim(Request("arrClass_Manage_" & ChannelDir)))
                Else
                    arrClass_Manage = arrClass_Manage & "," & ReplaceBadChar(Trim(Request("arrClass_Manage_" & ChannelDir)))
                End If
            End If
        End If
        rsChannel.MoveNext
    Loop
    Set rsChannel = Nothing
    
    rsAdmin("arrClass_View") = DelRightComma(Replace(arrClass_View, ",,", ","))
    rsAdmin("arrClass_Input") = DelRightComma(Replace(arrClass_Input, ",,", ","))
    rsAdmin("arrClass_Check") = DelRightComma(Replace(arrClass_Check, ",,", ","))
    rsAdmin("arrClass_Manage") = DelRightComma(Replace(arrClass_Manage, ",,", ","))
    If HouseEnable = True Then
        rsAdmin("arrClass_House") = arrHouseClass_View & "|||" & arrHouseClass_Input & "|||" & arrHouseClass_Check & "|||" & arrHouseClass_Manage
    End If
    rsAdmin("arrClass_GuestBook") = arrKind_Modify & "|||" & arrKind_Del & "|||" & arrKind_Move & "|||" & arrKind_Check & "|||" & arrKind_Quintessence & "|||" & arrKind_SetOnTop & "|||" & arrKind_Reply

    rsAdmin.Update
    rsAdmin.Close
    Set rsAdmin = Nothing
    Call WriteEntry(1, AdminName, "��������Ա��" & strAdminName)
    Call main
End Sub

Sub SaveModifyPwd()
    Dim UserID, UserName, Password, PwdConfirm, EnableMultiLogin
    Dim rsAdmin, sqlAdmin
    
    '��֤��ȫ��
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ƿ��ύ������!</li>"
    End If

    UserID = Trim(Request("ID"))
    UserName = Trim(Request("UserName"))
    Password = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    EnableMultiLogin = Trim(Request("EnableMultiLogin"))
    If UserID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
    Else
        UserID = PE_CLng(UserID)
    End If
    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ǰ̨��Ա������Ϊ�գ�</li>"
    Else
        If CheckBadChar(UserName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ǰ̨��Ա���к��зǷ��ַ���</li>"
        End If
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����벻��Ϊ�գ�</li>"
    End If
    If PwdConfirm <> Password Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ȷ�������������������ͬ��</li>"
    End If
    If CheckBadChar(Password) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������к��зǷ��ַ���</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    Dim rsUser
    Set rsUser = Conn.Execute("Select * from PE_User where UserName='" & UserName & "'")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ǰ̨��Ա��</li>"
    End If
    Set rsUser = Nothing
    If FoundErr = True Then Exit Sub
    
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˹���Ա��</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub
    End If
    rsAdmin("UserName") = UserName
    rsAdmin("password") = MD5(Password, 16)
    If EnableMultiLogin = "Yes" Then
        rsAdmin("EnableMultiLogin") = True
    Else
        rsAdmin("EnableMultiLogin") = False
    End If
    rsAdmin.Update
    rsAdmin.Close
    Set rsAdmin = Nothing
    Call WriteEntry(1, AdminName, "�޸Ĺ���Ա���룬ID��" & UserID)

    Call main
End Sub

Sub SaveModifyPurview()
    Dim UserID, UserName, Purview, EnableMultiLogin
    Dim AdminPurview_Channel, AdminPurview_Others, ChannelPurview_Others
    Dim OldAdminPurview_Channel
    Dim arrClass_View, arrClass_Input, arrClass_Check, arrClass_Manage
    Dim arrHouseClass_View, arrHouseClass_Input, arrHouseClass_Check, arrHouseClass_Manage
    Dim rsAdmin, sqlAdmin
    Dim arrKind_Modify, arrKind_Del, arrKind_Move, arrKind_Check, arrKind_Quintessence, arrKind_SetOnTop, arrKind_Reply, HouseEnable

    '��֤��ȫ��
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ƿ��ύ������!</li>"
    End If
    
    UserID = Trim(Request("ID"))
    Purview = Trim(Request("purview"))
    EnableMultiLogin = Trim(Request("EnableMultiLogin"))
    AdminPurview_Others = ReplaceBadChar(Trim(Request("AdminPurview_Others")))
    ChannelPurview_Others = ReplaceBadChar(Trim(Request("ChannelPurview_Others")))
    If ChannelPurview_Others <> "" Then
        AdminPurview_Others = AdminPurview_Others & "," & ChannelPurview_Others
    End If
    
    If UserID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
    Else
        UserID = PE_CLng(UserID)
    End If
    If Purview = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ԱȨ�޲���Ϊ�գ�</li>"
    Else
        Purview = PE_CLng(Purview)
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˹���Ա��</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub
    End If
    UserName = rsAdmin("UserName")
    rsAdmin("purview") = Purview
    If EnableMultiLogin = "Yes" Then
        rsAdmin("EnableMultiLogin") = True
    Else
        rsAdmin("EnableMultiLogin") = False
    End If
    rsAdmin("AdminPurview_Others") = AdminPurview_Others

    arrClass_View = ""
    arrClass_Input = ""
    arrClass_Check = ""
    arrClass_Manage = ""
    
    Dim sqlChannel, rsChannel
    HouseEnable = False
    sqlChannel = "select ChannelID,ChannelDir from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        ChannelID = rsChannel(0)
        ChannelDir = rsChannel(1)
        If rsChannel("ChannelID") = 998 Then
            HouseEnable = True
        End If
        AdminPurview_Channel = PE_CLng(Trim(Request("AdminPurview_" & ChannelDir)))
        rsAdmin("AdminPurview_" & ChannelDir) = AdminPurview_Channel
        If AdminPurview_Channel = 3 Then
            If ChannelID = 4 Then
                arrKind_Modify = ReplaceBadChar(Trim(Request("arrKind_Modify_" & ChannelDir)))
                arrKind_Del = ReplaceBadChar(Trim(Request("arrKind_Del_" & ChannelDir)))
                arrKind_Move = ReplaceBadChar(Trim(Request("arrKind_Move_" & ChannelDir)))
                arrKind_Check = ReplaceBadChar(Trim(Request("arrKind_Check_" & ChannelDir)))
                arrKind_Quintessence = ReplaceBadChar(Trim(Request("arrKind_Quintessence_" & ChannelDir)))
                arrKind_SetOnTop = ReplaceBadChar(Trim(Request("arrKind_SetOnTop_" & ChannelDir)))
                arrKind_Reply = ReplaceBadChar(Trim(Request("arrKind_Reply_" & ChannelDir)))
            ElseIf ChannelID = 998 Then
                arrHouseClass_View = ReplaceBadChar(Trim(Request("arrHouseClass_View_" & ChannelDir)))
                arrHouseClass_Input = ReplaceBadChar(Trim(Request("arrHouseClass_Input_" & ChannelDir)))
                arrHouseClass_Check = ReplaceBadChar(Trim(Request("arrHouseClass_Check_" & ChannelDir)))
                arrHouseClass_Manage = ReplaceBadChar(Trim(Request("arrHouseClass_Manage_" & ChannelDir)))
            Else
                If arrClass_View = "" Then
                    arrClass_View = ReplaceBadChar(Trim(Request("arrClass_View_" & ChannelDir)))
                Else
                    arrClass_View = arrClass_View & "," & ReplaceBadChar(Trim(Request("arrClass_View_" & ChannelDir)))
                End If
                If arrClass_Input = "" Then
                    arrClass_Input = ReplaceBadChar(Trim(Request("arrClass_Input_" & ChannelDir)))
                Else
                    arrClass_Input = arrClass_Input & "," & ReplaceBadChar(Trim(Request("arrClass_Input_" & ChannelDir)))
                End If
                If arrClass_Check = "" Then
                    arrClass_Check = ReplaceBadChar(Trim(Request("arrClass_Check_" & ChannelDir)))
                Else
                    arrClass_Check = arrClass_Check & "," & ReplaceBadChar(Trim(Request("arrClass_Check_" & ChannelDir)))
                End If
                If arrClass_Manage = "" Then
                    arrClass_Manage = ReplaceBadChar(Trim(Request("arrClass_Manage_" & ChannelDir)))
                Else
                    arrClass_Manage = arrClass_Manage & "," & ReplaceBadChar(Trim(Request("arrClass_Manage_" & ChannelDir)))
                End If
            End If
        End If
        rsChannel.MoveNext
    Loop
    Set rsChannel = Nothing
    
    rsAdmin("arrClass_View") = Replace(arrClass_View, ",,", ",")
    rsAdmin("arrClass_Input") = Replace(arrClass_Input, ",,", ",")
    rsAdmin("arrClass_Check") = Replace(arrClass_Check, ",,", ",")
    rsAdmin("arrClass_Manage") = Replace(arrClass_Manage, ",,", ",")
    If HouseEnable = True Then
        rsAdmin("arrClass_House") = arrHouseClass_View & "|||" & arrHouseClass_Input & "|||" & arrHouseClass_Check & "|||" & arrHouseClass_Manage
    End If
    rsAdmin("arrClass_GuestBook") = arrKind_Modify & "|||" & arrKind_Del & "|||" & arrKind_Move & "|||" & arrKind_Check & "|||" & arrKind_Quintessence & "|||" & arrKind_SetOnTop & "|||" & arrKind_Reply

    rsAdmin.Update
    rsAdmin.Close
    Set rsAdmin = Nothing
    Call WriteEntry(1, AdminName, "�޸Ĺ���ԱȨ�ޣ�ID��" & UserID)

    Call main
End Sub

Sub DelAdmin()
    Dim UserID
    Dim rsAdmin, sqlAdmin
    Dim rsChannel, sqlChannel

    '��֤��ȫ��
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ƿ��ύ������!</li>"
    End If

    UserID = Trim(Request("ID"))
    If IsValidID(UserID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ���Ĺ���ԱID</li>"
        Exit Sub
    End If
    If InStr(UserID, ",") > 0 Then
        Conn.Execute ("delete from PE_Admin where ID in (" & UserID & ")")
    Else
        Conn.Execute ("delete from PE_Admin where ID=" & UserID & "")
    End If
    Call WriteEntry(1, AdminName, "ɾ������Ա��ID��" & UserID)

    Call main
End Sub

Function IsOtherChecked(Purview_Others, strOthers)
    If CheckPurview_Other(Purview_Others, strOthers) = True Then
        IsOtherChecked = " checked"
    Else
        IsOtherChecked = ""
    End If
End Function

Sub ShowPwdStrength()
    Response.Write "<script language='JavaScript' src='PwdStrength.js'></script>" & vbCrLf
    Response.Write "<script language='JavaScript'>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "window.onerror = ignoreError;" & vbCrLf
    Response.Write "function ignoreError(){return true;}" & vbCrLf
    Response.Write "function EvalPwdStrength(oF,sP){" & vbCrLf
    Response.Write "  PadPasswd(oF,sP.length*2);" & vbCrLf
    Response.Write "  if(ClientSideStrongPassword(sP,gSimilarityMap,gDictionary)){DispPwdStrength(3,'cssStrong');}" & vbCrLf
    Response.Write "  else if(ClientSideMediumPassword(sP,gSimilarityMap,gDictionary)){DispPwdStrength(2,'cssMedium');}" & vbCrLf
    Response.Write "  else if(ClientSideWeakPassword(sP,gSimilarityMap,gDictionary)){DispPwdStrength(1,'cssWeak');}" & vbCrLf
    Response.Write "  else{DispPwdStrength(0,'cssPWD');}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function PadPasswd(oF,lPwd){" & vbCrLf
    Response.Write "  if(typeof oF.PwdPad=='object'){var sPad='IfYouAreReadingThisYouHaveTooMuchFreeTime';var lPad=sPad.length-lPwd;oF.PwdPad.value=sPad.substr(0,(lPad<0)?0:lPad);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function DispPwdStrength(iN,sHL){" & vbCrLf
    Response.Write "  if(iN>3){ iN=3;}for(var i=0;i<4;i++){ var sHCR='cssPWD';if(i<=iN){ sHCR=sHL;}if(i>0){ GEId('idSM'+i).className=sHCR;}GEId('idSMT'+i).style.display=((i==iN)?'inline':'none');}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function GEId(sID){return document.getElementById(sID);}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<style>" & vbCrLf
    Response.Write "input{FONT-FAMILY:����;FONT-SIZE: 9pt;}" & vbCrLf
    Response.Write ".cssPWD{background-color:#EBEBEB;border-right:solid 1px #BEBEBE;border-bottom:solid 1px #BEBEBE;}" & vbCrLf
    Response.Write ".cssWeak{background-color:#FF4545;border-right:solid 1px #BB2B2B;border-bottom:solid 1px #BB2B2B;}" & vbCrLf
    Response.Write ".cssMedium{background-color:#FFD35E;border-right:solid 1px #E9AE10;border-bottom:solid 1px #E9AE10;}" & vbCrLf
    Response.Write ".cssStrong{background-color:#3ABB1C;border-right:solid 1px #267A12;border-bottom:solid 1px #267A12;}" & vbCrLf
    Response.Write ".cssPWT{width:132px;}" & vbCrLf
    Response.Write "</style>" & vbCrLf
    Response.Write "<table cellpadding='0' cellspacing='0' class='cssPWT' style='height:16px'><tr valign='bottom'><td id='idSM1' width='33%' class='cssPWD' align='center'><span style='font-size:1px'>&nbsp;</span><span id='idSMT1' style='display:none;'>��</span></td><td id='idSM2' width='34%' class='cssPWD' align='center' style='border-left:solid 1px #fff'><span style='font-size:1px'>&nbsp;</span><span id='idSMT0' style='display:inline;font-weight:normal;color:#666'>��</span><span id='idSMT2' style='display:none;'>��</span></td><td id='idSM3' width='33%' class='cssPWD' align='center' style='border-left:solid 1px #fff'><span style='font-size:1px'>&nbsp;</span><span id='idSMT3' style='display:none;'>ǿ</span></td></tr></table>"
End Sub

Sub ShowPurviewTips()
    Response.Write "<b><font color='#FF0000'>Ƶ��Ȩ��˵����</font></b>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td width='150'><b>������Ŀ</b></td>" & vbCrLf
    Response.Write "    <td><b>Ƶ������ԱȨ��</b></td>" & vbCrLf
    Response.Write "    <td><b>��Ŀ�ܱ�Ȩ��</b></td>" & vbCrLf
    Response.Write "    <td><b>��Ŀ����ԱȨ��</b></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�����Ϣ</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>������������Ŀ�����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>������������Ŀ�����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>ֻ������¼��Ȩ�޵���Ŀ�����Ϣ</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�޸���Ϣ</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>�����޸�������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>�����޸�������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>ֻ���޸��й���Ȩ�޵���Ŀ�е���Ϣ�������Լ���ӵ���Ϣ</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>ɾ����Ϣ</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>����ɾ��������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>����ɾ��������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>ֻ�����й���Ȩ�޵���Ŀ��ɾ����Ϣ�������Լ���ӵ���Ϣ</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�����Ϣ</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>�������������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>�������������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>ֻ���������Ȩ�޵���Ŀ���ƶ���Ϣ</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>������Ϣ���̶����Ƽ���</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>���Թ���������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>���Թ���������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>ֻ�����й���Ȩ�޵���Ŀ�й�����Ϣ</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>����HTML��������</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>��������������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>��������������Ŀ����Ϣ</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>ֻ�����й���Ȩ�޵���Ŀ��������Ϣ</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>����HTML���Զ���</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>������Ϣʱ�Զ�����HTML</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>������Ϣʱ�Զ�����HTML</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>������Ϣʱ�Զ�����HTML</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�ƶ���Ϣ</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>������������</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>ר����Ϣ����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>���۹���</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>����վ����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>��Ŀ����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>ר�����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�ϴ��ļ�����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>ģ�����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>JS�ļ�����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�����˵�����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�ؼ��ֹ���</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>���߹���</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>��Դ����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td><b>������ĿXML����</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>�Զ����ֶ�</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>������</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>����ָ��</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>����</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

%>
