<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

strFileName = "Admin_Admin.asp"

Response.Write "<html><head><title>管理员管理</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("管 理 员 管 理", 10049)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td height='30'><a href='Admin_Admin.asp'>管理员管理首页</a>&nbsp;|&nbsp;<a href='Admin_Admin.asp?Action=Add'>新增管理员</a>"
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
    Call WriteEntry(1, AdminName, "管理员操作失败，失败原因：" & ErrMsg)
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
        Response.Write "没有任何管理员！"
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
    Response.Write "  <form name='myform' method='Post' action='Admin_Admin.asp' onsubmit=""return confirm('确定要删除选中的管理员吗？');"">"
    Response.Write "     <td>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title' height='22'>"
    Response.Write "    <td  width='30'><strong>选中</strong></td>"
    Response.Write "    <td  width='30' height='22'><strong>序号</strong></td>"
    Response.Write "    <td><strong>管理员名</strong></td>"
    Response.Write "    <td><strong>前台会员名</strong></td>"
    Response.Write "    <td width='70'><strong>权 限</strong></td>"
    Response.Write "    <td width='55'><strong>多人登录</strong></td>"
    Response.Write "    <td width='95'><strong>最后登录IP</strong></td>"
    Response.Write "    <td width='115'><strong>最后登录时间</strong></td>"
    Response.Write "    <td width='55'><strong>登录次数</strong></td>"
    Response.Write "    <td width='180'><strong>操 作</strong></td>"
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
              Response.Write "<font color=blue>超级管理员</font>"
            Case 2
              Response.Write "普通管理员"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='55'>"
        If rsAdminList("EnableMultiLogin") = True Then
            Response.Write "<font color='green'>允许</font>"
        Else
            Response.Write "<font color='red'>不允许</font>"
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
        Response.Write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rsAdminList("ID") & "'>修改密码及设置</a>&nbsp;&nbsp;"
        Response.Write "<a href='Admin_Admin.asp?Action=ModifyPurview&ID=" & rsAdminList("ID") & "'>修改权限</a>&nbsp;&nbsp;"
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
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> 选中本页显示的所有管理员</td>"
    Response.Write "    <td><input name='Action' type='hidden' id='Action' value='Del'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='删除选中的管理员'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个管理员", True)
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
    Response.Write "      alert('用户名不能为空！');" & vbCrLf
    Response.Write "   document.form1.username.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if(document.form1.Password.value==''){" & vbCrLf
    Response.Write "      alert('密码不能为空！');" & vbCrLf
    Response.Write "   document.form1.Password.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value)){" & vbCrLf
    Response.Write "      alert('初始密码与确认密码不同！');" & vbCrLf
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
    Response.Write "      alert('密码不能为空！');" & vbCrLf
    Response.Write "   document.form1.Password.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value)){" & vbCrLf
    Response.Write "      alert('初始密码与确认密码不同！');" & vbCrLf
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
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>新 增 管 理 员</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>管理员名：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='AdminName' type='text'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>前台会员名：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='username' type='text'> <a href='Admin_User.asp?Action=AddUser'>添加新会员</a></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>初始密码：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input type='password' name='Password' onkeyup='javascript:EvalPwdStrength(document.forms[0],this.value);' onmouseout='javascript:EvalPwdStrength(document.forms[0],this.value);' onblur='javascript:EvalPwdStrength(document.forms[0],this.value);'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>密码强度：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'>"
    Call ShowPwdStrength
    Response.Write "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>确认密码：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input type='password' name='PwdConfirm'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%'>&nbsp;</td>"
    Response.Write "      <td width='88%'><input name='EnableMultiLogin' type='checkbox' value='Yes'>允许多人同时使用此帐号登录</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='12%' align='right' class='tdbg'><strong>权限设置： </strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='Purview' type='radio' value='1' onClick=""PurviewDetail.style.display='none'"">超级管理员：拥有所有权限。某些权限（如管理员管理、网站信息配置、网站选项配置等管理权限）只有超级管理员才有。"
    Response.Write "  <br><input type='radio' name='Purview' value='2' checked  onClick=""PurviewDetail.style.display=''"">普通管理员：需要详细指定每一项管理权限</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2'><table id='PurviewDetail' width='100%' border='0' cellspacing='10' cellpadding='0' style='display:'>"
    Response.Write "        <tr>"
    Response.Write "          <td colspan='2' align='center'><strong>管 理 员 权 限 详 细 设 置</strong></td>"
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
        Response.Write "<fieldset><legend>此管理员在【<font color='red'>" & ChannelName & "</font>】频道的权限：</legend><table width='100%'><tr>"
        If rsChannel("ModuleType") = 4 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='';"">频道管理员：拥有栏目总编的权限，并可以管理IP来访和屏蔽广告</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目总编：拥有所有栏目的管理权限，并可以管理留言类别</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目管理员：只拥有部分留言类别管理权限</td>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""OtherPurview_" & ChannelDir & ".style.display='';"">频道管理员：拥有所有人才招聘模块的管理权限，可以发布职位信息和管理人才</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""OtherPurview_" & ChannelDir & ".style.display='none';"">职位信息管理员：拥有职位信息的发布和修改权限，但不能管理人才和删除职位及相关信息(相关人才信息会一起被删除)</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""OtherPurview_" & ChannelDir & ".style.display='none';"">人才信息管理员：拥有所有人才信息的管理权限包括删除和修改应聘者的简历</td>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""OtherPurview_" & ChannelDir & ".style.display='';table_" & ChannelDir & ".style.display='none'"">频道管理员：拥有所有栏目的管理权限</td>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目管理员：只拥有部分栏目管理权限</td>"
        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""OtherPurview_" & ChannelDir & ".style.display='';"">频道管理员：拥有所有栏目的管理权限，并可以添加栏目和专题</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目总编：拥有所有栏目的管理权限，但不能添加栏目和专题</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目管理员：只拥有部分栏目管理权限</td>"
        Else
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='';"">频道管理员：拥有所有栏目的管理权限，并可以添加栏目和专题</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目总编：拥有所有栏目的管理权限，但不能添加栏目和专题</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"">栏目管理员：只拥有部分栏目管理权限</td>"
        End If
        Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='4' checked onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"">在此频道里无任何管理权限</td></tr>"
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
            'Response.Write "<b>更多权限：</b>"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'>广告管理"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>模板管理&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>注：</font></strong>栏目权限采用继承制度，即在某一栏目拥有某项管理权限，则在此栏目的所有子栏目中都拥有这项管理权限，并可在子栏目中指定更多的管理权限。"
            Response.Write "  <input name='arrHouseClass_View_" & ChannelDir & "' type='hidden' id='arrHouseClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Input_" & ChannelDir & "' type='hidden' id='arrHouseClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Check_" & ChannelDir & "' type='hidden' id='arrHouseClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrHouseClass_Manage_" & ChannelDir & "' type='hidden' id='arrHouseClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='ClassConfig_" & ChannelDir & "'>房产栏目配置管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Area_" & ChannelDir & "'>房产区域管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>模板管理&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>模板管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'>顶部菜单&nbsp;"
            Response.Write "</td></tr></table>"
            
            Response.Write "<Table>"
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>注：</font></strong>栏目权限采用继承制度，即在某一栏目拥有某项管理权限，则在此栏目的所有子栏目中都拥有这项管理权限，并可在子栏目中指定更多的管理权限。"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></Table>"

            Response.Write "</fieldset></td></tr>"

        Else
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>注：</font></strong>栏目权限采用继承制度，即在某一栏目拥有某项管理权限，则在此栏目的所有子栏目中都拥有这项管理权限，并可在子栏目中指定更多的管理权限。"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2' style='display:none'><tr><td>"
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'>模板管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='JsFile_" & ChannelDir & "'>JS文件管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'>顶部菜单&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Keyword_" & ChannelDir & "'>关键字管理&nbsp;"
            If rsChannel("ModuleType") = 5 Then
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Producer_" & ChannelDir & "'>厂商管理&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Trademark_" & ChannelDir & "'>品牌管理&nbsp;"
            Else
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Author_" & ChannelDir & "'>作者管理&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Copyfrom_" & ChannelDir & "'>来源管理&nbsp;"
            End If
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='XML_" & ChannelDir & "'>更新XML&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Field_" & ChannelDir & "'>自定义字段&nbsp;"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'>广告管理"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    
    
    Response.Write "   <tr><td><fieldset><legend>此管理员的其他网站管理权限：<input name='chkAll' type='checkbox' id='chkAll' value='Yes' onclick='SelectAll(this.form)'>选中所有权限</legend>"
    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ModifyPwd' checked>修改自己密码</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Channel'>网站频道管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Collection'>采集管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Message'>短消息管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='MailList'>邮件列表管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD'>网站广告管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='FriendSite'>友情链接管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Announce'>网站公告管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Vote'>网站调查管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Counter'>网站统计管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Skin'>网站风格管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template0'>通用模板管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Label'>自定义标签管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowPage'>自定义页面管理</td>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Cache'>网站缓存管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='KeyLink'>站内链接管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Rtext'>字符过滤管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='UserGroup'>会员组管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Card'>充值卡管理</td>"

    If FoundInArr(AllModules, "Classroom", ",") Then
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Equipment'>室场登记管理</td>"
    End If
    If FoundInArr(AllModules, "Sdms", ",") Then
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='InfoManage'>学生信息管理</td>"
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ScoreManage'>学生成绩管理</td>"
        Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='TestManage'>考试管理</td>"
    End If
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>会员管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_View'>查看会员信息</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_ModifyInfo'>修改会员信息</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_MofidyPurview'>修改会员权限</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Lock'>锁住/解锁会员</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Del'>删除会员</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Update'>升级为客户</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Money'>会员资金管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Point'>会员" & PointName & "管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Valid'>会员有效期管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ConsumeLog'>会员消费明细</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='RechargeLog'>会员有效期明细</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>商城日常操作管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_View'>查看订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Confirm'>确认订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Modify'>修改订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Del'>删除订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Payment'>收款处理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Invoice'>开发票</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Deliver'>订单配送（实物）</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Download'>订单配送（软件）</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SendCard'>订单配送（点卡）</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_End'>结清订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Transfer'>订单过户</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Print'>订单打印</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Count'>订单统计</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_OrderItem'>销售明细情况</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SaleCount'>销售统计/排行</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Payment'>在线支付管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bankroll'>资金明细查询</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Deliver'>发退货记录</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Transfer'>订单过户记录</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PresentProject'>促销方案管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PaymentType'>付款方式管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='DeliverType'>送货方式管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bank'>银行帐户管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Refund'>退款处理</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShoppingCart'>购物车管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddPayment'>虚拟货币支付</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AgentPayment'>代理商余额支付</td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"


    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>客户关系管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_View'>查看客户信息</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Add'>添加客户</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyOwn'>修改属于自己的客户信息</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyAll'>修改所有客户信息</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Del'>删除客户</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_View'>查看服务记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Add'>添加服务记录</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyOwn'>修改自己添加的服务记录</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyAll'>修改所有服务记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Del'>删除服务记录</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_View'>查看投诉记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Add'>添加投诉记录</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyOwn'>修改自己添加的投诉记录</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyAll'>修改所有投诉记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Del'>删除投诉记录</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_View'>查看回访记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Add'>添加回访记录</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyOwn'>修改自己添加的回访记录</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyAll'>修改所有回访记录</td>"
    'Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Del'>删除投诉记录</td>"
    Response.Write "          </tr>"

    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>手机短信管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToMember'>发送给会员</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToContacter'>发送给联系人</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToConsignee'>发送给订单中的收货人</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToOther'>发送给其他人</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewMessageLog'>查看发送结果</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SMS_MessageReceive'>查看接收到的短信</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>问卷调查管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewSurvey'>查看问卷</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddSurvey'>创建问卷</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurvey'>管理问卷（修改、删除）</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowSurveyCountData'>查看调查结果</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurveyTemplate'>问卷模板管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ImportSurveyQuestion'>问卷题目导入</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ExportSurveyQuestion'>问卷题目导出</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewListQuestion'>查看问卷题目列表</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"


    Response.Write "      </fieldset></td></tr>"
    Response.Write "  </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Admin.asp';"" style='cursor:hand;'></td>"
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
        ErrMsg = ErrMsg & "<li>请指定要修改的管理员ID</li>"
        Exit Sub
    Else
        UserID = PE_CLng(UserID)
    End If
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不存在此管理员！</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub

    End If
    Call ShowJS_Check
    Response.Write "<form method='post' action='Admin_Admin.asp' name='form1' onsubmit='return CheckModifyPwd();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修 改 管 理 员 密 码</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>管理员名：</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'>" & rsAdmin("AdminName") & "<input name='ID' type='hidden' value='" & rsAdmin("ID") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>前台会员名：</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'><input name='UserName' type='text' value='" & rsAdmin("UserName") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td width='40%' class='tdbg'><strong>新 密 码：</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'><input type='password' name='Password' onkeyup='javascript:EvalPwdStrength(document.forms[0],this.value);' onmouseout='javascript:EvalPwdStrength(document.forms[0],this.value);' onblur='javascript:EvalPwdStrength(document.forms[0],this.value);'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>密码强度：</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'>"
    Call ShowPwdStrength
    Response.Write "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='40%' class='tdbg'><strong>确认密码：</strong></td>"
    Response.Write "      <td width='60%' class='tdbg'><input type='password' name='PwdConfirm'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='40%'>&nbsp;</td>"
    Response.Write "      <td width='60%'><input name='EnableMultiLogin' type='checkbox' value='Yes'"
    If rsAdmin("EnableMultiLogin") = True Then Response.Write " checked"
    Response.Write ">允许多人同时使用此帐号登录</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModifyPwd'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "        <input  type='submit' name='Submit' value='保存修改结果' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Admin.asp'"" style='cursor:hand;'></td>"
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
        ErrMsg = ErrMsg & "<li>请指定要修改的管理员ID</li>"
        Exit Sub
    Else
        UserID = PE_CLng(UserID)
    End If
    
    
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不存在此管理员！</li>"
        rsAdmin.Close
        Set rsAdmin = Nothing
        Exit Sub
    End If

    PO = rsAdmin("AdminPurview_Others")
    Call ShowJS_Check
    
    Response.Write "<form method='post' action='Admin_Admin.asp' name='form1' onsubmit='javascript:CheckModifyPurview();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修 改 管 理 员 权 限</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='12%' class='tdbg'><strong>管理员名：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'>" & rsAdmin("AdminName") & "<input name='ID' type='hidden' value='" & rsAdmin("ID") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr> "
    Response.Write "      <td width='12%' class='tdbg'><strong>前台会员名：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'>" & rsAdmin("UserName") & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='12%'>&nbsp;</td>"
    Response.Write "      <td width='88%'><input name='EnableMultiLogin' type='checkbox' value='Yes'"
    If rsAdmin("EnableMultiLogin") = True Then Response.Write " checked"
    Response.Write ">允许多人同时使用此帐号登录</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='12%' class='tdbg'><strong>权限设置：</strong></td>"
    Response.Write "      <td width='88%' class='tdbg'><input name='Purview' type='radio' value='1' onClick=""PurviewDetail.style.display='none'"""
    If rsAdmin("Purview") = 1 Then Response.Write "checked"
    Response.Write ">超级管理员：拥有所有权限。某些权限（如管理员管理、网站信息配置、网站选项配置等管理权限）只有超级管理员才有<br>"
    Response.Write "<input type='radio' name='Purview' value='2' onClick=""PurviewDetail.style.display=''"""
    If rsAdmin("Purview") = 2 Then Response.Write "checked"
    Response.Write ">普通管理员：需要详细指定每一项管理权限</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2'><table id='PurviewDetail' width='100%' border='0' cellspacing='10' cellpadding='0'"
    If rsAdmin("Purview") = 1 Then Response.Write "style='display:none'"
    Response.Write "><tr><td colspan='2' align='center'><strong>管 理 员 权 限 详 细 设 置</strong></td></tr>"

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
        Response.Write "<fieldset><legend>此管理员在【<font color='red'>" & ChannelName & "</font>】频道的权限：</legend><table width='100%'><tr>"
        Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='1' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='';"""
        If AdminPurview_Channel = 1 Then Response.Write " checked"
        If rsChannel("ModuleType") = 4 Then
            Response.Write ">频道管理员：拥有栏目总编的权限，并可以管理IP来访和屏蔽广告</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">栏目总编：拥有所有栏目的管理权限，并可以管理留言类别</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">栏目管理员：只拥有部分留言类别管理权限</td>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write ">频道管理员：拥有所有人才招聘模块的管理权限，可以发布职位信息和管理人才</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">职位信息管理员：拥有职位信息的发布和修改权限，但不能管理人才和删除职位及相关信息(相关人才信息会一起被删除)</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">人才信息管理员：拥有所有人才信息的管理权限包括删除和修改应聘者的简历</td>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write ">频道管理员：拥有所有栏目的管理权限</td>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">栏目管理员：只拥有部分栏目管理权限</td>"
        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write ">频道管理员：拥有所有栏目的管理权限，并可以添加栏目和专题</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">栏目总编：拥有所有栏目的管理权限，但不能添加栏目和专题</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">栏目管理员：只拥有部分栏目管理权限</td>"
        Else
            Response.Write ">频道管理员：拥有所有栏目的管理权限，并可以添加栏目和专题</td>"
            Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='2' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 2 Then Response.Write " checked"
            Response.Write ">栏目总编：拥有所有栏目的管理权限，但不能添加栏目和专题</td></tr>"
            Response.Write "<tr><td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='3' onclick=""table_" & ChannelDir & ".style.display='';OtherPurview_" & ChannelDir & ".style.display='none';"""
            If AdminPurview_Channel = 3 Then Response.Write " checked"
            Response.Write ">栏目管理员：只拥有部分栏目管理权限</td>"
        End If
        Response.Write "<td width='50%'><input type='radio' name='AdminPurview_" & ChannelDir & "' value='4' onclick=""table_" & ChannelDir & ".style.display='none';OtherPurview_" & ChannelDir & ".style.display='none';"""
        If AdminPurview_Channel = 4 Then Response.Write " checked"
        Response.Write ">在此频道里无任何管理权限</td></tr>"
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
            'Response.Write "<b>更多权限：</b>"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'" & IsOtherChecked(PO, "AD_" & ChannelDir) & ">广告管理"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 8 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">模板管理&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        ElseIf rsChannel("ModuleType") = 7 Then
            Response.Write "<td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>注：</font></strong>栏目权限采用继承制度，即在某一栏目拥有某项管理权限，则在此栏目的所有子栏目中都拥有这项管理权限，并可在子栏目中指定更多的管理权限。"
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
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='ClassConfig_" & ChannelDir & "'" & IsOtherChecked(PO, "ClassConfig_" & ChannelDir) & ">房产栏目配置管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Area_" & ChannelDir & "'" & IsOtherChecked(PO, "Area_" & ChannelDir) & ">房产区域管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">模板管理&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"

        ElseIf rsChannel("ModuleType") = 6 Then
            Response.Write "</td></tr></table>"
            Response.Write "<table id='OtherPurview_" & ChannelDir & "' width='100%' border='0' cellspacing='1' cellpadding='2'"
            If AdminPurview_Channel > 1 Then
                Response.Write " style='display:none'"
            End If
            Response.Write "><tr><td>"
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">模板管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'" & IsOtherChecked(PO, "Menu_" & ChannelDir) & ">顶部菜单&nbsp;"
            Response.Write "</td></tr></table>"
            Response.Write "<Table>"
            Response.Write "<tr id='table_" & ChannelDir & "' style='display:none'><td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "''></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>注：</font></strong>栏目权限采用继承制度，即在某一栏目拥有某项管理权限，则在此栏目的所有子栏目中都拥有这项管理权限，并可在子栏目中指定更多的管理权限。"
            Response.Write "  <input name='arrClass_View_" & ChannelDir & "' type='hidden' id='arrClass_View_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Input_" & ChannelDir & "' type='hidden' id='arrClass_Input_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Check_" & ChannelDir & "' type='hidden' id='arrClass_Check_" & ChannelDir & "'>"
            Response.Write "  <input name='arrClass_Manage_" & ChannelDir & "' type='hidden' id='arrClass_Manage_" & ChannelDir & "'>"
            Response.Write "</td></tr></Table>"
            Response.Write "</fieldset></td></tr>"
        Else
            Response.Write "<td width='50%'>"
            Response.Write "<iframe id='frm" & ChannelDir & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Admin&ChannelID=" & ChannelID & "&Action=Modify&UserID=" & UserID & "'></iframe></td>"
            Response.Write "<td><strong><font color='#0000FF'>注：</font></strong>栏目权限采用继承制度，即在某一栏目拥有某项管理权限，则在此栏目的所有子栏目中都拥有这项管理权限，并可在子栏目中指定更多的管理权限。"
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
            Response.Write "<b>更多权限：</b>"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template_" & ChannelDir & "'" & IsOtherChecked(PO, "Template_" & ChannelDir) & ">模板管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='JsFile_" & ChannelDir & "'" & IsOtherChecked(PO, "JsFile_" & ChannelDir) & ">JS文件管理&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Menu_" & ChannelDir & "'" & IsOtherChecked(PO, "Menu_" & ChannelDir) & ">顶部菜单&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Keyword_" & ChannelDir & "'" & IsOtherChecked(PO, "Keyword_" & ChannelDir) & ">关键字管理&nbsp;"
            If rsChannel("ModuleType") = 5 Then
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Producer_" & ChannelDir & "'" & IsOtherChecked(PO, "Producer_" & ChannelDir) & ">厂商管理&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Trademark_" & ChannelDir & "'" & IsOtherChecked(PO, "Trademark_" & ChannelDir) & ">品牌管理&nbsp;"
            Else
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Author_" & ChannelDir & "'" & IsOtherChecked(PO, "Author_" & ChannelDir) & ">作者管理&nbsp;"
                Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Copyfrom_" & ChannelDir & "'" & IsOtherChecked(PO, "Copyfrom_" & ChannelDir) & ">来源管理&nbsp;"
            End If
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='XML_" & ChannelDir & "'" & IsOtherChecked(PO, "XML_" & ChannelDir) & ">更新XML&nbsp;"
            Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='Field_" & ChannelDir & "'" & IsOtherChecked(PO, "Field_" & ChannelDir) & ">自定义字段&nbsp;"
            'Response.Write "<input name='ChannelPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD_" & ChannelDir & "'" & IsOtherChecked(PO, "AD_" & ChannelDir) & ">广告管理"
            Response.Write "</td></tr></table>"
            Response.Write "</fieldset></td></tr>"
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    Response.Write "   <tr><td><fieldset><legend>此管理员的其他网站管理权限：<input name='chkAll' type='checkbox' id='chkAll' value='Yes' onclick='SelectAll(this.form)'>选中所有权限</legend>"
    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ModifyPwd'" & IsOtherChecked(PO, "ModifyPwd") & ">修改自己密码</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Channel'" & IsOtherChecked(PO, "Channel") & ">网站频道管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Collection'" & IsOtherChecked(PO, "Collection") & ">采集管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Message'" & IsOtherChecked(PO, "Message") & ">短消息管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='MailList'" & IsOtherChecked(PO, "MailList") & ">邮件列表管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AD'" & IsOtherChecked(PO, "AD") & ">网站广告管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='FriendSite'" & IsOtherChecked(PO, "FriendSite") & ">友情链接管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Announce'" & IsOtherChecked(PO, "Announce") & ">网站公告管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Vote'" & IsOtherChecked(PO, "Vote") & ">网站调查管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Counter'" & IsOtherChecked(PO, "Counter") & ">网站统计管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Skin'" & IsOtherChecked(PO, "Skin") & ">网站风格管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Template'" & IsOtherChecked(PO, "Template") & ">通用模板管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Label'" & IsOtherChecked(PO, "Label") & ">自定义标签管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowPage'" & IsOtherChecked(PO, "ShowPage") & ">自定义页面管理</td>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Cache'" & IsOtherChecked(PO, "Cache") & ">网站缓存管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='KeyLink'" & IsOtherChecked(PO, "KeyLink") & ">站内链接管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Rtext'" & IsOtherChecked(PO, "Rtext") & ">字符过滤管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='UserGroup'" & IsOtherChecked(PO, "UserGroup") & ">会员组管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"	
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Card'" & IsOtherChecked(PO, "Card") & ">充值卡管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Equipment'" & IsOtherChecked(PO, "Equipment") & ">室场登记管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='InfoManage'" & IsOtherChecked(PO, "InfoManage") & ">学生信息管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ScoreManage'" & IsOtherChecked(PO, "ScoreManage") & ">学生成绩管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='TestManage'" & IsOtherChecked(PO, "TestManage") & ">考试管理</td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>会员管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_View'" & IsOtherChecked(PO, "User_View") & ">查看会员信息</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_ModifyInfo'" & IsOtherChecked(PO, "User_ModifyInfo") & ">修改会员信息</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_MofidyPurview'" & IsOtherChecked(PO, "User_MofidyPurview") & ">修改会员权限</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Lock'" & IsOtherChecked(PO, "User_Lock") & ">锁住/解锁会员</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Del'" & IsOtherChecked(PO, "User_Del") & ">删除会员</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Update'" & IsOtherChecked(PO, "User_Update") & ">升级为客户</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Money'" & IsOtherChecked(PO, "User_Money") & ">会员资金管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Point'" & IsOtherChecked(PO, "User_Point") & ">会员" & PointName & "管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='User_Valid'" & IsOtherChecked(PO, "User_Valid") & ">会员有效期管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ConsumeLog'" & IsOtherChecked(PO, "ConsumeLog") & ">会员消费明细</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='RechargeLog'" & IsOtherChecked(PO, "RechargeLog") & ">会员有效期明细</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>商城日常操作管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_View'" & IsOtherChecked(PO, "Order_View") & ">查看订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Confirm'" & IsOtherChecked(PO, "Order_Confirm") & ">确认订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Modify'" & IsOtherChecked(PO, "Order_Modify") & ">修改订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Del'" & IsOtherChecked(PO, "Order_Del") & ">删除订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Payment'" & IsOtherChecked(PO, "Order_Payment") & ">收款处理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Invoice'" & IsOtherChecked(PO, "Order_Invoice") & ">开发票</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Deliver'" & IsOtherChecked(PO, "Order_Deliver") & ">订单配送（实物）</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Download'" & IsOtherChecked(PO, "Order_Download") & ">订单配送（软件）</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SendCard'" & IsOtherChecked(PO, "Order_SendCard") & ">订单配送（点卡）</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_End'" & IsOtherChecked(PO, "Order_End") & ">结清订单</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Transfer'" & IsOtherChecked(PO, "Order_Transfer") & ">订单过户</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Print'" & IsOtherChecked(PO, "Order_Print") & ">订单打印</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Count'" & IsOtherChecked(PO, "Order_Count") & ">订单统计</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_OrderItem'" & IsOtherChecked(PO, "Order_OrderItem") & ">销售明细情况</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_SaleCount'" & IsOtherChecked(PO, "Order_SaleCount") & ">销售统计/排行</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Payment'" & IsOtherChecked(PO, "Payment") & ">在线支付管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bankroll'" & IsOtherChecked(PO, "Bankroll") & ">资金明细查询</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Deliver'" & IsOtherChecked(PO, "Deliver") & ">发退货记录</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Transfer'" & IsOtherChecked(PO, "Transfer") & ">订单过户记录</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PresentProject'" & IsOtherChecked(PO, "PresentProject") & ">促销方案管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='PaymentType'" & IsOtherChecked(PO, "PaymentType") & ">付款方式管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='DeliverType'" & IsOtherChecked(PO, "DeliverType") & ">送货方式管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Bank'" & IsOtherChecked(PO, "Bank") & ">银行帐户管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Order_Refund'" & IsOtherChecked(PO, "Order_Refund") & ">退款处理</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShoppingCart'" & IsOtherChecked(PO, "ShoppingCart") & ">购物车管理</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddPayment'" & IsOtherChecked(PO, "AddPayment") & ">虚拟货币支付</td>"
    Response.Write "            <td width='16%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AgentPayment'" & IsOtherChecked(PO, "AgentPayment") & ">代理商余额支付</td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "            <td width='16%'></td>"
    Response.Write "          </tr>"

    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>客户关系管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_View'" & IsOtherChecked(PO, "Client_View") & ">查看客户信息</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Add'" & IsOtherChecked(PO, "Client_Add") & ">添加客户</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyOwn'" & IsOtherChecked(PO, "Client_ModifyOwn") & ">修改属于自己的客户信息</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_ModifyAll'" & IsOtherChecked(PO, "Client_ModifyAll") & ">修改所有客户信息</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Client_Del'" & IsOtherChecked(PO, "Client_Del") & ">删除客户</td>"
    Response.Write "          </tr>"
    
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_View'" & IsOtherChecked(PO, "Service_View") & ">查看服务记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Add'" & IsOtherChecked(PO, "Service_Add") & ">添加服务记录</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyOwn'" & IsOtherChecked(PO, "Service_ModifyOwn") & ">修改自己添加的服务记录</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_ModifyAll'" & IsOtherChecked(PO, "Service_ModifyAll") & ">修改所有服务记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Service_Del'" & IsOtherChecked(PO, "Service_Del") & ">删除服务记录</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_View'" & IsOtherChecked(PO, "Complain_View") & ">查看投诉记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Add'" & IsOtherChecked(PO, "Complain_Add") & ">添加投诉记录</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyOwn'" & IsOtherChecked(PO, "Complain_ModifyOwn") & ">修改自己添加的投诉记录</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_ModifyAll'" & IsOtherChecked(PO, "Complain_ModifyAll") & ">修改所有投诉记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Complain_Del'" & IsOtherChecked(PO, "Complain_Del") & ">删除投诉记录</td>"
    Response.Write "          </tr>"

    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_View'" & IsOtherChecked(PO, "Call_View") & ">查看回访记录</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Add'" & IsOtherChecked(PO, "Call_Add") & ">添加回访记录</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyOwn'" & IsOtherChecked(PO, "Call_ModifyOwn") & ">修改自己添加的回访记录</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_ModifyAll'" & IsOtherChecked(PO, "Call_ModifyAll") & ">修改所有回访记录</td>"
    'Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='Call_Del'>删除投诉记录</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>手机短信管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToMember'" & IsOtherChecked(PO, "SendSMSToMember") & ">发送给会员</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToContacter'" & IsOtherChecked(PO, "SendSMSToContacter") & ">发送给联系人</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToConsignee'" & IsOtherChecked(PO, "SendSMSToConsignee") & ">发送给订单中的收货人</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SendSMSToOther'" & IsOtherChecked(PO, "SendSMSToOther") & ">发送给其他人</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewMessageLog'" & IsOtherChecked(PO, "ViewMessageLog") & ">查看发送结果</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='SMS_MessageReceive'" & IsOtherChecked(PO, "SMS_MessageReceive") & ">查看接收到的短信</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "        <table width='100%' border='0' cellspacing='1' cellpadding='2'>"
    Response.Write "          <tr><td colspan='6'>&nbsp;</td></tr>"
    Response.Write "          <tr><td colspan='6'>问卷调查管理权限</td></tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewSurvey'" & IsOtherChecked(PO, "ViewSurvey") & ">查看问卷</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='AddSurvey'" & IsOtherChecked(PO, "AddSurvey") & ">创建问卷</td>"
    Response.Write "            <td width='26%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurvey'" & IsOtherChecked(PO, "ManageSurvey") & ">管理问卷（修改、删除）</td>"
    Response.Write "            <td width='20%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ShowSurveyCountData'" & IsOtherChecked(PO, "ShowSurveyCountData") & ">查看调查结果</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ManageSurveyTemplate'" & IsOtherChecked(PO, "ManageSurveyTemplate") & ">问卷模板管理</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ImportSurveyQuestion'" & IsOtherChecked(PO, "ImportSurveyQuestion") & ">问卷题目导入</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ExportSurveyQuestion'" & IsOtherChecked(PO, "ExportSurveyQuestion") & ">问卷题目导出</td>"
    Response.Write "            <td width='18%'><input name='AdminPurview_Others' type='checkbox' id='AdminPurview_Others' value='ViewListQuestion'" & IsOtherChecked(PO, "ViewListQuestion") & ">查看问卷题目列表</td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "      </fieldset></td></tr>"
    Response.Write "  </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModifyPurview'>"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "    <input  type='submit' name='Submit' value='保存修改结果' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Admin.asp'"" style='cursor:hand;'></td>"
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

    '验证安全码
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>非法提交的数据!</li>"
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
        ErrMsg = ErrMsg & "<li>管理员名不能为空！</li>"
    Else
        If CheckBadChar(strAdminName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>管理员名中含有非法字符！</li>"
        End If
    End If
    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>前台会员名不能为空！</li>"
    Else
        If CheckBadChar(UserName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>前台会员名中含有非法字符！</li>"
        End If
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>初始密码不能为空！</li>"
    End If
    If PwdConfirm <> Password Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>确认密码必须与初始密码相同！</li>"
    End If
    If CheckBadChar(Password) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>初始密码中含有非法字符！</li>"
    End If
    If Purview = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>管理员权限不能为空！</li>"
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
        ErrMsg = ErrMsg & "<li>找不到指定的前台会员！</li>"
    End If
    Set rsUser = Nothing
    If FoundErr = True Then Exit Sub
    
    sqlAdmin = "Select * from PE_Admin where AdminName='" & strAdminName & "'"
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If Not (rsAdmin.BOF And rsAdmin.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库中已经存在此管理员！</li>"
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
    Call WriteEntry(1, AdminName, "新增管理员：" & strAdminName)
    Call main
End Sub

Sub SaveModifyPwd()
    Dim UserID, UserName, Password, PwdConfirm, EnableMultiLogin
    Dim rsAdmin, sqlAdmin
    
    '验证安全码
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>非法提交的数据!</li>"
    End If

    UserID = Trim(Request("ID"))
    UserName = Trim(Request("UserName"))
    Password = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    EnableMultiLogin = Trim(Request("EnableMultiLogin"))
    If UserID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的管理员ID</li>"
    Else
        UserID = PE_CLng(UserID)
    End If
    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>前台会员名不能为空！</li>"
    Else
        If CheckBadChar(UserName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>前台会员名中含有非法字符！</li>"
        End If
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>新密码不能为空！</li>"
    End If
    If PwdConfirm <> Password Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>确认密码必须与新密码相同！</li>"
    End If
    If CheckBadChar(Password) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>新密码中含有非法字符！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    Dim rsUser
    Set rsUser = Conn.Execute("Select * from PE_User where UserName='" & UserName & "'")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的前台会员！</li>"
    End If
    Set rsUser = Nothing
    If FoundErr = True Then Exit Sub
    
    sqlAdmin = "Select * from PE_Admin where ID=" & UserID
    Set rsAdmin = Server.CreateObject("Adodb.RecordSet")
    rsAdmin.Open sqlAdmin, Conn, 1, 3
    If rsAdmin.BOF And rsAdmin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不存在此管理员！</li>"
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
    Call WriteEntry(1, AdminName, "修改管理员密码，ID：" & UserID)

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

    '验证安全码
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>非法提交的数据!</li>"
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
        ErrMsg = ErrMsg & "<li>请指定要修改的管理员ID</li>"
    Else
        UserID = PE_CLng(UserID)
    End If
    If Purview = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>管理员权限不能为空！</li>"
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
        ErrMsg = ErrMsg & "<li>不存在此管理员！</li>"
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
    Call WriteEntry(1, AdminName, "修改管理员权限，ID：" & UserID)

    Call main
End Sub

Sub DelAdmin()
    Dim UserID
    Dim rsAdmin, sqlAdmin
    Dim rsChannel, sqlChannel

    '验证安全码
    If CheckSecretCode(Trim(Request.Form("Scode"))) <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>非法提交的数据!</li>"
    End If

    UserID = Trim(Request("ID"))
    If IsValidID(UserID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的管理员ID</li>"
        Exit Sub
    End If
    If InStr(UserID, ",") > 0 Then
        Conn.Execute ("delete from PE_Admin where ID in (" & UserID & ")")
    Else
        Conn.Execute ("delete from PE_Admin where ID=" & UserID & "")
    End If
    Call WriteEntry(1, AdminName, "删除管理员，ID：" & UserID)

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
    Response.Write "input{FONT-FAMILY:宋体;FONT-SIZE: 9pt;}" & vbCrLf
    Response.Write ".cssPWD{background-color:#EBEBEB;border-right:solid 1px #BEBEBE;border-bottom:solid 1px #BEBEBE;}" & vbCrLf
    Response.Write ".cssWeak{background-color:#FF4545;border-right:solid 1px #BB2B2B;border-bottom:solid 1px #BB2B2B;}" & vbCrLf
    Response.Write ".cssMedium{background-color:#FFD35E;border-right:solid 1px #E9AE10;border-bottom:solid 1px #E9AE10;}" & vbCrLf
    Response.Write ".cssStrong{background-color:#3ABB1C;border-right:solid 1px #267A12;border-bottom:solid 1px #267A12;}" & vbCrLf
    Response.Write ".cssPWT{width:132px;}" & vbCrLf
    Response.Write "</style>" & vbCrLf
    Response.Write "<table cellpadding='0' cellspacing='0' class='cssPWT' style='height:16px'><tr valign='bottom'><td id='idSM1' width='33%' class='cssPWD' align='center'><span style='font-size:1px'>&nbsp;</span><span id='idSMT1' style='display:none;'>弱</span></td><td id='idSM2' width='34%' class='cssPWD' align='center' style='border-left:solid 1px #fff'><span style='font-size:1px'>&nbsp;</span><span id='idSMT0' style='display:inline;font-weight:normal;color:#666'>无</span><span id='idSMT2' style='display:none;'>中</span></td><td id='idSM3' width='33%' class='cssPWD' align='center' style='border-left:solid 1px #fff'><span style='font-size:1px'>&nbsp;</span><span id='idSMT3' style='display:none;'>强</span></td></tr></table>"
End Sub

Sub ShowPurviewTips()
    Response.Write "<b><font color='#FF0000'>频道权限说明：</font></b>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td width='150'><b>操作项目</b></td>" & vbCrLf
    Response.Write "    <td><b>频道管理员权限</b></td>" & vbCrLf
    Response.Write "    <td><b>栏目总编权限</b></td>" & vbCrLf
    Response.Write "    <td><b>栏目管理员权限</b></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>添加信息</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以在所有栏目添加信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以在所有栏目添加信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>只能在有录入权限的栏目添加信息</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>修改信息</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以修改所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以修改所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>只能修改有管理权限的栏目中的信息，或者自己添加的信息</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>删除信息</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以删除所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以删除所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>只能在有管理权限的栏目中删除信息，或者自己添加的信息</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>审核信息</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以审核所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以审核所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>只能在有审核权限的栏目中移动信息</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>管理信息（固顶、推荐）</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以管理所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以管理所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>只能在有管理权限的栏目中管理信息</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>生成HTML（批量）</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以生成所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>可以生成所有栏目的信息</font></td>" & vbCrLf
    Response.Write "    <td><font color='#0000FF'>只能在有管理权限的栏目中生成信息</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>生成HTML（自动）</b></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>保存信息时自动生成HTML</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>保存信息时自动生成HTML</font></td>" & vbCrLf
    Response.Write "    <td><font color='#009900'>保存信息时自动生成HTML</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>移动信息</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>批量设置属性</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>专题信息管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>评论管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>回收站管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>栏目管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>专题管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>上传文件管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#009900'>可以</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>模板管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>JS文件管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>顶部菜单管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>关键字管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>作者管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>来源管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td><b>更新栏目XML数据</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>自定义字段</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='150'><b>广告管理</b></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#0000FF'>另行指定</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "    <td align='center'><font color='#FF0000'>不能</font></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

%>
