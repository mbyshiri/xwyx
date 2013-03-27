<!--#include file="Admin_Common.asp"-->
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

Dim ManageType

Response.write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'><title>设置栏目权限</title></head>"
Response.write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
Response.write "<form name='myform' method='post' action=''>"

ManageType = Trim(Request("ManageType"))

If Action = "Modify" Then
    Select Case ManageType
    Case "Admin"
        UserID = PE_CLng(Trim(Request("UserID")))
        If UserID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的会员ID</li>"
        Else
            Dim sql, rs
            sql = "Select * from PE_Admin where ID=" & UserID
            Set rs = Server.CreateObject("Adodb.RecordSet")
            rs.Open sql, Conn, 1, 3
            If rs.BOF And rs.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>不存在此会员！</li>"
            Else
                AdminName = rs("UserName")
                AdminPurview = rs("Purview")
                AdminPurview_Channel = rs("AdminPurview_" & ChannelDir)
                arrClass_View = rs("arrClass_View")
                arrClass_Input = rs("arrClass_Input")
                arrClass_Check = rs("arrClass_Check")
                arrClass_Manage = rs("arrClass_Manage")
                If ChannelID = 4 Then
                    Dim arrKind_GuestBook
                    arrKind_GuestBook = Split(rs("arrClass_GuestBook"), "|||")
                End If
                If ChannelID = 998 Then
                    If IsNull(rs("arrClass_House")) Then
                        ReDim arrKind_House(3)
                    Else
                        arrKind_House = Split(rs("arrClass_House"), "|||")
                    End If
                End If
            End If
            rs.Close
            Set rs = Nothing
        End If
    Case "Group"
        GroupID = PE_CLng(Trim(Request("GroupID")))
        If GroupID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的会员组ID</li>"
        Else
            Dim rsGroup, sqlGroup
            sqlGroup = "Select arrClass_Browse,arrClass_View,arrClass_Input from PE_UserGroup where GroupID=" & GroupID
            Set rsGroup = Server.CreateObject("Adodb.RecordSet")
            rsGroup.Open sqlGroup, Conn, 1, 1
            If rsGroup.BOF And rsGroup.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>不存在此会员组！</li>"
            Else
                arrClass_Browse = rsGroup(0)
                arrClass_View = rsGroup(1)
                arrClass_Input = rsGroup(2)
            End If
            rsGroup.Close
            Set rsGroup = Nothing
        End If
    Case "User"
        UserID = PE_CLng(Trim(Request("UserID")))
        If UserID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的会员ID</li>"
        Else
            Dim rsUser, sqlUser
            sqlUser = "Select arrClass_Browse,arrClass_View,arrClass_Input from PE_User where UserID=" & UserID
            Set rsUser = Server.CreateObject("Adodb.RecordSet")
            rsUser.Open sqlUser, Conn, 1, 1
            If rsUser.BOF And rsUser.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>不存在此会员组！</li>"
            Else
                arrClass_Browse = rsUser(0)
                arrClass_View = rsUser(1)
                arrClass_Input = rsUser(2)
            End If
            rsUser.Close
            Set rsUser = Nothing
        End If
    
    End Select
End If
If FoundErr = True Then
    Response.write ErrMsg
    Response.End
End If

Response.write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
If ManageType = "Admin" Then
    If ChannelID = 4 Then
        Response.write "  <tr align='center' class='title' height='22'>"
        Response.write "    <td><strong>栏目名称</strong></td>"
        Response.write "    <td width='30'><strong>修改</strong></td>"
        Response.write "    <td width='30'><strong>删除</strong></td>"
        Response.write "    <td width='30'><strong>移动</strong></td>"
        Response.write "    <td width='30'><strong>审核</strong></td>"
        Response.write "    <td width='30'><strong>精华</strong></td>"
        Response.write "    <td width='30'><strong>固顶</strong></td>"
        Response.write "    <td width='30' height='22'><strong>回复</strong></td>"
        Response.write "  </tr>"
        Response.write "  <tr style='display:none'>"
        Response.write "    <td width='30' align='center'><input name='Purview_Modify' type='checkbox' id='Purview_Modify' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Del' type='checkbox' id='Purview_Del' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Move' type='checkbox' id='Purview_Move' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Quintessence' type='checkbox' id='Purview_Quintessence' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_SetOnTop' type='checkbox' id='Purview_SetOnTop' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Reply' type='checkbox' id='Purview_Reply' value='0'></td>"
        Response.write "  </tr>"
    ElseIf ChannelID = 999 Then
        Response.write "  <tr align='center' class='title' height='22'>"
        Response.write "    <td><strong>栏目名称</strong></td>"
        'Response.Write "    <td width='30'><strong>查看</strong></td>"
        Response.write "    <td width='30'><strong>录入</strong></td>"
        Response.write "    <td width='30'><strong>审核</strong></td>"
        Response.write "    <td width='30' height='22'><strong>管理</strong></td>"
        Response.write "  </tr>"
        Response.write "  <tr style='display:none'>"
        'Response.Write "    <td width='30' align='center'></td>"
        Response.write "<input name='Purview_View' type='Hidden' id='Purview_View' value='0'>"
        Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Manage' type='checkbox' id='Purview_Manage' value='0'></td>"
        Response.write "  </tr>"
    Else
        Response.write "  <tr align='center' class='title' height='22'>"
        Response.write "    <td><strong>栏目名称</strong></td>"
        Response.write "    <td width='30'><strong>查看</strong></td>"
        Response.write "    <td width='30'><strong>录入</strong></td>"
        Response.write "    <td width='30'><strong>审核</strong></td>"
        Response.write "    <td width='30' height='22'><strong>管理</strong></td>"
        Response.write "  </tr>"
        Response.write "  <tr style='display:none'>"
        Response.write "    <td width='30' align='center'><input name='Purview_View' type='checkbox' id='Purview_View' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='0'></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Manage' type='checkbox' id='Purview_Manage' value='0'></td>"
        Response.write "  </tr>"
    End If
Else
    Response.write "  <tr align='center' class='title' height='22'>"
    Response.write "    <td><strong>栏目名称</strong></td>"
    Response.write "    <td width='30'><strong>浏览</strong></td>"
    Response.write "    <td width='30'><strong>查看</strong></td>"
    Response.write "    <td width='30'><strong>发布</strong></td>"
    Response.write "  </tr>"
    Response.write "  <tr class='tdbg'>"
    Response.write "        <td><img src='../Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'><b><font color='#FF0000'>整个频道</font><b></td>"
    Response.write "    <td width='30' align='center'><input name='Purview_Browse' type='checkbox' id='Purview_Browse' value='" & ChannelDir & "all'"
    If (Action = "Modify" And FoundInArr(arrClass_Browse, ChannelDir & "all", ",")) Then Response.write " checked"
    Response.write " onclick=""if (this.checked==true){for(var i=1;i<document.myform.Purview_Browse.length;i++){document.myform.Purview_Browse[i].checked=false;}}"""
    Response.write "></td>"
    Response.write "    <td width='30' align='center'><input name='Purview_View' type='checkbox' id='Purview_View' value='" & ChannelDir & "all'"
    If (Action = "Modify" And FoundInArr(arrClass_View, ChannelDir & "all", ",")) Then Response.write " checked"
    Response.write " onclick=""if (this.checked==true){for(var i=1;i<document.myform.Purview_View.length;i++){document.myform.Purview_View[i].checked=false;}}"""
    Response.write "></td>"
    Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='" & ChannelDir & "all'"
    If (Action = "Modify" And FoundInArr(arrClass_Input, ChannelDir & "all", ",")) Then Response.write " checked"
    Response.write " onclick=""if (this.checked==true){for(var i=1;i<document.myform.Purview_Input.length;i++){document.myform.Purview_Input[i].checked=false;}}"""
    Response.write "></td>"
    Response.write "  </tr>"
    Response.write "  <tr style='display:none'>"
    Response.write "    <td width='30' align='center'><input name='Purview_Browse' type='checkbox' id='Purview_Browse' value='0' disabled></td>"
    Response.write "    <td width='30' align='center'><input name='Purview_View' type='checkbox' id='Purview_View' value='0' disabled></td>"
    Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='0' disabled></td>"
    Response.write "  </tr>"
End If

If ManageType = "Admin" And ChannelID = 4 Then
    Dim rsGuestKind
    Set rsGuestKind = Conn.Execute("select * from PE_GuestKind order by OrderID,KindID")
    Do While Not rsGuestKind.EOF
        Response.write "  <tr class='tdbg'>"
        Response.write "    <td width='130' align='center'>" & rsGuestKind("KindName") & "</td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Modify' type='checkbox' id='Purview_Modify' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(0), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Del' type='checkbox' id='Purview_Del' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(1), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Move' type='checkbox' id='Purview_Move' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(2), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(3), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Quintessence' type='checkbox' id='Purview_Quintessence' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(4), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_SetOnTop' type='checkbox' id='Purview_SetOnTop' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(5), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "    <td width='30' align='center'><input name='Purview_Reply' type='checkbox' id='Purview_Reply' value='" & rsGuestKind("KindID") & "'"
        If Action = "Modify" And AdminPurview_Channel = 3 Then
            If FoundInArr(arrKind_GuestBook(6), rsGuestKind("KindID"), ",") = True Then Response.write " checked"
        End If
        Response.write "></td>"
        Response.write "  </tr>"
        rsGuestKind.MoveNext
    Loop
    Set rsGuestKind = Nothing
ElseIf ChannelID = 998 Then
    Dim rsHouseClass, sqlHouseClass
    sqlHouseClass = "select * from PE_HouseConfig order by ClassID"
    Set rsHouseClass = Conn.Execute(sqlHouseClass)
    Do While Not rsHouseClass.EOF
        Response.write "  <tr class='tdbg'>"
        Response.write "    <td>"
        Response.write "&nbsp;&nbsp;<b>"
        Response.write rsHouseClass("ClassName")
        Response.write "    </td>"
        If ManageType = "Admin" Then
            Response.write "    <td width='30' align='center'><input name='Purview_View' type='checkbox' id='Purview_View' value='" & rsHouseClass("ClassID") & "'"
            If Action = "Modify" And AdminPurview_Channel = 3 Then
                If FoundInArr(arrKind_House(0), rsHouseClass("ClassID"), ",") = True Then Response.write " checked"
            End If
            Response.write "></td>"
            Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='" & rsHouseClass("ClassID") & "'"
            If Action = "Modify" And AdminPurview_Channel = 3 Then
                If FoundInArr(arrKind_House(1), rsHouseClass("ClassID"), ",") = True Then Response.write " checked"
            End If
            Response.write "></td>"
            Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='" & rsHouseClass("ClassID") & "'"
            If Action = "Modify" And AdminPurview_Channel = 3 Then
                If FoundInArr(arrKind_House(2), rsHouseClass("ClassID"), ",") = True Then Response.write " checked"
            End If
            Response.write "></td>"
            Response.write "    <td width='30' align='center'><input name='Purview_Manage' type='checkbox' id='Purview_Manage' value='" & rsHouseClass("ClassID") & "'"
            If Action = "Modify" And AdminPurview_Channel = 3 Then
                If FoundInArr(arrKind_House(3), rsHouseClass("ClassID"), ",") = True Then Response.write " checked"
            End If
            Response.write "></td>"
        End If
        Response.write "  </tr>"
        rsHouseClass.MoveNext
    Loop
    Set rsHouseClass = Nothing
Else
    Dim arrShowLine(20), i
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    Dim sqlClass, rsClass, iDepth
    sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    Do While Not rsClass.EOF
        Response.write "  <tr class='tdbg'>"
        Response.write "    <td>"
        iDepth = rsClass("Depth")
        If rsClass("NextID") > 0 Then
            arrShowLine(iDepth) = True
        Else
            arrShowLine(iDepth) = False
        End If
        If iDepth > 0 Then
            For i = 1 To iDepth
                If i = iDepth Then
                    If rsClass("NextID") > 0 Then
                        Response.write "<img src='../images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
                    Else
                        Response.write "<img src='../images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
                    End If
                Else
                    If arrShowLine(i) = True Then
                        Response.write "<img src='../images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
                    Else
                        Response.write "<img src='../images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
                    End If
                End If
            Next
          End If
        If rsClass("Child") > 0 Then
            Response.write "<img src='../Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
        Else
            Response.write "<img src='../Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
        End If
        If rsClass("Depth") = 0 Then
            Response.write "<b>"
        End If
        Response.write rsClass("ClassName")
        Response.write "    </td>"
        If ManageType = "Admin" Then
            If ChannelID = 999 Then
                Response.write "<input name='Purview_View' type='Hidden' id='Purview_View' value='" & rsClass("ClassID") & "'>"
                Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_Input, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
                Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_Check, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
                Response.write "    <td width='30' align='center'><input name='Purview_Manage' type='checkbox' id='Purview_Manage' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_Manage, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
            Else
                Response.write "    <td width='30' align='center'><input name='Purview_View' type='checkbox' id='Purview_View' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_View, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
                Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_Input, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
                Response.write "    <td width='30' align='center'><input name='Purview_Check' type='checkbox' id='Purview_Check' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_Check, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
                Response.write "    <td width='30' align='center'><input name='Purview_Manage' type='checkbox' id='Purview_Manage' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And AdminPurview_Channel = 3 Then
                    If FoundInArr(arrClass_Manage, rsClass("ClassID"), ",") = True Then Response.write " checked"
                End If
                Response.write "></td>"
            End If
        Else
            Response.write "    <td width='30' align='center'>"
            If rsClass("ClassPurview") < 2 Then
                Response.write "<input name='Purview_Browse' type='checkbox' id='Purview' value='' checked disabled"
            Else
                Response.write "<input name='Purview_Browse' type='checkbox' id='Purview_Browse' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And FoundInArr(arrClass_Browse, rsClass("ClassID"), ",") Then Response.write " checked"
            End If
            Response.write " onclick='document.myform.Purview_Browse[0].checked=false;'></td>"
            Response.write "    <td width='30' align='center'>"
            If rsClass("ClassPurview") = 0 Then
                Response.write "<input name='Purview_View' type='checkbox' id='Purview' value='' checked disabled"
            Else
                Response.write "<input name='Purview_View' type='checkbox' id='Purview_View' value='" & rsClass("ClassID") & "'"
                If Action = "Modify" And FoundInArr(arrClass_View, rsClass("ClassID"), ",") Then Response.write " checked"
            End If
            Response.write " onclick='document.myform.Purview_View[0].checked=false;'></td>"
            Response.write "    <td width='30' align='center'><input name='Purview_Input' type='checkbox' id='Purview_Input' value='" & rsClass("ClassID") & "'"
            If Action = "Modify" And FoundInArr(arrClass_Input, rsClass("ClassID"), ",") Then Response.write " checked"
                Response.write " onclick='document.myform.Purview_Input[0].checked=false;'></td>"
            End If
        Response.write "  </tr>"
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
End If
Response.write "</form>"
Response.write "</body></html>"
Call CloseConn
%>
