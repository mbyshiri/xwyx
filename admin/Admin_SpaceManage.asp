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

Dim KindFileName, ComUrl, SelectedName, KindID, oldKInd


XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

KindID = PE_CLng(Trim(Request("KindID")))

Select Case Action
Case "Add"
    SelectedName = "新增聚合空间"
Case "Modify"
    SelectedName = "修改聚合空间"
Case "Check"
    SelectedName = "审核聚合空间"
Case "Kind", "OrderKind"
    SelectedName = "聚合空间类别"
Case "AddKind"
    SelectedName = "新增空间分类"
Case "ModifyKing"
    SelectedName = "修改空间分类"
Case Else
    SelectedName = "聚合空间管理"
End Select

UserID = PE_CLng(Trim(Request("UserID")))

strFileName = "Admin_SpaceManage.asp?Action=" & Action
If UserID > 0 Then strFileName = strFileName & "&UserID=" & UserID
KindFileName = strFileName
ComUrl = "Admin_SpaceManage.asp?UserID=" & UserID

If KindID > 0 Then
    strFileName = strFileName & "&KindID=" & KindID
    ComUrl = ComUrl & "&KindID=" & KindID
End If
If Keyword <> "" Then
    strFileName = strFileName & "&Field=" & strField & "&keyword=" & Keyword
    KindFileName = KindFileName & "&Field=" & strField & "&keyword=" & Keyword
    ComUrl = ComUrl & "&Field=" & strField & "&keyword=" & Keyword
End If


If CurrentPage > 1 Then ComUrl = ComUrl & "&page=" & CurrentPage

Response.Write "<html><head><title>" & SelectedName & "</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
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

Response.Write "function CheckInput(){" & vbCrLf
Response.Write "  if(document.myform.BlogName.value==''){" & vbCrLf
Response.Write "      alert('聚合空间名不能为空！');" & vbCrLf
Response.Write "      document.myform.BlogName.focus();" & vbCrLf
Response.Write "      return false;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "   document.myform.Intro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "function changemode(){" & vbCrLf
Response.Write "    var dbname=document.myform.addtype.value;" & vbCrLf
Response.Write "    if(dbname=='2'){" & vbCrLf
Response.Write "        url.style.display='';" & vbCrLf
Response.Write "    }else{" & vbCrLf
Response.Write "        url.style.display='none';" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf

Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle(SelectedName, 10048)

Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='80' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td height='30'><a href='Admin_SpaceManage.asp'>空间管理首页</a>&nbsp;|&nbsp;<a href='Admin_SpaceManage.asp?Action=Check'>审核空间</a>&nbsp;|&nbsp;<a href='Admin_SpaceManage.asp?Action=Kind'>空间分类管理</a>&nbsp;|&nbsp;<a href='Admin_SpaceManage.asp?Action=AddKind'>添加空间分类</a></td>" & vbCrLf

Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd"
    Call SaveAdd
Case "SaveModify"
    Call SaveModify
Case "Dis"
    Call SetStat(1)
Case "En"
    Call SetStat(2)
Case "DisElite"
    Call SetStat(3)
Case "EnElite"
    Call SetStat(4)
Case "DisTop"
    Call SetStat(5)
Case "EnTop"
    Call SetStat(6)
Case "Del"
    Call Del
Case "Kind"
    Call Kind
Case "AddKind"
    Call AddKind
Case "ModifyKind"
    Call ModifyKind
Case "SaveAddKind", "SaveModifyKind"
    Call SaveKind
Case "DelKind", "ClearKind"
    Call DelKind
Case "OrderKind"
    Call OrderKind
Case Else
    Call main
    End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsBlog, sqlBlog, rsUser, tempname
    Dim iCount

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'>" & GetKindList(KindID) & "</td></tr></table><br>"
    If UserID > 0 Then
        tempname = "模块"
    Else
        tempname = "空间"
    End If
    Response.Write "  <form name='myform' method='Post' action='Admin_SpaceManage.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title' height='22'>"
    Response.Write "    <td width='30'><strong>选中</strong></td>"
    Response.Write "    <td width='35'><strong>序号</strong></td>"
    Response.Write "    <td width='80'><strong>分类</strong></td>"
    Response.Write "    <td><strong>名称</strong></td>"
    Response.Write "    <td width='35'><strong>点击</strong></td>"
    Response.Write "    <td width='70'><strong>状态</strong></td>"
    Response.Write "    <td width='200'><strong>操 作</strong></td>"
    If UserID > 0 Then
        Response.Write "<td width='80'><strong>排 序</strong></td></tr>"
    Else
        Response.Write "<td width='80'><strong>所属会员</strong></td></tr>"
    End If
    
    Set rsBlog = Server.CreateObject("Adodb.RecordSet")
    If UserID > 0 Then
        sqlBlog = "select * from PE_Space Where Type>0"
    Else
        sqlBlog = "select * from PE_Space Where Type=1"
    End If
    If KindID > 0 Then sqlBlog = sqlBlog & " and ClassID=" & KindID
    If Keyword <> "" Then
        Select Case strField
        Case "name"
            sqlBlog = sqlBlog & " and Name like '%" & Keyword & "%' "
        Case "address"
            sqlBlog = sqlBlog & " and Address like '%" & Keyword & "%' "
        Case "Phone"
            sqlBlog = sqlBlog & " and Tel like '%" & Keyword & "%' "
        Case "intro"
            sqlBlog = sqlBlog & " and Intro like '%" & Keyword & "%' "
        Case Else
            sqlBlog = sqlBlog & " and Name like '%" & Keyword & "%' "
        End Select
    End If
    If UserID > 0 Then sqlBlog = sqlBlog & " and UserID = " & UserID
    If Action = "Check" Then
        sqlBlog = sqlBlog & " and Passed = " & PE_False
    Else
        If UserID = 0 Then sqlBlog = sqlBlog & " and Passed = " & PE_True
    End If
    If UserID > 0 Then
        sqlBlog = sqlBlog & " order by Type,onTop " & PE_OrderType & ",OrderID"
    Else
        sqlBlog = sqlBlog & " order by ID Desc"
    End If
    rsBlog.Open sqlBlog, Conn, 1, 1
    If rsBlog.BOF And rsBlog.EOF Then
        rsBlog.Close
        Set rsBlog = Nothing
        If UserID > 0 Then
            Response.Write "  <tr class='tdbg'><td colspan='10' align='center'><br>该用户尚未添加任何" & tempname & "！<br><br></td></tr>"
        Else
            If Action = "Check" Then
                Response.Write "  <tr class='tdbg'><td colspan='10' align='center'><br>没有任何申请中的" & tempname & "！<br><br></td></tr>"
            Else
                Response.Write "  <tr class='tdbg'><td colspan='10' align='center'><br>没有任何" & tempname & "！<br><br></td></tr>"
            End If
        End If
        Response.Write "</Table>"
        Exit Sub
    End If
    
    totalPut = rsBlog.RecordCount
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
            rsBlog.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    Do While Not rsBlog.EOF
        If rsBlog("Type") = 1 Then
            Response.Write "  <tr align='center' bgcolor='#ffbbbb' onmouseout=""this.style.backgroundColor='#ffbbbb'"" onmouseover=""this.style.backgroundColor='#bbbbbb'"">"
        Else
            Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        End If
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsBlog("ID") & "'  onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsBlog("ID") & "</td>"
        Response.Write "    <td>" & GetKingName(rsBlog("ClassID")) & "</td>"
        Response.Write "    <td>" & GetSubStr(rsBlog("Name"), 24, False) & "</td>"
        Response.Write "    <td>" & rsBlog("Hits") & "</td><td>"
        If rsBlog("Passed") = True Then
            Response.Write "<font color=""green"">√</font>"
        Else
            Response.Write "<font color=""red"">×</font>"
        End If
        If rsBlog("onTop") = True Then
            Response.Write "&nbsp;<font color=""blue"">固</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        If rsBlog("IsElite") = True Then
            Response.Write "&nbsp;<font color=""green"">荐</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td><td>"
        Response.Write "<a href='Admin_SpaceManage.asp?Action=Modify&ID=" & rsBlog("ID") & "&UserID=" & UserID & "'>修改</a>"
        If rsBlog("Passed") = True Then
            Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=Dis&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "'>禁用</a>"
        Else
            Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=En&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "'>启用</a>"
        End If
        If rsBlog("onTop") = True Then
            Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=DisTop&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "'>解固</a>"
        Else
            Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=EnTop&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "'>固顶</a>"
        End If
        If rsBlog("IsElite") = True Then
            Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=DisElite&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "'>取消推荐</a>"
        Else
            Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=EnElite&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "'>设为推荐</a>"
        End If
        Response.Write "&nbsp;&nbsp;<a href='Admin_SpaceManage.asp?Action=Del&ID=" & rsBlog("ID") & "&UserID=" & UserID & "&KindID=" & KindID & "&page=" & CurrentPage & "&Field=" & strField & "&keyword=" & Keyword & "' onClick=""return confirm('确定要删除" & tempname & rsBlog("Name") & "吗？');"">删除</a>"
        Response.Write "</td>"
        If UserID > 0 Then
            If rsBlog("Type") < 2 Then
                Response.Write "<td>根栏目</td>"
            Else
                Response.Write "<td><input name='OrderID" & rsBlog("ID") & "' type='text' id='OrderID" & rsBlog("ID") & "' value='" & rsBlog("OrderID") & "' size='4' maxlength='4' style='text-align:center'><input type='submit' name='Submit' value='修改' onClick=""document.myform.Action.value='order|" & rsBlog("ID") & "'""></td>"
            End If
        Else
            Set rsUser = Conn.Execute("select Top 1 UserName from PE_User Where UserID=" & rsBlog("UserID"))
            Response.Write "<td><a href='Admin_User.asp?Action=Show&UserID=" & rsBlog("UserID") & "'>" & rsUser("UserName") & "</a> <a href='Admin_SpaceManage.asp?UserID=" & rsBlog("UserID") & "'>详</a></td>"
        End If
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsBlog.MoveNext
    Loop
    If UserID = 0 Then rsUser.Close: Set rsUser = Nothing
    rsBlog.Close: Set rsBlog = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> 选中本页显示的所有" & tempname & "</td>"
    Response.Write "    <td><input name='Action' type='hidden' id='Action' value='Del'>"
    Response.Write "<input name='UserID' type='hidden' id='UserID' value=" & UserID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='删除选中的" & tempname & "'>"
    If UserID > 0 Then Response.Write "&nbsp;&nbsp;<input name='add' type='button' id='add' value='增加一个" & tempname & "' onClick=""window.location.href='Admin_SpaceManage.asp?Action=Add&UserID=" & UserID & "';"" style='cursor:hand;'>"
    Response.Write "</td></tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个" & tempname, True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>" & tempname & "搜索：</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SpaceManage.asp'>"
    Response.Write "<input name='Action' type='hidden' id='Action' value='" & Action & "'>"
    Response.Write "<input name='UserID' type='hidden' id='UserID' value=" & UserID & ">"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>名称</option><option value='address'>地址</option><option value='Phone'>电话</option><option value='intro'>简介</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "关键字"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='搜索'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub Add()
    Dim rsUser, rsBlog, rsBlogClass, UName
    If UserID = 0 Then
        Call main
        Exit Sub
    End If
    Set rsUser = Conn.Execute("select Top 1 UserName,Blog from PE_User Where UserID=" & UserID)
    If rsUser.BOF And rsUser.EOF Then
        Call main
    Else
        UName = rsUser("UserName")
        Set rsBlog = Conn.Execute("select Top 1 * from PE_Space Where Type=1 and Passed=" & PE_True & " and UserID=" & UserID)
        If rsBlog.BOF And rsBlog.EOF Then
            Call PopCalendarInit
            Response.Write "<form method='post' action='Admin_SpaceManage.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
            Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>[<font color=red>" & rsUser("UserName") & "</font>] 新 增 聚 合 空 间</strong></div></td>"
            Response.Write "    </tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>空间名称：</strong><input name='BlogName' type='text' size='20' maxlength='20'> <font color='#FF0000'>*</font></td></tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'><table><tr><td>&nbsp;<strong>空间首页<br>&nbsp;显示项目：</strong></td><td>"
            Set rsBlogClass = Conn.Execute("select * from PE_Channel Where Disabled=" & PE_False & " and ModuleType>0 and ModuleType<4 order by OrderID")
            Do While Not rsBlogClass.EOF
                Response.Write "<input type='checkbox' name='Showitem' value='" & rsBlogClass("ChannelID") & "' checked>我在" & rsBlogClass("ChannelName") & "频道发表的作品<br>"
                rsBlogClass.MoveNext
            Loop
            Response.Write "</td></tr></table></td></tr><tr class='tdbg'>"
            Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong>分类：</strong><select name='BlogType'>" & GetKingOpti(0) & "</select></td>"
            Response.Write "    <td rowspan='9 align='center' valign='top' class='tdbg'>"
            Response.Write "        <table width='180' height='200' border='1'>"
            Response.Write "            <tr><td width='100%' align='center'><img id='showphoto' src='" & InstallDir & "Space/default.gif' width='150' height='172'></td></tr>"
            Response.Write "        </table>"
            Response.Write "        <input name='Photo' type='text' size='25'><strong>：照 片 地 址</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AdminBlogPic&Uname=" & UName & "' frameborder=0 scrolling=no width='285' height='25'></iframe>"
            Response.Write "     </td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>日期：</strong><input name='BirthDay' type='text' size='20' maxlength='20' value='" & FormatDateTime(Date, 2) & "' maxlength='20'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BirthDay, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>地址：</strong><input name='Address' type='text' size='20' maxlength='20'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>电话：</strong><input name='Tel' type='text' size='20' maxlength='20'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>传真：</strong><input name='Fax' type='text' size='20' maxlength='20'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>单位：</strong><input name='Company' type='text' size='20' maxlength='20'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>部门：</strong><input name='Department' type='text' size='20' maxlength='20'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>邮编：</strong><input name='ZipCode' type='text' size='20' maxlength='20'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>ＱＱ：</strong><input name='QQ' type='text'></td></tr>"
            Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>邮件：</strong><input name='Email' type='text' size='20' maxlength='20'></td><td><strong> 主页：</strong><input name='HomePage' type='text'></td></tr>"
            Response.Write "  <tr class='tdbg'> "
            Response.Write "    <td colspan='2'>&nbsp;<strong>简介</strong>↓<br>"
            Response.Write "      <textarea name='Intro' id='Intro' cols='72' rows='9' style='display:none'></textarea>"
            Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='300' ></iframe>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            Response.Write "  <tr>"
            Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
            Response.Write "    <input name='UserID' type='hidden' id='UserID' value=" & UserID & ">"
            Response.Write "    <input name='addtype' type='hidden' id='addtype' value=1>"
            Response.Write "    <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
            Response.Write "    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_SpaceManage.asp';"" style='cursor:hand;'></td>"
            Response.Write "  </tr>"
            Response.Write "</table></form>"
        Else
            Call PopCalendarInit
            Response.Write "<form method='post' action='Admin_SpaceManage.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
            Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
            Response.Write "    <tr class='title'> "
            Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>[<font color=red>" & rsUser("UserName") & "</font>] 新 增 聚 合 空 间 模 块</strong></div></td>"
            Response.Write "    </tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>模块名称：</strong><input name='BlogName' type='text' size='20' maxlength='20'> <font color='#FF0000'>*</font></td></tr>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>模块类型：</strong><select name='addtype' onChange=""changemode()""><option value=2>外部RSS数据</option><option value=3>我的日志</option><option value=4>我的音乐</option><option value=5>我的图书</option><option value=6>我的图片</option><option value=7>我的连接</option></select></td></tr>"
            Response.Write "<tbody id='url' style='display:'>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>ＲＳＳ地址：</strong><input name='LinkUrl' type='text' size='67' maxlength='100' value='http://'> <font color='#FF0000'>* RSS来源地址</font></td></tr></tbody>"
            Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>显示条数：</strong><input name='ListNum' type='text' size='2' maxlength='2' value='10'> <font color='#FF0000'>* 前台显示条数</font></td></tr>"
            Response.Write "  <tr class='tdbg'> "
            Response.Write "    <td colspan='2'>&nbsp;<strong>简介</strong>↓<br>"
            Response.Write "      <textarea name='Intro' id='Intro' cols='72' rows='9' style='display:none'></textarea>"
            Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='300' ></iframe>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            Response.Write "  <tr>"
            Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
            Response.Write "    <input name='UserID' type='hidden' id='UserID' value=" & UserID & ">"
            Response.Write "    <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
            Response.Write "    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_SpaceManage.asp';"" style='cursor:hand;'></td>"
            Response.Write "  </tr>"
            Response.Write "</table></form>"
       End If
       Set rsBlog = Nothing
    End If
    Set rsUser = Nothing
End Sub

Sub Modify()
    Dim BlogID
    Dim rsBlog, rsBlogClass, sqlBlog, rsUser, UName
    BlogID = PE_CLng(Trim(Request("ID")))
    If BlogID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的聚合空间</li>"
        Exit Sub
    End If
    sqlBlog = "Select * from PE_Space where ID=" & BlogID
    Set rsBlog = Server.CreateObject("Adodb.RecordSet")
    rsBlog.Open sqlBlog, Conn, 1, 3
    If rsBlog.BOF And rsBlog.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不存在此聚合空间！</li>"
    Else
        Set rsUser = Conn.Execute("select Top 1 UserName from PE_User Where UserID=" & rsBlog("UserID"))
        If rsUser.BOF And rsUser.EOF Then
            Call main
        Else
            If rsBlog("type") > 1 Then
                Response.Write "<form method='post' action='Admin_SpaceManage.asp' name='myform' onsubmit='return CheckInput();'>"
                Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
                Response.Write "    <tr class='title'> "
                Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>修 改 聚 合 空 间 模 块</strong></font></div></td>"
                Response.Write "    </tr>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>模块名称：</strong><input name='BlogName' type='text' size='20' value='" & rsBlog("Name") & "'><font color='#FF0000'>*</font></td></tr>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>模块类型：</strong><select name='addtype' onChange=""changemode()""><option value=2"
                If rsBlog("type") = 2 Then Response.Write " selected"
                Response.Write ">外部RSS数据</option><option value=3"
                If rsBlog("type") = 3 Then Response.Write " selected"
                Response.Write ">我的日志</option><option value=4"
                If rsBlog("type") = 4 Then Response.Write " selected"
                Response.Write ">我的音乐</option><option value=5"
                If rsBlog("type") = 5 Then Response.Write " selected"
                Response.Write ">我的图书</option><option value=6"
                If rsBlog("type") = 6 Then Response.Write " selected"
                Response.Write ">我的图片</option><option value=7"
                If rsBlog("type") = 7 Then Response.Write " selected"
                Response.Write ">我的连接</option></select></td></tr>"
                If rsBlog("type") = 2 Then
                    Response.Write "<tbody id='url' style='display:'>"
                Else
                    Response.Write "<tbody id='url' style='display:none'>"
                End If
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>ＲＳＳ地址：</strong><input name='LinkUrl' type='text' size='67' maxlength='100' value='" & rsBlog("LinkUrl") & "'> <font color='#FF0000'>* RSS来源地址</font></td></tr></tbody>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>显示条数：</strong><input name='ListNum' type='text' size='2' maxlength='2' value='" & rsBlog("ListNum") & "'> <font color='#FF0000'>* 前台显示条数</font></td></tr>"
                Response.Write "  <tr class='tdbg'> "
                Response.Write "    <td colspan='2'>&nbsp;<strong>简介</strong>↓<br>"
                Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>"
                If Trim(rsBlog("Intro") & "") <> "" Then Response.Write Server.HTMLEncode(rsBlog("Intro"))
                Response.Write "      </textarea>"
                Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
                Response.Write "    </td>"
                Response.Write "  </tr>"
                Response.Write "    <tr>"
                Response.Write "      <td colspan='2' align='center' class='tdbg'>"
                Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
                Response.Write "      <input name='UserID' type='hidden' id='UserID' value=" & UserID & ">"
                Response.Write "      <input name='addtype' type='hidden' id='addtype' value=0>"
                Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsBlog("ID") & ">"
                Response.Write "      <input  type='submit' name='Submit' value='保存修改结果' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_SpaceManage.asp'"" style='cursor:hand;'></td>"
                Response.Write "    </tr>"
                Response.Write "  </table>"
                Response.Write "</form>"
            Else
                Call PopCalendarInit
                Response.Write "<form method='post' action='Admin_SpaceManage.asp' name='myform' onsubmit='return CheckInput();'>"
                Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
                Response.Write "    <tr class='title'> "
                Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>修 改 聚 合 空 间</strong></font></div></td>"
                Response.Write "    </tr>"
                Response.Write "  <tr class='tdbg'><td colspan='2'>&nbsp;<strong>空间名称：</strong><input name='BlogName' type='text' size='20' maxlength='20' value='" & rsBlog("Name") & "'> <font color='#FF0000'>*</font></td></tr>"
                Response.Write "  <tr class='tdbg'><td colspan='2'><table><tr><td>&nbsp;<strong>空间首页<br>&nbsp;显示项目：</strong></td><td>"
                Set rsBlogClass = Conn.Execute("select * from PE_Channel Where Disabled=" & PE_False & " and ModuleType>0 and ModuleType<4 order by OrderID")
                Do While Not rsBlogClass.EOF
                    Response.Write "<input type='checkbox' name='Showitem' value='" & rsBlogClass("ChannelID") & "'"
                    If FoundInArr(rsBlog("LinkUrl"), rsBlogClass("ChannelID"), ",") Then Response.Write " checked"
                    Response.Write ">我在" & rsBlogClass("ChannelName") & "频道发表的作品<br>"
                    rsBlogClass.MoveNext
                Loop
                Response.Write "</td></tr></table></td></tr><tr class='tdbg'>"

                Response.Write "  <tr class='tdbg'> "
                Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong>分类：</strong><select name='BlogType'>" & GetKingOpti(rsBlog("ClassID")) & "</select></td>"
                Response.Write "    <td rowspan='9' align='center' valign='top' class='tdbg'>"
                Response.Write "        <table width='180' height='200' border='1'>"
                Response.Write "            <tr><td width='100%' align='center'>"
                If IsNull(rsBlog("Photo")) Then
                    Response.Write "<img id='showphoto' src='" & InstallDir & "Space/default.gif' width='150' height='172'>"
                Else
                    Response.Write "<img id='showphoto' src='" & rsBlog("Photo") & "' width='150' height='172'>"
                End If
                Response.Write "        </td></tr></table>"
                Response.Write "        <input name='Photo' type='text' size='25' value='" & rsBlog("Photo") & "'><strong>：照 片 地 址</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AdminBlogPic&Uname=" & rsUser("UserName") & "' frameborder=0 scrolling=no width='285' height='25'></iframe>"
                Response.Write "     </td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>日期：</strong><input name='BirthDay' type='text'  value='" & rsBlog("BirthDay") & "' maxlength='20'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BirthDay, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>地址：</strong><input name='Address' type='text'  value='" & rsBlog("Address") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>电话：</strong><input name='Tel' type='text' value='" & rsBlog("Tel") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>传真：</strong><input name='Fax' type='text' value='" & rsBlog("Fax") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>单位：</strong><input name='Company' type='text' value='" & rsBlog("Company") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>部门：</strong><input name='Department' type='text' value='" & rsBlog("Department") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>邮编：</strong><input name='ZipCode' type='text' value='" & rsBlog("ZipCode") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>ＱＱ：</strong><input name='QQ' type='text' value='" & rsBlog("QQ") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'><td>&nbsp;<strong>邮件：</strong><input name='Email' type='text' value='" & rsBlog("Email") & "'></td><td><strong> 主页：</strong><input name='HomePage' type='text' value='" & rsBlog("HomePage") & "'></td></tr>"
                Response.Write "  <tr class='tdbg'> "
                Response.Write "    <td colspan='2'>&nbsp;<strong>简介</strong>↓<br>"
                Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>"
                If Trim(rsBlog("Intro") & "") <> "" Then Response.Write Server.HTMLEncode(rsBlog("Intro"))
                Response.Write "      </textarea>"
                Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
                Response.Write "    </td>"
                Response.Write "  </tr>"
                Response.Write "    <tr>"
                Response.Write "      <td colspan='2' align='center' class='tdbg'>"
                Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
                Response.Write "      <input name='addtype' type='hidden' id='addtype' value=1>"
                Response.Write "      <input name='UserID' type='hidden' id='UserID' value=" & UserID & ">"
                Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsBlog("ID") & ">"
                Response.Write "      <input  type='submit' name='Submit' value='保存修改结果' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_SpaceManage.asp'"" style='cursor:hand;'></td>"
                Response.Write "    </tr>"
                Response.Write "  </table>"
                Response.Write "</form>"
            End If
        End If
    End If
    rsBlog.Close
    Set rsBlog = Nothing
End Sub

Sub SaveAdd()
    Dim BlogName, Birthday, Address, Tel, Fax, Company, Department, ZipCode, Homepage, Email, QQ, Intro, Photo, BlogType, LinkUrl
    Dim rsBlog, sqlBlog, BlogID, isFirst, addtype, listnum, UserName
    isFirst = False

    BlogName = Trim(Request.Form("BlogName"))
    If BlogName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>空间名称不能为空！</li>"
    Else
        BlogName = ReplaceBadChar(BlogName)
    End If
    UserID = PE_CLng(Trim(Request.Form("UserID")))
    If UserID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>未指定用户！</li>"
    End If

    addtype = PE_CLng(Trim(Request.Form("addtype")))
    If addtype = 0 Or addtype = 2 Then
        addtype = 2
        LinkUrl = Trim(Request("LinkUrl"))
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>连接地址不能为空！</li>"
        End If
    ElseIf addtype = 1 Then
        LinkUrl = Trim(Request.Form("Showitem"))
        Birthday = Trim(Request.Form("BirthDay"))
        Photo = Trim(Request.Form("Photo"))
        Address = Trim(Request.Form("Address"))
        Tel = Trim(Request.Form("Tel"))
        Fax = Trim(Request.Form("Fax"))
        Company = Trim(Request.Form("Company"))
        Department = Trim(Request.Form("Department"))
        ZipCode = Trim(Request.Form("ZipCode"))
        Homepage = Trim(Request.Form("HomePage"))
        Email = Trim(Request.Form("Email"))
        QQ = Trim(Request.Form("QQ"))
    End If
    Intro = Trim(Request.Form("Intro"))
    listnum = PE_CLng(Trim(Request.Form("ListNum")))
    If listnum = 0 Then listnum = 10

    BlogType = PE_CLng(Trim(Request.Form("BlogType")))
    If addtype = 1 Then
        Set rsBlog = Conn.Execute("Select Top 1 UserID from PE_Space where UserID=" & UserID & " and Type=1")
        If rsBlog.BOF And rsBlog.EOF Then
            Dim blogdir
            Set blogdir = Conn.Execute("Select Top 1 UserName from PE_User where UserID=" & UserID)
            If Not (blogdir.BOF And blogdir.EOF) Then
                isFirst = True
                UserName = blogdir("UserName")
            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>未指定用户！</li>"
            End If
            Set blogdir = Nothing
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    Set rsBlog = Server.CreateObject("Adodb.RecordSet")
    sqlBlog = "Select * from PE_Space"
    rsBlog.Open sqlBlog, Conn, 1, 3
    BlogID = GetNewID("PE_Space", "ID")
    rsBlog.addnew
        rsBlog("ID") = BlogID
        rsBlog("UserID") = UserID
        If BlogType > 0 Then rsBlog("ClassID") = BlogType
        rsBlog("Name") = BlogName
        If addtype = 1 And Birthday <> "" Then
            rsBlog("BirthDay") = Birthday
        Else
            rsBlog("BirthDay") = Now()
        End If
        If addtype = 1 Then
            rsBlog("Address") = Address
            rsBlog("Tel") = Tel
            rsBlog("Fax") = Fax
            rsBlog("Company") = Company
            rsBlog("Department") = Department
            rsBlog("ZipCode") = ZipCode
            rsBlog("HomePage") = Homepage
            rsBlog("Email") = Email
            rsBlog("QQ") = PE_CLng(QQ)
            If Photo <> "" Then rsBlog("Photo") = Photo
        End If
        rsBlog("Intro") = Intro
        If Trim(LinkUrl & "") = "" Then
           rsBlog("LinkUrl") = Null
        Else
           rsBlog("LinkUrl") = LinkUrl
        End If
        If isFirst = True Then
            rsBlog("Type") = 1
            rsBlog("OrderID") = 1
        Else
            rsBlog("Type") = addtype
            rsBlog("OrderID") = 2
        End If
        rsBlog("LastUseTime") = Now()
        rsBlog("Passed") = True
        rsBlog("listnum") = listnum
    rsBlog.Update
    rsBlog.Close
    Set rsBlog = Nothing
    If addtype = 1 And isFirst = True Then
        Conn.Execute ("update PE_User set Blog=" & PE_True & " where UserID=" & UserID)
        Call CreateBlogDir(UserID, UserName)
    End If
    Call CloseConn
    Response.Redirect ComUrl
End Sub

Sub SaveModify()
    Dim BlogName, BlogID, Birthday, Address, Tel, Fax, Company, Department, ZipCode, Homepage, Email, QQ, Intro, Photo, BlogType, LinkUrl
    Dim rsBlog, sqlBlog, addtype, listnum
    BlogName = Trim(Request.Form("BlogName"))
    If BlogName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>空间名称不能为空！</li>"
    End If
    BlogID = Trim(Request.Form("ID"))
    If BlogID <> "" Then
        If InStr(BlogID, ",") > 0 Then
            BlogID = ReplaceBadChar(BlogID)
        Else
            BlogID = PE_CLng(BlogID)
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的聚合空间！</li>"
    End If
    addtype = PE_CLng(Trim(Request.Form("addtype")))
    If addtype = 0 Or addtype = 2 Then
        addtype = 2
        LinkUrl = Trim(Request.Form("LinkUrl"))
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>连接地址不能为空！</li>"
        End If
    ElseIf addtype = 1 Then
        LinkUrl = Trim(Request.Form("Showitem"))
        Birthday = Trim(Request.Form("BirthDay"))
        Photo = Trim(Request.Form("Photo"))
        Address = Trim(Request.Form("Address"))
        Tel = Trim(Request.Form("Tel"))
        Fax = Trim(Request.Form("Fax"))
        Company = Trim(Request.Form("Company"))
        Department = Trim(Request.Form("Department"))
        ZipCode = Trim(Request.Form("ZipCode"))
        Homepage = Trim(Request.Form("HomePage"))
        Email = Trim(Request.Form("Email"))
        QQ = Trim(Request.Form("QQ"))
    End If
    Intro = Trim(Request.Form("Intro"))
    listnum = PE_CLng(Trim(Request.Form("ListNum")))
    If listnum = 0 Then listnum = 10

    If FoundErr = True Then
        Exit Sub
    End If

    BlogType = PE_CLng(Trim(Request.Form("BlogType")))
    Set rsBlog = Server.CreateObject("Adodb.RecordSet")
    sqlBlog = "Select * from PE_Space where ID=" & BlogID
    rsBlog.Open sqlBlog, Conn, 1, 3
    If Not (rsBlog.BOF And rsBlog.EOF) Then
        If BlogName <> "" Then rsBlog("Name") = BlogName
        If BlogType > 0 Then rsBlog("ClassID") = BlogType
        If addtype = 1 And Birthday <> "" Then
            rsBlog("BirthDay") = Birthday
        End If
        If rsBlog("Type") > 1 Then rsBlog("Type") = addtype
        If addtype = 1 Then
            rsBlog("Address") = Address
            rsBlog("Tel") = Tel
            rsBlog("Fax") = Fax
            rsBlog("Company") = Company
            rsBlog("Department") = Department
            rsBlog("ZipCode") = ZipCode
            rsBlog("HomePage") = Homepage
            rsBlog("Email") = Email
            rsBlog("QQ") = PE_CLng(QQ)
            If Photo <> "" Then rsBlog("Photo") = Photo
        End If
        rsBlog("Intro") = Intro
        If Trim(LinkUrl & "") = "" Then
           rsBlog("LinkUrl") = Null
        Else
           rsBlog("LinkUrl") = LinkUrl
        End If
        rsBlog("listnum") = listnum
        rsBlog.Update
    End If
    rsBlog.Close
    Set rsBlog = Nothing
    Call CloseConn
    Response.Redirect ComUrl
End Sub

Sub SetStat(istat)
    Dim BlogID, OrderID, tmporderid, fl, UserName, UserID
    BlogID = PE_CLng(Trim(Request("ID")))
    If BlogID = 0 And istat < 7 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要操作的聚合空间</li>"
        Exit Sub
    End If
    istat = PE_CLng(istat)
    If istat = 1 Or istat = 2 Then
        Dim blogdir
        Set blogdir = Conn.Execute("Select Top 1 A.ID,C.UserName,C.UserID from PE_Space A left join PE_User C on A.UserID=C.UserID where A.ID=" & BlogID)
        If Not (blogdir.BOF And blogdir.EOF) Then
            UserName = blogdir("UserName")
            UserID = blogdir("UserID")
        End If
        Set blogdir = Nothing
    End If

    Dim spacename
    spacename = Replace(LCase(UserName & UserID), ".", "")

    Select Case istat
    Case 1
        If fso.FolderExists(Server.MapPath(InstallDir & "Space/" & spacename & "/")) = False Then
            Call CreateBlogDir(UserID, UserName)
        End If
        Conn.Execute ("update PE_Space set Passed=" & PE_False & " where ID=" & BlogID & "")
        Set fl = fso.CreateTextFile(Server.MapPath(InstallDir & "Space/" & spacename & "/index.asp"), True)
        fl.WriteLine ("审核中...")
        fl.Close
        Set fl = Nothing
    Case 2
        If fso.FolderExists(Server.MapPath(InstallDir & "Space/" & spacename & "/")) = False Then
            Call CreateBlogDir(UserID, UserName)
        End If
        Conn.Execute ("update PE_Space set Passed=" & PE_True & ",LastUseTime=" & PE_Now & " where ID=" & BlogID & "")
        Set fl = fso.GetFile(Server.MapPath(InstallDir & "Space/Default/index.asp"))
        fl.Copy Server.MapPath(InstallDir & "Space/" & spacename & "/index.asp"), True
        Set fl = Nothing
    Case 3
        Conn.Execute ("update PE_Space set IsElite=" & PE_False & " where ID=" & BlogID & "")
    Case 4
        Conn.Execute ("update PE_Space set IsElite=" & PE_True & " where ID=" & BlogID & "")
    Case 5
        Conn.Execute ("update PE_Space set onTop=" & PE_False & " where ID=" & BlogID & "")
    Case 6
        Conn.Execute ("update PE_Space set onTop=" & PE_True & " where ID=" & BlogID & "")
    Case 7
        tmporderid = Split(Action, "|")
        If UBound(tmporderid) = 1 Then
            BlogID = PE_CLng(tmporderid(1))
            OrderID = Trim(Request("OrderID" & BlogID))
            If OrderID > 0 And BlogID > 0 Then Conn.Execute ("update PE_Space set OrderID=" & OrderID & " where ID=" & BlogID & "")
        End If
    End Select
    Call CloseConn
    Response.Redirect ComUrl
End Sub

Sub Del()
    Dim BlogID
    BlogID = Trim(Request("ID"))
    If IsValidID(BlogID) = False Then
        BlogID = ""
    End If
    If BlogID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的聚合空间</li>"
        Exit Sub
    End If
    Call DelBlogDir(BlogID)
    If InStr(BlogID, ",") > 0 Then
        Conn.Execute ("delete from PE_Space where ID in (" & BlogID & ")")
    Else
        Conn.Execute ("delete from PE_Space where ID=" & BlogID & "")
    End If
    Call CloseConn
    Response.Redirect ComUrl
End Sub

Sub CreateBlogDir(UID, UName)
    If PE_CLng(UID) = 0 Or Trim(UName & "") = "" Then Exit Sub
    On Error Resume Next
    Dim fsfl, fl, strDir, spacename
    spacename = Replace(LCase(UName & UID), ".", "")

    strDir = InstallDir & "Space/" & spacename & "/"
    If fso.FolderExists(Server.MapPath(strDir)) = False Then fso.CreateFolder Server.MapPath(strDir)

    Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & "Space/Default/"))
    For Each fl In fsfl.Files
        fl.Copy Server.MapPath(strDir & fl.name), True
    Next

    Set fsfl = fso.CreateTextFile(Server.MapPath(strDir & "config.xml"), True)
    fsfl.WriteLine ("<?" & "xml version=""1.0"" encoding=""gb2312""" & "?>")
    fsfl.WriteLine ("<" & "body" & ">")
    fsfl.WriteLine ("<" & "baseconfig" & ">")
    fsfl.WriteLine ("<" & "userid" & ">" & UID & "</" & "userid" & ">")
    fsfl.WriteLine ("</" & "baseconfig" & ">")
    fsfl.WriteLine ("</" & "body" & ">")
    fsfl.Close
    Set fsfl = Nothing
End Sub

Sub DelBlogDir(BID)
    Dim UsRs, tmporderid, i, tempuserid, spacename
    On Error Resume Next
    If Trim(BID & "") = "" Then Exit Sub
    If InStr(BID, ",") > 0 Then
        tmporderid = Split(BID, ",")
        For i = 0 To UBound(tmporderid)
            Set UsRs = Conn.Execute("select top 1 A.ID,A.UserID,A.Type,C.UserID,C.UserName from PE_Space A left join PE_User C on A.UserID=C.UserID where A.ID=" & PE_CLng(tmporderid(i)) & " and A.Type=1")
            If Not (UsRs.BOF And UsRs.EOF) Then
                tempuserid = UsRs(1)

                spacename = Replace(LCase(UsRs(4) & tempuserid), ".", "")

                If fso.FolderExists(Server.MapPath(InstallDir & "Space/" & spacename & "/")) Then
                    fso.DeleteFolder Server.MapPath(InstallDir & "Space/" & spacename & "/")
                End If
                '删除全部数据
                Conn.Execute ("delete from PE_Space Where UserID=" & tempuserid)
                Conn.Execute ("delete from PE_SpaceBook Where UserID=" & tempuserid)
                Conn.Execute ("delete from PE_SpaceDiary Where UserID=" & tempuserid)
                Conn.Execute ("delete from PE_SpaceMusic Where UserID=" & tempuserid)
                Conn.Execute ("update PE_User Set Blog=" & PE_False & " Where UserID=" & tempuserid)
            End If
        Next
    Else
        Set UsRs = Conn.Execute("select top 1 A.UserID,A.Type,C.UserName from PE_Space A left join PE_User C on A.UserID=C.UserID where A.ID=" & PE_CLng(BID) & " and A.Type=1")
        If Not (UsRs.BOF And UsRs.EOF) Then
            tempuserid = UsRs(0)
            spacename = Replace(LCase(UsRs(2) & tempuserid), ".", "")
            If fso.FolderExists(Server.MapPath(InstallDir & "Space/" & spacename & "/")) Then
                fso.DeleteFolder Server.MapPath(InstallDir & "Space/" & spacename & "/")
            End If
            '删除全部数据
            Conn.Execute ("delete from PE_Space Where UserID=" & tempuserid)
            Conn.Execute ("delete from PE_SpaceBook Where UserID=" & tempuserid)
            Conn.Execute ("delete from PE_SpaceDiary Where UserID=" & tempuserid)
            Conn.Execute ("delete from PE_SpaceMusic Where UserID=" & tempuserid)
            Conn.Execute ("update PE_User Set Blog=" & PE_False & " Where UserID=" & tempuserid)
        End If
    End If
    Set UsRs = Nothing
End Sub

'*********
'*模块类别管理
'*********

Sub Kind()
    Dim KindID, rsGKind, sqlGKind
    sqlGKind = "select * from PE_Spacekind order by OrderID"
    Set rsGKind = Conn.Execute(sqlGKind)

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td width='50' align='center'><strong>分类ID</strong></td>"
    Response.Write "    <td width='150' align='center'><strong>分类名称</strong></td>"
    Response.Write "    <td align='center'><strong>分类说明</strong></td>"
    Response.Write "    <td width='150' align='center'><strong>常规操作</strong></td>"
    Response.Write "    <td width='100' align='center'><strong>排序操作</strong></td>" & vbCrLf
    Response.Write "  </tr>"
    If rsGKind.BOF And rsGKind.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='5' align='center'>您还没有添加任何分类!</td><tr>" & vbCrLf
    Else
        Do While Not rsGKind.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='50' align='center'>" & rsGKind("KindID") & "</td>"
            Response.Write "    <td width='150' align='center'>" & rsGKind("KindName") & "</td>"
            Response.Write "    <td>" & PE_HTMLEncode(rsGKind("ReadMe")) & "</td>"
            Response.Write "    <td width='150' align='center'>"
            Response.Write "<a href='Admin_SpaceManage.asp?action=ModifyKind&ID=" & rsGKind("KindID") & "'>修改</a>&nbsp;&nbsp;"
            Response.Write "<a href='Admin_SpaceManage.asp?Action=DelKind&ID=" & rsGKind("KindID") & "' onClick=""return confirm('确定要删除此分类吗？删除此模块后原属于此分类的聚合空间将不属于任何分类。');"">删除</a>&nbsp;&nbsp;"
            Response.Write "<a href='Admin_SpaceManage.asp?Action=ClearKind&ID=" & rsGKind("KindID") & "' onClick=""return confirm('确定要清空此分类中的聚合空间吗？');"">清空</a>"
            Response.Write "</td>"
            Response.Write "<form name='orderform' method='post' action='Admin_SpaceManage.asp'>"
            Response.Write "    <td width='100' align='center'>      <input name='OrderID' type='text' id='OrderID' value='" & rsGKind("OrderID") & "' size='4' maxlength='4' style='text-align:center '>"
            Response.Write "      <input name='ID' type='hidden' id='ID' value='" & rsGKind("KindID") & "'>"
            Response.Write "    <input type='submit' name='Submit' value='修改'>"
            Response.Write "    <input name='Action' type='hidden' id='Action' value='OrderKind'></td></form>"
            Response.Write "</tr>"
            rsGKind.MoveNext
        Loop
    End If
    Response.Write "</table>"
    rsGKind.Close
    Set rsGKind = Nothing
End Sub

Sub AddKind()
    Response.Write "<form method='post' action='Admin_SpaceManage.asp' name='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>添加聚合空间模块分类</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>分类名称：</strong></td>"
    Response.Write "      <td class='tdbg'><input name='KindName' type='text' id='KindName' size='49' maxlength='30'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>分类说明</strong><br>鼠标移至分类名称上时将显示设定的说明文字（不支持HTML）</td>"
    Response.Write "      <td class='tdbg'><textarea name='ReadMe' cols='40' rows='5' id='ReadMe'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAddKind'>"
    Response.Write "        <input  type='submit' name='Submit' value=' 添 加 '>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_SpaceManage.asp?Action=Kind'"" style='cursor:hand;'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyKind()
    Dim KindID, rsGKind, sqlGKind
    KindID = Trim(Request("ID"))
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的分类ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    sqlGKind = "Select * from PE_Spacekind Where KindID=" & KindID
    Set rsGKind = Server.CreateObject("Adodb.RecordSet")
    rsGKind.Open sqlGKind, Conn, 1, 3
    If rsGKind.BOF And rsGKind.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的分类，可能已经被删除！</li>"
    Else
        Response.Write "<form method='post' action='Admin_SpaceManage.asp' name='myform'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>修改聚合空间模块分类</strong></div></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg'><strong>分类名称：</strong></td>"
        Response.Write "      <td class='tdbg'><input name='KindName' type='text' id='KindName' value='" & rsGKind("KindName") & "' size='49' maxlength='30'><input name='KindID' type='hidden' id='KindID' value='" & rsGKind("KindID") & "'></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg'><strong>分类说明</strong><br>鼠标移至分类名称上时将显示设定的说明文字（不支持HTML）</td>"
        Response.Write "      <td class='tdbg'><textarea name='ReadMe' cols='40' rows='5' id='ReadMe'>" & rsGKind("ReadMe") & "</textarea></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModifyKind'>"
        Response.Write "        <input name='ID' type='hidden' id='ID' value=" & KindID & "><input  type='submit' name='Submit' value='保存修改结果'>&nbsp;&nbsp;"
        Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_SpaceManage.asp?Action=Kind'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsGKind.Close
    Set rsGKind = Nothing
End Sub

Sub SaveKind()
    Dim KindID, KindName, ReadMe, rs, mrs, intMaxID, OrderID
    KindName = ReplaceBadChar(Trim(Request("KindName")))
    ReadMe = Trim(Request("ReadMe"))
    If KindName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>分类名称不能为空！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If Action = "SaveAddKind" Then
        intMaxID = PE_CLng(Conn.Execute("select max(KindID) from PE_Spacekind")(0)) + 1
        
        Set mrs = Conn.Execute("select max(OrderID) from PE_Spacekind")
        If IsNull(mrs(0)) Then
            OrderID = 1
        Else
            OrderID = mrs(0) + 1
        End If
        Set mrs = Nothing
        
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open "Select * from PE_Spacekind Where KindName='" & KindName & "'", Conn, 1, 3
        If Not (rs.BOF And rs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>分类名称已经存在！</li>"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.addnew
        rs("KindID") = intMaxID
        rs("OrderID") = OrderID
    ElseIf Action = "SaveModifyKind" Then
        KindID = Trim(Request("ID"))
        If KindID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的分类ID！</li>"
            Exit Sub
        Else
            KindID = PE_CLng(KindID)
        End If
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open "Select * from PE_Spacekind Where KindID=" & KindID, Conn, 1, 3
        If rs.BOF And rs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的分类，可能已经被删除！</li>"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If
    rs("KindName") = KindName
    rs("ReadMe") = ReadMe
    rs.Update
    rs.Close
    Set rs = Nothing
    Call CloseConn
    Response.Redirect "Admin_SpaceManage.asp?Action=Kind"
End Sub

Sub DelKind()
    Dim KindID
    KindID = Trim(Request("ID"))
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的分类ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    If FoundErr = True Then Exit Sub

    If Action = "DelKind" Then
        Conn.Execute ("delete from PE_Spacekind where KindID=" & KindID)
    End If
    Conn.Execute ("update PE_Space set ClassID=0 where ClassID=" & KindID)
    Call CloseConn
    Response.Redirect "Admin_SpaceManage.asp?Action=Kind"
End Sub

Sub OrderKind()
    Dim KindID, OrderID
    KindID = Trim(Request("ID"))
    OrderID = Trim(Request("OrderID"))
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定分类ID</li>"
    Else
        KindID = PE_CLng(KindID)
    End If
    If OrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定排序顺序ID</li>"
    Else
        OrderID = PE_CLng(OrderID)
    End If
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_Spacekind set OrderID=" & OrderID & " where KindID=" & KindID & "")
    Call CloseConn
    Response.Redirect "Admin_SpaceManage.asp?Action=Kind"
End Sub

Function GetKindList(iKindID)
    Dim strtmp, rskind

    If iKindID = 0 Then
        strtmp = "| <font color=red>全部分类</font>"
    Else
        strtmp = "| <a href='" & KindFileName & "'>全部分类</a>"
    End If
    Set rskind = Conn.Execute("select KindID,KindName from PE_SpaceKind order by OrderID")
    Do While Not rskind.EOF
        If iKindID = rskind("KindID") Then
            strtmp = strtmp & "| <font color=red>" & rskind("KindName") & "</font>"
        Else
            strtmp = strtmp & "| <a href='" & KindFileName & "&KindID=" & rskind("KindID") & "'>" & rskind("KindName") & "</a>"
        End If
        rskind.MoveNext
    Loop
    Set rskind = Nothing
    GetKindList = strtmp & " |"
End Function

Function GetKingOpti(iselected)
    Dim strtmp, rskind
    Set rskind = Conn.Execute("select KindID,KindName from PE_SpaceKind order by OrderID")
    Do While Not rskind.EOF
        strtmp = strtmp & "<option value=" & rskind("KindID")
        If iselected = rskind("KindID") Then
            strtmp = strtmp & " selected"
        End If
        strtmp = strtmp & ">" & rskind("KindName") & "</option>"
        rskind.MoveNext
    Loop
    Set rskind = Nothing
    strtmp = strtmp & "<option"
    If iselected = 0 Then
        strtmp = strtmp & " selected"
    End If
    strtmp = strtmp & ">不属于任何分类</option>"
    GetKingOpti = strtmp
End Function

Function GetKingName(iselected)
    Dim strtmp, rskind, KindS

    If oldKInd = "" Then oldKInd = "0|||无分类"

    KindS = Split(oldKInd, "|||")
    If KindS(0) <> iselected Then
        Set rskind = Conn.Execute("select top 1 KindID,KindName from PE_SpaceKind Where KindID=" & iselected)
        If Not (rskind.BOF And rskind.EOF) Then
            strtmp = rskind("KindName")
        Else
            strtmp = "无分类"
        End If
        oldKInd = iselected & "|||" & strtmp
        Set rskind = Nothing
    Else
        strtmp = KindS(1)
    End If
    GetKingName = strtmp
End Function
%>
