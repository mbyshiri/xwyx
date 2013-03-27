<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "FriendSite"   '其他权限

Dim KindID, KindName, KindType, KindTypeName, ShowType, Passed

Passed = Trim(Request("Passed"))
KindID = Trim(Request("KindID"))
KindType = Trim(Request("KindType"))
ShowType = Trim(Request("ShowType"))

If Passed = "" Then
    Passed = Session("Passed")
End If
Session("Passed") = Passed
If IsValidID(KindID) = False Then
    KindID = ""
End If
If KindType = "" Then
    KindType = 1
Else
    KindType = PE_CLng(KindType)
End If
If KindType = 1 Then
    KindTypeName = "类别"
ElseIf KindType = 2 Then
    KindTypeName = "专题"
End If
If ShowType = "" Then
    ShowType = Session("ShowType")
End If
If ShowType = "" Then ShowType = "0"
Session("ShowType") = ShowType
FileName = "Admin_FriendSite.asp?Action=" & Action
strFileName = FileName & "&KindID=" & KindID & "&Field=" & strField & "&Keyword=" & Keyword

'页面头部HTML代码
Response.Write "<html><head><title>友情链接管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("友 情 链 接 管 理", 10022)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td>"
Response.Write "    <a href='Admin_FriendSite.asp?ShowType=0'>友情链接管理首页</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_FriendSite.asp?Action=Add'>添加友情链接</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_FriendSite.asp?Action=FsKind&KindType=1'>链接类别管理</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_FriendSite.asp?Action=AddFsKind&KindType=1'>添加链接类别</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_FriendSite.asp?Action=FsKind&KindType=2'>链接专题管理</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_FriendSite.asp?Action=AddFsKind&KindType=2'>添加链接专题</a>&nbsp;|&nbsp;"
Response.Write "    <a href='Admin_FriendSite.asp?Action=Order'>友情链接排序</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
If Action = "" Then
    Response.Write "  <form name='form3' method='Post' action='" & FileName & "'>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='70' height='30' ><strong>管理选项：</strong></td>" & vbCrLf
    Response.Write "    <td>" & vbCrLf
    Response.Write "      <input name='Passed' type='radio' value='All' onclick='submit();' " & IsRadioChecked(Passed, "All") & ">所有链接&nbsp;&nbsp;"
    Response.Write "      <input name='Passed' type='radio' value='False' onclick='submit();' " & IsRadioChecked(Passed, "False") & ">未审核的链接&nbsp;&nbsp;"
    Response.Write "      <input name='Passed' type='radio' value='True' onclick='submit();' " & IsRadioChecked(Passed, "True") & ">已审核的链接&nbsp;&nbsp;|&nbsp;&nbsp;"
    'Response.Write "      <input name='ShowType' type='radio' value='0' onclick='submit();' " & IsRadioChecked(ShowType, "0") & ">所有链接&nbsp;&nbsp;"
    'Response.Write "      <input name='ShowType' type='radio' value='1' onclick='submit();' " & IsRadioChecked(ShowType, "1") & ">LOGO链接&nbsp;&nbsp;"
    'Response.Write "      <input name='ShowType' type='radio' value='2' onclick='submit();' " & IsRadioChecked(ShowType, "2") & ">文字链接&nbsp;&nbsp;|&nbsp;&nbsp;"
    Response.Write "      <input name='KindType' type='radio' value='1' onclick='submit();' " & IsRadioChecked(KindType, 1) & ">按类别分类&nbsp;&nbsp;"
    Response.Write "      <input name='KindType' type='radio' value='2' onclick='submit();' " & IsRadioChecked(KindType, 2) & ">按专题分类"
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  </form>" & vbCrLf
End If
Response.Write "</table>" & vbCrLf

'执行的操作
Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveFriendSite
Case "SetElite", "CancelElite", "SetPassed", "CancelPassed", "Del"
    Call SetProperty
Case "MoveToKind"
    Call MoveToKind
Case "FsKind"
    Call FsKind
Case "AddFsKind"
    Call AddFsKind
Case "ModifyFsKind"
    Call ModifyFsKind
Case "DelFsKind"
    Call DelFsKind
Case "ClearFsKind"
    Call ClearFsKind
Case "SaveAddFsKind", "SaveModifyFsKind"
    Call SaveFsKind
Case "Order"
    Call Order
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsFriendSite, sqlFriendSite
    If KindID <> "" Then
        Dim tKind
        Set tKind = Conn.Execute("select * from PE_FsKind where KindID=" & KindID)
        If tKind.BOF And tKind.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的类别</li>"
            Exit Sub
        Else
            KindName = tKind("KindName")
        End If
    End If

    Call ShowJS_FriendSite
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetFsKindList() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetFriendSitePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_FriendSite.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "    <td>"
    Response.Write "      <table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "        <tr class='title' height='22'> "
    Response.Write "          <td width='30' align='center'><strong>选中</strong></td>"
    Response.Write "          <td width='80' align='center'><strong>链接" & KindTypeName & "</strong></td>"
    Response.Write "          <td width='60' align='center'><strong>链接类型</strong></td>"
    Response.Write "          <td align='center'><strong>网站名称</strong></td>"
    Response.Write "          <td width='100' align='center'><strong>网站LOGO</strong></td>"
    Response.Write "          <td width='60' align='center'><strong>站长</strong></td>"
    Response.Write "          <td width='40' align='center'><strong>点击数</strong></td>"
    Response.Write "          <td width='40' align='center'><strong>状态</strong></td>"
    Response.Write "          <td width='40' align='center'><strong>已审核</strong></td>"
    Response.Write "          <td width='100' align='center'><strong>操作</strong></td>"
    Response.Write "        </tr>"

    sqlFriendSite = "select ID,KindID,SpecialID,LinkType,SiteName,SiteUrl,SiteIntro,LogoUrl,SiteAdmin,SiteEmail,Stars,Hits,Elite,Passed,UpdateTime from PE_FriendSite where "
    If ShowType = "1" Then
        sqlFriendSite = sqlFriendSite & " LinkType=1 "
    ElseIf ShowType = "2" Then
        sqlFriendSite = sqlFriendSite & " LinkType=2 "
    Else
        sqlFriendSite = sqlFriendSite & " 1=1 "
    End If
    If KindID <> "" Then
        If KindType = 1 Then
            sqlFriendSite = sqlFriendSite & " and KindID=" & KindID
        ElseIf KindType = 2 Then
            sqlFriendSite = sqlFriendSite & " and SpecialID=" & KindID
        End If
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "SiteName"
            sqlFriendSite = sqlFriendSite & " and SiteName like '%" & Keyword & "%' "
        Case "SiteUrl"
            sqlFriendSite = sqlFriendSite & " and SiteUrl like '%" & ReplaceUrlBadChar(Trim(Request("keyword"))) & "%' "
        Case "SiteAdmin"
            sqlFriendSite = sqlFriendSite & " and SiteAdmin like '%" & Keyword & "%' "
        Case "SiteIntro"
            sqlFriendSite = sqlFriendSite & " and SiteIntro like '%" & Keyword & "%' "
        End Select
    End If
    If Passed = "True" Then
        sqlFriendSite = sqlFriendSite & " and Passed=" & PE_True & ""
    ElseIf Passed = "False" Then
        sqlFriendSite = sqlFriendSite & " and Passed=" & PE_False & ""
    End If
    sqlFriendSite = sqlFriendSite & " order by ID desc"

    Set rsFriendSite = Server.CreateObject("ADODB.Recordset")
    rsFriendSite.Open sqlFriendSite, Conn, 1, 1
    If rsFriendSite.BOF And rsFriendSite.EOF Then
        totalPut = 0
        If ShowType = "1" Then
            Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何LOGO链接！<br><br></td></tr>"
        ElseIf ShowType = "2" Then
            Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何文字链接！<br><br></td></tr>"
        Else
            Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何友情链接！<br><br></td></tr>"
        End If
    Else
        totalPut = rsFriendSite.RecordCount
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
                rsFriendSite.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim FriendSiteNum
        FriendSiteNum = 0
        Do While Not rsFriendSite.EOF
            Response.Write "        <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "          <td width='30' align='center'>"
            Response.Write "            <input name='ID' type='checkbox' onclick='unselectall()' id='ID' value='" & rsFriendSite("ID") & "'>"
            Response.Write "          </td>"
            Response.Write "          <td width='80' align='center'>"
            If KindType = 1 Then
                Response.Write GetKindName(rsFriendSite("KindID"))
            ElseIf KindType = 2 Then
                Response.Write GetKindName(rsFriendSite("SpecialID"))
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='60' align='center'>"
            If rsFriendSite("LinkType") = 1 Then
                Response.Write "            <a href='Admin_FriendSite.asp?KindType=" & KindType & "&ShowType=1'>LOGO链接</a>"
            Else
                Response.Write "            <a href='Admin_FriendSite.asp?KindType=" & KindType & "&ShowType=2'>文字链接</a>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td>"
            Response.Write "            <a href='" & rsFriendSite("SiteUrl") & "' target='blank' title='"
            Response.Write "网站名称：" & rsFriendSite("SiteName") & vbCrLf
            Response.Write "网站地址：" & rsFriendSite("SiteUrl") & vbCrLf
            Response.Write "评分等级："
            If rsFriendSite("Stars") = 0 Or IsNull(rsFriendSite("Stars")) Then
                Response.Write "无" & vbCrLf
            Else
                Response.Write String(rsFriendSite("Stars"), "★") & vbCrLf
            End If
            Response.Write "点 击 数：" & rsFriendSite("Hits") & vbCrLf
            Response.Write "更新时间：" & rsFriendSite("UpdateTime") & vbCrLf
            Response.Write "网站简介：" & rsFriendSite("SiteIntro")
            Response.Write "'>" & rsFriendSite("SiteName") & "</a>"
            Response.Write "          </td>"
            Response.Write "          <td width='100' align='center'>"
            If rsFriendSite("LogoUrl") <> "" And rsFriendSite("LogoUrl") <> "http://" Then
                If LCase(Right(rsFriendSite("LogoUrl"), 3)) = "swf" Then
                    Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#versFriendSiteion=5,0,0,0' width='88' height='31'><param name='movie' value='" & rsFriendSite("LogoUrl") & "'><param name='quality' value='high'><embed src='" & rsFriendSite("LogoUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='88' height='31'></embed></object>"
                Else
                    Response.Write "<a href='" & rsFriendSite("SiteUrl") & "' target='_blank' title='" & rsFriendSite("LogoUrl") & "'><img src='" & rsFriendSite("LogoUrl") & "' width='88' height='31' border='0'></a>"
                End If
            Else
                Response.Write "&nbsp;"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='60' align='center'>"
            Response.Write "            <a href='mailto:" & rsFriendSite("SiteEmail") & "'>" & rsFriendSite("SiteAdmin") & "</a>"
            Response.Write "          </td>"
            Response.Write "          <td width='40' align='center'>" & rsFriendSite("Hits") & "</td>"
            Response.Write "          <td width='40' align='center'>"
            If rsFriendSite("Elite") = True Then
                Response.Write "<font color=green>推荐</font> "
            Else
                Response.Write "&nbsp;"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='40' align='center'>"
            If rsFriendSite("Passed") = True Then
                Response.Write "<b>√</b>"
            Else
                Response.Write "<font color=red><b>×</b></font>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='100' align='center'>"
            If rsFriendSite("Passed") = False Then
                Response.Write "            <a href='Admin_FriendSite.asp?Action=SetPassed&ID=" & rsFriendSite("ID") & "'>审核通过</a>&nbsp;"
            Else
                Response.Write "            <a href='Admin_FriendSite.asp?Action=CancelPassed&ID=" & rsFriendSite("ID") & "'>取消审核</a>&nbsp;"
            End If
            Response.Write "            <a href='Admin_FriendSite.asp?Action=Modify&ID=" & rsFriendSite("ID") & "'>修改</a><br>"
            If rsFriendSite("Elite") = False Then
                Response.Write "            <a href='Admin_FriendSite.asp?Action=SetElite&ID=" & rsFriendSite("ID") & "'>设为推荐</a>&nbsp;"
            Else
                Response.Write "            <a href='Admin_FriendSite.asp?Action=CancelElite&ID=" & rsFriendSite("ID") & "'>取消推荐</a>&nbsp;"
            End If
            Response.Write "            <a href='Admin_FriendSite.asp?Action=Del&ID=" & rsFriendSite("ID") & "' onclick=""return confirm('确定要删除此友情链接站点吗？');"">删除</a>"
            Response.Write "          </td>"
            Response.Write "        </tr>"
            FriendSiteNum = FriendSiteNum + 1
            If FriendSiteNum >= MaxPerPage Then Exit Do
            rsFriendSite.MoveNext
        Loop
    End If
    rsFriendSite.Close
    Set rsFriendSite = Nothing
    Response.Write "      </table>"
    Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='160' height='30'>"
    Response.Write "            <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中本页所有友情链接"
    Response.Write "          </td>"
    Response.Write "          <td>"
    Response.Write "            <input type='submit' value='删除选定链接' name='submit' onClick=""document.myform.Action.value='Del'"">&nbsp;"
    Response.Write "            <input type='submit' value='设为推荐链接' name='submit' onClick=""document.myform.Action.value='SetElite'"">&nbsp;"
    Response.Write "            <input type='submit' value='取消推荐链接' name='submit' onClick=""document.myform.Action.value='CancelElite'"">&nbsp;"
    Response.Write "            <input type='submit' value='移动选定的链接 ->' name='submit' onClick=""document.myform.Action.value='MoveToKind'""><select name='KindID' id='KindID'>" & GetFsKind_Option(KindType, 0) & "</select>"
    Response.Write "            <input name='KindType' type='hidden' id='KindType' value='" & KindType & "'>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"

    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个友情链接", True)
    End If
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right'><strong>友情链接搜索：</strong></td>"
    Response.Write "    <td>" & GetFsKindSearchForm() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"
End Sub

Sub ShowJS_FriendSite()
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function unselectall(){" & vbCrLf
    Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.Name != 'chkAll'&&e.disabled==false)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "function ConfirmDel(){" & vbCrLf
    Response.Write " if(document.myform.Action.value=='Del'){" & vbCrLf
    Response.Write "     if(confirm('确定要删除选中的友情链接吗？'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='MoveToKind'){" & vbCrLf
    Response.Write "     if(confirm('确定要将选中的友情链接移动到指定的" & KindTypeName & "吗？'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub ShowJS_AddModify()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(document.myform.SiteName.value==''){" & vbCrLf
    Response.Write "    alert('请输入网站名称！');" & vbCrLf
    Response.Write "    document.myform.SiteName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.SiteUrl.value=='' || document.myform.SiteUrl.value=='http://'){" & vbCrLf
    Response.Write "    alert('请输入网站地址！');" & vbCrLf
    Response.Write "    document.myform.SiteUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.Action.value=='SaveAdd'&&document.myform.SitePassword.value==''){" & vbCrLf
    Response.Write "    alert('请输入网站密码！');" & vbCrLf
    Response.Write "    document.myform.SitePassword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.Action.value=='SaveAdd'&&document.myform.SitePwdConfirm.value==''){" & vbCrLf
    Response.Write "    alert('请输入确认密码！');" & vbCrLf
    Response.Write "    document.myform.SitePwdConfirm.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.SitePwdConfirm.value!=document.myform.SitePassword.value){" & vbCrLf
    Response.Write "    alert('网站密码与确认密码不一致！');" & vbCrLf
    Response.Write "    document.myform.SitePwdConfirm.focus();" & vbCrLf
    Response.Write "    document.myform.SitePwdConfirm.select();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Add()
    Call ShowJS_AddModify
    Response.Write "<form method='post' name='myform' onsubmit='return CheckForm()' action='Admin_FriendSite.asp'>"
    Response.Write "  <table border='0' cellpadding='2' cellspacing='1' align='center' width='100%' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>添加友情链接</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>链接所属类别：</strong></td>"
    Response.Write "      <td><select name='KindID' id='KindID'>" & GetFsKind_Option(1, 0) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>链接所属专题：</strong></td>"
    Response.Write "      <td><select name='SpecialID' id='SpecialID'>" & GetFsKind_Option(2, 0) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站名称：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteName' id='SiteName' size='60' maxlength='50' value=''> <font color='#FF0000'> *</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站地址：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteUrl' id='SiteUrl' size='80' maxlength='100' value='http://'> <font color='#FF0000'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站Logo地址：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='LogoUrl' id='LogoUrl' size='80' maxlength='100' value='http://'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>站长姓名：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteAdmin' id='SiteAdmin' size='40'  maxlength='25' value=''>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>电子邮件：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteEmail' id='SiteEmail' size='40'  maxlength='50' value=''>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站密码：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='password' name='SitePassword' id='SitePassword' size='30' maxlength='20' value='123456'> <font color='#FF0000'>*</font> 用于修改信息时用。默认密码为：123456"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>确认密码：</strong></td>"
    Response.Write "      <td><input type='password' name='SitePwdConfirm' id='SitePwdConfirm' size='30' maxlength='20' value='123456'> <font color='#FF0000'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站简介：</strong></td>"
    Response.Write "      <td><textarea name='SiteIntro' id='SiteIntro' cols='67' rows='4'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站评分等级：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <select name='Stars' id='Stars'>" & GetStars(0) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>点 击 数：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Hits' id='Hits' size='10' maxlength='10' value='0'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>录入时间：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='UpdateTime' id='UpdateTime' value='" & Now() & "' maxlength='50'> 时间格式为“年-月-日 时:分:秒”，如：<font color='#0000FF'>2003-5-12 12:32:47</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>是否推荐站点：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='Elite' value='yes' checked> 是&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='Elite' value='no'> 否&nbsp;&nbsp;"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>是否审核通过：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='Passed' value='yes' checked> 是&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='Passed' value='no'> 否&nbsp;&nbsp;"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input type='submit' value=' 确 定 ' name='submit'>&nbsp;&nbsp;"
    Response.Write "        <input type='reset' value=' 重 填 ' name='reset'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim ID, rsFriendSite, sqlFriendSite
    ID = PE_CLng(Trim(Request("ID")))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定友情链接ID</li>"
        Exit Sub
    End If
    sqlFriendSite = "select * from PE_FriendSite where ID=" & ID
    Set rsFriendSite = Conn.Execute(sqlFriendSite)
    If rsFriendSite.BOF And rsFriendSite.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到友情链接！</li>"
        rsFriendSite.Close
        Set rsFriendSite = Nothing
        Exit Sub
    End If

    Call ShowJS_AddModify
    Response.Write "<form method='post' name='myform' onsubmit='return CheckForm()' action='Admin_FriendSite.asp'>"
    Response.Write "  <table border='0' cellpadding='2' cellspacing='1' align='center' width='100%' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修改友情链接</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>链接所属类别：</strong></td>"
    Response.Write "      <td><select name='KindID' id='KindID'>" & GetFsKind_Option(1, rsFriendSite("KindID")) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>链接所属专题：</strong></td>"
    Response.Write "      <td><select name='SpecialID' id='SpecialID'>" & GetFsKind_Option(2, rsFriendSite("SpecialID")) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站名称：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteName' id='SiteName' size='60' maxlength='50' value='" & rsFriendSite("SiteName") & "'> <font color='#FF0000'> *</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站地址：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteUrl' id='SiteUrl' size='80' maxlength='100' value='" & rsFriendSite("SiteUrl") & "'> <font color='#FF0000'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站Logo地址：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='LogoUrl' id='LogoUrl' size='80' maxlength='100' value='" & rsFriendSite("LogoUrl") & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>站长姓名：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteAdmin' id='SiteAdmin' size='40'  maxlength='25' value='" & rsFriendSite("SiteAdmin") & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>电子邮件：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='SiteEmail' id='SiteEmail' size='40'  maxlength='50' value='" & rsFriendSite("SiteEmail") & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站密码：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='password' name='SitePassword' id='SitePassword' size='30' maxlength='20'> <font color='#FF0000'>若不修改，请保持为空</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>确认密码：</strong></td>"
    Response.Write "      <td><input type='password' name='SitePwdConfirm' id='SitePwdConfirm' size='30' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站简介：</strong></td>"
    Response.Write "      <td><textarea name='SiteIntro' id='SiteIntro' cols='67' rows='4'>" & PE_ConvertBR(rsFriendSite("SiteIntro")) & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>网站评分等级：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <select name='Stars' id='Stars'>" & GetStars(rsFriendSite("Stars")) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>点 击 数：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='Hits' id='Hits' size='10' maxlength='10' value='" & rsFriendSite("Hits") & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>录入时间：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='text' name='UpdateTime' id='UpdateTime' value='" & rsFriendSite("UpdateTime") & "' maxlength='50'> 时间格式为“年-月-日 时:分:秒”，如：<font color='#0000FF'>2003-5-12 12:32:47</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>是否推荐站点：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='Elite' value='yes' " & IsRadioChecked(rsFriendSite("Elite"), True) & "> 是&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='Elite' value='no' " & IsRadioChecked(rsFriendSite("Elite"), False) & "> 否&nbsp;&nbsp;"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' align='right'><strong>是否审核通过：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='Passed' value='yes' " & IsRadioChecked(rsFriendSite("Passed"), True) & "> 是&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='Passed' value='no' " & IsRadioChecked(rsFriendSite("Passed"), False) & "> 否&nbsp;&nbsp;"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & rsFriendSite("ID") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input type='submit' value=' 修 改 ' name='submit'>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_FriendSite.asp'"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsFriendSite.Close
    Set rsFriendSite = Nothing
End Sub

Sub SaveFriendSite()
    Dim rsFriendSite, sqlFriendSite
    Dim ID, KindID, SpecialID, LinkType, SiteName, SiteUrl, SiteIntro, LogoUrl
    Dim SiteAdmin, SiteEmail, SitePassword, SitePwdConfirm, Stars, Hits, UpdateTime, Elite, Passed

    ID = PE_CLng(Trim(Request.Form("ID")))
    KindID = PE_CLng(Trim(Request.Form("KindID")))
    SpecialID = PE_CLng(Trim(Request.Form("SpecialID")))
    SiteName = Trim(Request.Form("SiteName"))
    SiteUrl = Trim(Request.Form("SiteUrl"))
    SiteIntro = Trim(Request.Form("SiteIntro"))
    LogoUrl = Trim(Request.Form("LogoUrl"))
    SiteAdmin = Trim(Request.Form("SiteAdmin"))
    SiteEmail = Trim(Request.Form("SiteEmail"))
    SitePassword = Trim(Request.Form("SitePassword"))
    SitePwdConfirm = Trim(Request.Form("SitePwdConfirm"))
    Stars = PE_CLng(Trim(Request.Form("Stars")))
    Hits = PE_CLng(Trim(Request.Form("Hits")))
    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))
    Elite = Trim(Request.Form("Elite"))
    Passed = Trim(Request.Form("Passed"))

    If SiteName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>网站名称不能为空！</li>"
    End If
    If SiteUrl = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>网站地址不能为空！</li>"
    End If
    If SiteEmail <> "" And IsValidEmail(SiteEmail) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Email地址错误!</li>"
    End If

    If Action = "SaveAdd" Then
        If SitePassword = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>网站密码不能为空！</li>"
        End If
        If SitePwdConfirm = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>确认密码不能为空！</li>"
        End If
        If SitePwdConfirm <> SitePassword Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>网站密码与确认密码不一致！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    If LogoUrl = "" Or LogoUrl = "http://" Then
        LinkType = 2
    Else
        LinkType = 1
    End If
    SiteName = ReplaceBadChar(SiteName)
    SiteUrl = ReplaceUrlBadChar(SiteUrl)
    LogoUrl = ReplaceUrlBadChar(LogoUrl)
    SiteAdmin = PE_HTMLEncode(SiteAdmin)
    SiteEmail = PE_HTMLEncode(SiteEmail)
    SiteIntro = PE_HTMLEncode(SiteIntro)
    Elite = CBool(Elite = "yes")
    Passed = CBool(Passed = "yes")

    Set rsFriendSite = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        sqlFriendSite = "select top 1 * from PE_FriendSite where SiteName='" & SiteName & "' and SiteUrl='" & SiteUrl & "'"
        rsFriendSite.Open sqlFriendSite, Conn, 1, 3
        If Not (rsFriendSite.BOF And rsFriendSite.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>你要添加的网站已经存在！</li>"
            rsFriendSite.Close
            Set rsFriendSite = Nothing
            Exit Sub
        End If
        rsFriendSite.addnew
    ElseIf Action = "SaveModify" Then
        If ID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定友情链接ID</li>"
            Exit Sub
        End If
        sqlFriendSite = "select * from PE_FriendSite where ID=" & ID
        rsFriendSite.Open sqlFriendSite, Conn, 1, 3
        If rsFriendSite.BOF And rsFriendSite.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的友情链接！</li>"
            rsFriendSite.Close
            Set rsFriendSite = Nothing
            Exit Sub
        End If
    End If
    rsFriendSite("KindID") = KindID
    rsFriendSite("SpecialID") = SpecialID
    rsFriendSite("LinkType") = LinkType
    rsFriendSite("SiteName") = SiteName
    rsFriendSite("SiteUrl") = SiteUrl
    rsFriendSite("SiteIntro") = SiteIntro
    rsFriendSite("LogoUrl") = LogoUrl
    rsFriendSite("SiteAdmin") = SiteAdmin
    rsFriendSite("SiteEmail") = SiteEmail
    If Action = "SaveAdd" Or (Action = "SaveModify" And SitePassword <> "") Then
        rsFriendSite("SitePassword") = MD5(SitePassword, 16)
    End If
    If Action = "SaveAdd" Then
        rsFriendSite("OrderID") = GetNewID("PE_FriendSite", "OrderID")
    End If
    rsFriendSite("Stars") = Stars
    rsFriendSite("Hits") = Hits
    rsFriendSite("UpdateTime") = UpdateTime
    rsFriendSite("Elite") = Elite
    rsFriendSite("Passed") = Passed
    rsFriendSite.Update
    rsFriendSite.Close
    Set rsFriendSite = Nothing
    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect "Admin_FriendSite.asp"
End Sub

Sub SetProperty()
    Dim ID, sqlProperty, rsProperty
    ID = Trim(Request("ID"))
    If IsValidID(ID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定友情链接！</li>"
        Exit Sub
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    End If

    If InStr(ID, ",") > 0 Then
        sqlProperty = "select * from PE_FriendSite where ID in (" & ID & ")"
    Else
        sqlProperty = "select * from PE_FriendSite where ID=" & ID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetElite"
            rsProperty("Elite") = True
        Case "CancelElite"
            rsProperty("Elite") = False
        Case "SetPassed"
            rsProperty("Passed") = True
        Case "CancelPassed"
            rsProperty("Passed") = False
        Case "Del"
            rsProperty.Delete
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing
    
    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub MoveToKind()
    Dim ID
    ID = Trim(Request("ID"))
    If IsValidID(ID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要移动的友情链接！</li>"
        Exit Sub
    End If
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定目标类别！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    If KindType = 1 Then
        Conn.Execute ("update PE_FriendSite set KindID=" & KindID & " where ID in (" & ID & ")")
    ElseIf KindType = 2 Then
        Conn.Execute ("update PE_FriendSite set SpecialID=" & KindID & " where ID in (" & ID & ")")
    End If
    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect ComeUrl
End Sub


Sub FsKind()
    Dim rsFsKind, sqlFsKind
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "    <td width='200' align='center'><strong>" & KindTypeName & "名称</strong></td>"
    Response.Write "    <td align='center'><strong>" & KindTypeName & "说明</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>包含链接数</strong></td>"
    Response.Write "    <td width='120' align='center'><strong>常规操作</strong></td>"
    Response.Write "  </tr>"

    sqlFsKind = "select * from PE_FsKind where KindType=" & KindType & " order by KindID"
    Set rsFsKind = Conn.Execute(sqlFsKind)
    Do While Not rsFsKind.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td width='30' align='center'>" & rsFsKind("KindID") & "</td>"
        Response.Write "    <td width='200' align='center'>"
        Response.Write "      <a href='Admin_FriendSite.asp?KindID=" & rsFsKind("KindID") & "' title='点击进入管理此" & KindTypeName & "的友情链接'>" & PE_HTMLEncode(rsFsKind("KindName")) & "</a>"
        Response.Write "    </td>"
        Response.Write "    <td>" & rsFsKind("ReadMe") & "</td>"
        Response.Write "    <td width='80' align='center'>" & GetLinkNum(KindType, rsFsKind("KindID")) & "</td>"
        Response.Write "    <td width='120' align='center'>"
        Response.Write "      <a href='Admin_FriendSite.asp?action=ModifyFsKind&KindType=" & KindType & "&KindID=" & rsFsKind("KindID") & "'>修改</a>&nbsp;"
        Response.Write "      <a href='Admin_FriendSite.asp?Action=DelFsKind&KindType=" & KindType & "&KindID=" & rsFsKind("KindID") & "' onClick=""return confirm('确定要删除此" & KindTypeName & "吗？删除此" & KindTypeName & "后原属于此" & KindTypeName & "的友情链接将不属于任何" & KindTypeName & "。');"">删除</a>&nbsp;"
        Response.Write "      <a href='Admin_FriendSite.asp?Action=ClearFsKind&KindType=" & KindType & "&KindID=" & rsFsKind("KindID") & "' onClick=""return confirm('确定要清空此" & KindTypeName & "中的友情链接吗？本操作将原属于此" & KindTypeName & "的友情链接改为不属于任何" & KindTypeName & "。');"">清空</a>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        rsFsKind.MoveNext
    Loop
    rsFsKind.Close
    Set rsFsKind = Nothing
    Response.Write "</table>"
End Sub

Sub AddFsKind()
    Response.Write "<form name='myform' method='post' action='Admin_FriendSite.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>添加友情链接" & KindTypeName & "</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>" & KindTypeName & "名称：</strong></td>"
    Response.Write "      <td class='tdbg'>"
    Response.Write "        <input name='KindName' type='text' id='KindName' size='49' maxlength='30'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>" & KindTypeName & "说明</strong><br>鼠标移至" & KindTypeName & "名称上时将显示设定的说明文字（不支持HTML）</td>"
    Response.Write "      <td class='tdbg'>"
    Response.Write "        <textarea name='ReadMe' cols='40' rows='5' id='ReadMe'></textarea>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAddFsKind'>"
    Response.Write "        <input name='KindType' type='hidden' id='KindType' value='" & KindType & "'>"
    Response.Write "        <input  type='submit' name='Submit' value=' 添 加 '>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_FriendSite.asp'"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyFsKind()
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的" & KindTypeName & "ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    Dim rsFsKind, sqlFsKind
    sqlFsKind = "Select * from PE_FsKind Where KindID=" & KindID
    Set rsFsKind = Conn.Execute(sqlFsKind)
    If rsFsKind.BOF And rsFsKind.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的" & KindTypeName & "！</li>"
        rsFsKind.Close
        Set rsFsKind = Nothing
        Exit Sub
    End If

    Response.Write "<form name='myform' method='post' action='Admin_FriendSite.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>修改友情链接" & KindTypeName & "</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>" & KindTypeName & "名称：</strong></td>"
    Response.Write "      <td class='tdbg'>"
    Response.Write "        <input name='KindName' type='text' id='KindName' size='49' maxlength='30' value='" & rsFsKind("KindName") & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg'><strong>" & KindTypeName & "说明</strong><br>鼠标移至" & KindTypeName & "名称上时将显示设定的说明文字（不支持HTML）</td>"
    Response.Write "      <td class='tdbg'>"
    Response.Write "        <textarea name='ReadMe' cols='40' rows='5' id='ReadMe'>" & PE_ConvertBR(rsFsKind("ReadMe")) & "</textarea>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'>"
    Response.Write "        <input name='KindID' type='hidden' id='KindID' value='" & rsFsKind("KindID") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyFsKind'>"
    Response.Write "        <input name='KindType' type='hidden' id='KindType' value='" & KindType & "'>"
    Response.Write "        <input  type='submit' name='Submit' value='保存修改结果'>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_FriendSite.asp'"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"

    rsFsKind.Close
    Set rsFsKind = Nothing
End Sub

Sub SaveFsKind()
    Dim KindID, KindName, ReadMe
    Dim rsFsKind, sqlFsKind
    KindID = PE_CLng(Trim(Request.Form("KindID")))
    KindName = Trim(Request.Form("KindName"))
    ReadMe = Trim(Request.Form("ReadMe"))

    If KindName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & KindTypeName & "名称不能为空！</li>"
    Else
        If CheckBadChar(KindName) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & KindTypeName & "名称中含有非法字符！</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    
    KindName = PE_HTMLEncode(KindName)
    ReadMe = PE_HTMLEncode(ReadMe)

    Set rsFsKind = Server.CreateObject("Adodb.RecordSet")
    If Action = "SaveAddFsKind" Then
        sqlFsKind = "select top 1 * from PE_FsKind"
        rsFsKind.Open sqlFsKind, Conn, 1, 3
        rsFsKind.addnew
        Dim mrs
        Set mrs = Conn.Execute("select max(KindID) from PE_FsKind")
        If IsNull(mrs(0)) Then
            KindID = 1
        Else
            KindID = mrs(0) + 1
        End If
        Set mrs = Nothing
        rsFsKind("KindID") = KindID
    ElseIf Action = "SaveModifyFsKind" Then
        If KindID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的" & KindTypeName & "ID！</li>"
            Exit Sub
        Else
            sqlFsKind = "Select * from PE_FsKind Where KindID=" & KindID
            rsFsKind.Open sqlFsKind, Conn, 1, 3
            If rsFsKind.BOF And rsFsKind.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到指定的" & KindTypeName & "！</li>"
                rsFsKind.Close
                Set rsFsKind = Nothing
                Exit Sub
            End If
        End If
    End If
    rsFsKind("KindName") = KindName
    rsFsKind("KindType") = KindType
    rsFsKind("ReadMe") = ReadMe
    rsFsKind.Update
    rsFsKind.Close
    Set rsFsKind = Nothing

    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect "Admin_FriendSite.asp?Action=FsKind&KindType=" & KindType
End Sub

Sub DelFsKind()
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的" & KindTypeName & "ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If

    Conn.Execute ("delete from PE_FsKind where KindID=" & KindID)
    Conn.Execute ("update PE_FriendSite set KindID=0 where KindID=" & KindID)
    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect "Admin_FriendSite.asp?Action=FsKind&KindType=" & KindType
End Sub

Sub ClearFsKind()
    If KindID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要清空的" & KindTypeName & "ID！</li>"
        Exit Sub
    Else
        KindID = PE_CLng(KindID)
    End If
    Conn.Execute ("update PE_FriendSite set KindID=0 where KindID=" & KindID)
    Call WriteSuccessMsg("清空此" & KindTypeName & "的友情链接成功。", ComeUrl)
    Call ClearSiteCache(0)
End Sub

Sub Order()
    Dim rsFriendSite, sqlFriendSite, iCount, i, j

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_FriendSite.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "    <td>"
    Response.Write "      <table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "        <tr class='title' height='22'> "
    Response.Write "          <td width='30' align='center'><strong>序号</strong></td>"
    Response.Write "          <td width='80' align='center'><strong>链接" & KindTypeName & "</strong></td>"
    Response.Write "          <td width='60' align='center'><strong>链接类型</strong></td>"
    Response.Write "          <td align='center'><strong>网站名称</strong></td>"
    Response.Write "          <td width='100' align='center'><strong>网站LOGO</strong></td>"
    Response.Write "          <td width='60' align='center'><strong>站长</strong></td>"
    Response.Write "          <td width='240' Colspan='2' align='center'><strong>操作</strong></td>"
    Response.Write "        </tr>"

    sqlFriendSite = "select ID,KindID,SpecialID,LinkType,SiteName,SiteUrl,SiteIntro,LogoUrl,SiteAdmin,SiteEmail,Stars,Hits,Elite,OrderID,Passed,UpdateTime from PE_FriendSite"
    sqlFriendSite = sqlFriendSite & " order by OrderID asc"

    Set rsFriendSite = Server.CreateObject("ADODB.Recordset")
    rsFriendSite.Open sqlFriendSite, Conn, 1, 1
    iCount = rsFriendSite.RecordCount
    j = 1
    If rsFriendSite.BOF And rsFriendSite.EOF Then
        If ShowType = "1" Then
            Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何LOGO链接！<br><br></td></tr>"
        ElseIf ShowType = "2" Then
            Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何文字链接！<br><br></td></tr>"
        Else
            Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何友情链接！<br><br></td></tr>"
        End If
    Else
        Do While Not rsFriendSite.EOF
            Response.Write "        <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "          <td width='30' align='center'>"
            Response.Write rsFriendSite("OrderID")
            Response.Write "          </td>"
            Response.Write "          <td width='80' align='center'>"
            If KindType = 1 Then
                Response.Write GetKindName(rsFriendSite("KindID"))
            ElseIf KindType = 2 Then
                Response.Write GetKindName(rsFriendSite("SpecialID"))
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='60' align='center'>"
            If rsFriendSite("LinkType") = 1 Then
                Response.Write "            <a href='Admin_FriendSite.asp?KindType=" & KindType & "&ShowType=1'>LOGO链接</a>"
            Else
                Response.Write "            <a href='Admin_FriendSite.asp?KindType=" & KindType & "&ShowType=2'>文字链接</a>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td>"
            Response.Write "            <a href='" & rsFriendSite("SiteUrl") & "' target='blank' title='"
            Response.Write "网站名称：" & rsFriendSite("SiteName") & vbCrLf
            Response.Write "网站地址：" & rsFriendSite("SiteUrl") & vbCrLf
            Response.Write "评分等级："
            If rsFriendSite("Stars") = 0 Or IsNull(rsFriendSite("Stars")) Then
                Response.Write "无" & vbCrLf
            Else
                Response.Write String(rsFriendSite("Stars"), "★") & vbCrLf
            End If
            Response.Write "点 击 数：" & rsFriendSite("Hits") & vbCrLf
            Response.Write "更新时间：" & rsFriendSite("UpdateTime") & vbCrLf
            Response.Write "网站简介：" & rsFriendSite("SiteIntro")
            Response.Write "'>" & rsFriendSite("SiteName") & "</a>"
            Response.Write "          </td>"
            Response.Write "          <td width='100' align='center'>"
            If rsFriendSite("LogoUrl") <> "" And rsFriendSite("LogoUrl") <> "http://" Then
                If LCase(Right(rsFriendSite("LogoUrl"), 3)) = "swf" Then
                    Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#versFriendSiteion=5,0,0,0' width='88' height='31'><param name='movie' value='" & rsFriendSite("LogoUrl") & "'><param name='quality' value='high'><embed src='" & rsFriendSite("LogoUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='88' height='31'></embed></object>"
                Else
                    Response.Write "<a href='" & rsFriendSite("SiteUrl") & "' target='_blank' title='" & rsFriendSite("LogoUrl") & "'><img src='" & rsFriendSite("LogoUrl") & "' width='88' height='31' border='0'></a>"
                End If
            Else
                Response.Write "&nbsp;"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='60' align='center'>"
            Response.Write "            <a href='mailto:" & rsFriendSite("SiteEmail") & "'>" & rsFriendSite("SiteAdmin") & "</a>"
            Response.Write "          </td>"
            Response.Write "<form action='Admin_FriendSite.asp?Action=UpOrder' method='post'>"
            Response.Write "          <td width='120' align='center'>"

            If j > 1 Then
                Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>"
                For i = 1 To j - 1
                    Response.Write "<option value=" & i & ">" & i & "</option>"
                Next
                Response.Write "</select>"
                Response.Write "<input type=hidden name=iFriendSiteID value=" & rsFriendSite("ID") & ">"
                Response.Write "<input type=hidden name=cOrderID value=" & rsFriendSite("OrderID") & ">&nbsp;<input type=submit name=Submit value=修改>"
            Else
                Response.Write "&nbsp;"
            End If
            Response.Write "</td></form>"
            Response.Write "<form action='Admin_FriendSite.asp?Action=DownOrder' method='post'>"
            Response.Write "  <td width='120' align='center'>"
            If iCount > j Then
                Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>"
                For i = 1 To iCount - j
                    Response.Write "<option value=" & i & ">" & i & "</option>"
                Next
                Response.Write "</select>"
                Response.Write "<input type=hidden name=iFriendSiteID value=" & rsFriendSite("ID") & ">"
                Response.Write "<input type=hidden name=cOrderID value=" & rsFriendSite("OrderID") & ">&nbsp;<input type=submit name=Submit value=修改>"
            Else
                Response.Write "&nbsp;"
            End If
            Response.Write "</td></form></tr>"
            j = j + 1
            rsFriendSite.MoveNext
        Loop
    End If
    rsFriendSite.Close
    Set rsFriendSite = Nothing
    Response.Write "      </table>"

    Response.Write "    </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"

    Response.Write "<br>"

End Sub


Sub UpOrder()
    Dim FriendSiteID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsFriendSite
    FriendSiteID = Trim(Request("iFriendSiteID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If FriendSiteID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        FriendSiteID = PE_CLng(FriendSiteID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cOrderID = PE_CLng(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_FriendSite")
    MaxOrderID = mrs(0) + 1
    '先将当前项目移至最后
    Conn.Execute ("update PE_FriendSite set OrderID=" & MaxOrderID & " where ID=" & FriendSiteID)
    
    '然后将位于当前项目以上的项目的OrderID依次加一，范围为要提升的数字
    sqlOrder = "select * from PE_FriendSite where OrderID<" & cOrderID & " order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Response.Redirect ("Admin_FriendSite.asp?Action=Order")
        Exit Sub        '如果当前项目已经在最上面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID
        Conn.Execute ("update PE_FriendSite set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前项目从最后移到相应位置
    Conn.Execute ("update PE_FriendSite set OrderID=" & tOrderID & " where ID=" & FriendSiteID)

    Response.Redirect ("Admin_FriendSite.asp?Action=Order")
    Call ClearSiteCache(0)
End Sub

Sub DownOrder()
    Dim FriendSiteID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsFriendSite, PrevID, NextID
    FriendSiteID = Trim(Request("iFriendSiteID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If FriendSiteID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        FriendSiteID = PE_CLng(FriendSiteID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cOrderID = PE_CLng(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_FriendSite")
    MaxOrderID = mrs(0) + 1
    '先将当前项目移至最后
    Conn.Execute ("update PE_FriendSite set OrderID=" & MaxOrderID & " where ID=" & FriendSiteID)
    
    '然后将位于当前项目以下的项目的OrderID依次减一，范围为要下降的数字
    sqlOrder = "select * from PE_FriendSite where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前项目已经在最下面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID
        Conn.Execute ("update PE_FriendSite set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前项目从最后移到相应位置
    Conn.Execute ("update PE_FriendSite set OrderID=" & tOrderID & " where ID=" & FriendSiteID)
    
    Response.Redirect ("Admin_FriendSite.asp?Action=Order")
    Call ClearSiteCache(0)
End Sub


Function GetLinkNum(iKindType, iKindID)
    Dim rsLinkNum
    If iKindType = 1 Then
        Set rsLinkNum = Conn.Execute("select count(ID) from PE_FriendSite where KindID=" & iKindID & "")
    ElseIf iKindType = 2 Then
        Set rsLinkNum = Conn.Execute("select count(ID) from PE_FriendSite where SpecialID=" & iKindID & "")
    Else
        GetLinkNum = 0
        Exit Function
    End If
    If IsNull(rsLinkNum(0)) Then
        GetLinkNum = 0
    Else
        GetLinkNum = rsLinkNum(0)
    End If
    Set rsLinkNum = Nothing
End Function

Function GetFsKindList()
    Dim rsFsKind, sqlFsKind, strFsKind, i
    sqlFsKind = "select * from PE_FsKind"
    If KindType > 0 Then
        sqlFsKind = sqlFsKind & " where KindType=" & KindType
    End If
    sqlFsKind = sqlFsKind & " order by KindID"
    Set rsFsKind = Conn.Execute(sqlFsKind)
    If rsFsKind.BOF And rsFsKind.EOF Then
        strFsKind = strFsKind & "没有任何" & KindTypeName
    Else
        i = 1
        strFsKind = "| "
        Do While Not rsFsKind.EOF
            If rsFsKind("KindID") = KindID Then
                strFsKind = strFsKind & "<a href='" & FileName & "&KindType=" & KindType & "&KindID=" & KindID & "'><font color=red>" & rsFsKind("KindName") & "</font></a>"
            Else
                strFsKind = strFsKind & "<a href='" & FileName & "&KindType=" & KindType & "&KindID=" & rsFsKind("KindID") & "'>" & rsFsKind("KindName") & "</a>"
            End If
            strFsKind = strFsKind & " | "
            i = i + 1
            If i Mod 10 = 0 Then
                strFsKind = strFsKind & "<br>"
            End If
            rsFsKind.MoveNext
        Loop
    End If
    rsFsKind.Close
    Set rsFsKind = Nothing
    GetFsKindList = strFsKind
End Function

Function GetFriendSitePath()
    Dim strPath
    strPath = "您现在的位置：&nbsp;<a href='Admin_FriendSite.asp?ShowType=0'>友情链接管理</a>&nbsp;&gt;&gt;&nbsp;"
    If KindType = 1 Then
        strPath = strPath & "按类别分类&nbsp;&gt;&gt;&nbsp;"
    ElseIf KindType = 2 Then
        strPath = strPath & "按专题分类&nbsp;&gt;&gt;&nbsp;"
    End If
    If KindID <> "" Then
        strPath = strPath & "<a href='" & FileName & "&KindID=" & KindID & "'>" & KindName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If Keyword = "" Then
        If ShowType = "1" Then
            strPath = strPath & "所有LOGO链接"
        ElseIf ShowType = "2" Then
            strPath = strPath & "所有文字链接"
        Else
            strPath = strPath & "所有友情链接"
        End If
    Else
        Select Case strField
            Case "SiteName"
                strPath = strPath & "网站名称中含有 <font color=red>" & Keyword & "</font> "
            Case "SiteUrl"
                strPath = strPath & "网站地址中含有 <font color=red>" & ReplaceUrlBadChar(Trim(Request("keyword"))) & "</font> "
            Case "SiteAdmin"
                strPath = strPath & "站长姓名中含有 <font color=red>" & Keyword & "</font> "
            Case "SiteIntro"
                strPath = strPath & "网站简介中含有 <font color=red>" & Keyword & "</font> "
        End Select
        If ShowType = "1" Then
            strPath = strPath & "的LOGO链接"
        ElseIf ShowType = "2" Then
            strPath = strPath & "的文字链接"
        Else
            strPath = strPath & "的友情链接"
        End If
    End If
    GetFriendSitePath = strPath
End Function

Function GetKindName(iKindID)
    Dim strKindName, rsFsKind, sqlFsKind
    If iKindID > 0 Then
        sqlFsKind = "select KindName from PE_FsKind where KindID=" & iKindID
        Set rsFsKind = Conn.Execute(sqlFsKind)
        If rsFsKind.BOF And rsFsKind.EOF Then
            strKindName = ""
        Else
            strKindName = "<a href='" & FileName & "&KindType=" & KindType & "&KindID=" & iKindID & "'>" & rsFsKind(0) & "</a>"
        End If
        rsFsKind.Close
        Set rsFsKind = Nothing
    End If
    GetKindName = strKindName
End Function

Function GetFsKind_Option(iKindType, KindID)
    Dim sqlFsKind, rsFsKind, strOption
    strOption = "<option value='0'"
    If KindID = "" Then
        strOption = strOption & " selected"
    End If
    If iKindType = 1 Then
        strOption = strOption & ">不属于任何类别</option>"
    ElseIf iKindType = 2 Then
        strOption = strOption & ">不属于任何专题</option>"
    End If
    sqlFsKind = "select * from PE_FsKind"
    If iKindType > 0 Then
        sqlFsKind = sqlFsKind & " where KindType=" & iKindType
    End If
    sqlFsKind = sqlFsKind & " order by KindID"
    Set rsFsKind = Conn.Execute(sqlFsKind)
    Do While Not rsFsKind.EOF
        If rsFsKind("KindID") = KindID Then
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "' selected>" & rsFsKind("KindName") & "</option>"
        Else
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "'>" & rsFsKind("KindName") & "</option>"
        End If
        rsFsKind.MoveNext
    Loop
    rsFsKind.Close
    Set rsFsKind = Nothing
    GetFsKind_Option = strOption
End Function

Function GetFsKindSearchForm()
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    strForm = strForm & "<tr><td height='28' align='center'> "
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='SiteName' selected>网站名称</option>"
    strForm = strForm & "<option value='SiteUrl'>网站地址</option>"
    strForm = strForm & "<option value='SiteAdmin'>站长姓名</option>"
    strForm = strForm & "<option value='SiteIntro'>网站简介</option>"
    strForm = strForm & "</select> "
    strForm = strForm & "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'> "
    strForm = strForm & "<input type='submit' name='Submit'  value='搜索'>"
    strForm = strForm & "<input name='KindID' type='hidden' id='KindID' value='" & KindID & "'>"
    strForm = strForm & "</td></tr></form></table>"
    GetFsKindSearchForm = strForm
End Function

Function IsRadioChecked(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function
%>
