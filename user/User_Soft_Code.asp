<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Common.Purview.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim SoftID, AuthorName, Status, ManageType
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created

Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview
Dim arrFields_Options, arrSoftType, arrSoftLanguage, arrCopyrightType, arrOperatingSystem
    

Sub Execute()
    ChannelID = PE_CLng(Trim(Request("ChannelID")))
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要查看的频道ID！</li>"
        Response.Write ErrMsg
        Exit Sub
    End If
    SoftID = Trim(Request("SoftID"))
    ClassID = PE_CLng(Trim(Request("ClassID")))
    Status = Trim(Request("Status"))
    AuthorName = Trim(Request("AuthorName"))
    strField = Trim(Request("Field"))
    If Status = "" Then
        Status = 9
    Else
        Status = PE_CLng(Status)
    End If
    If IsValidID(SoftID) = False Then
        SoftID = ""
    End If

    arrFields_Options = Split(",,,", ",")
    arrSoftType = ""
    arrSoftLanguage = ""
    arrCopyrightType = ""
    arrOperatingSystem = ""
    If Fields_Options & "" <> "" Then
        arrFields_Options = Split(Fields_Options, "$$$")
        If UBound(arrFields_Options) = 3 Then
            arrSoftType = Split(arrFields_Options(0), vbCrLf)
            arrSoftLanguage = Split(arrFields_Options(1), vbCrLf)
            arrCopyrightType = Split(arrFields_Options(2), vbCrLf)
            arrOperatingSystem = Split(arrFields_Options(3), vbCrLf)
        End If
    End If

    If Action = "" Then Action = "Manage"
    FileName = "User_Soft.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
    strFileName = FileName & "&ClassID=" & ClassID & "&Field=" & strField & "&keyword=" & Keyword
    If AuthorName <> "" Then
        AuthorName = ReplaceBadChar(AuthorName)
        strFileName = strFileName & "&AuthorName=" & AuthorName
    End If

    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

    Response.Write "<table align='center'><tr align='center' valign='top'>"
    If CheckUser_ChannelInput() Then
        Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Add'><img src='images/Soft_add.gif' border='0' align='absmiddle'><br>添加" & ChannelShortName & "</a></td>"
    End If
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=9'><img src='images/Soft_all.gif' border='0' align='absmiddle'><br>所有" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=-1'><img src='images/Soft_draft.gif' border='0' align='absmiddle'><br>草 稿</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=0'><img src='images/Soft_unpassed.gif' border='0' align='absmiddle'><br>待审核的" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=3'><img src='images/Soft_passed.gif' border='0' align='absmiddle'><br>已审核的" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=-2'><img src='images/Soft_reject.gif' border='0' align='absmiddle'><br>未被采用的" & ChannelShortName & "</a></td>"
    Response.Write "</tr></table>" & vbCrLf

    Select Case Action
    Case "Add"
        Call Add
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SaveSoft
    Case "Preview"
        Call Preview
    Case "Show"
        Call Show
    Case "Del"
        Call Del
    Case "Manage"
        Call main
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub main()

    Call GetClass
    If FoundErr = True Then Exit Sub

    Call ShowJS_Main(ChannelShortName)
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetRootClass() & "</td>"
    Response.Write "  </tr>" & GetChild_Root() & ""
    Response.Write "</table><br>"

    Call ShowContentManagePath(ChannelShortName & "管理")

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_Soft.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>选中</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    Response.Write "            <td align='center' ><strong>" & ChannelShortName & "名称</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>录入</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>下载数</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>" & ChannelShortName & "属性</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>审核状态</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>管理操作</strong></td>"
    Response.Write "          </tr>"

    Dim rsSoftList, sql
    sql = "select S.SoftID,S.ClassID,C.ClassName,C.ParentDir,C.ClassDir,S.SoftName,S.SoftVersion,S.Author,S.Keyword,S.UpdateTime,S.Inputer,S.Editor,S.Hits,S.SoftSize,S.OnTop,S.Elite,S.Status,S.Stars,S.InfoPoint from PE_Soft S"
    sql = sql & " left join PE_Class C on S.ClassID=C.ClassID where S.ChannelID=" & ChannelID & " and S.Deleted=" & PE_False & " and S.Inputer='" & UserName & "' "
    If AuthorName <> "" Then
        sql = sql & " and S.Author='" & AuthorName & "|' "
    End If
    Select Case Status
    Case 3
        sql = sql & " and S.Status=3"
    Case 0
        sql = sql & " and (S.Status=0 Or S.Status=1 Or S.Status=2)"
    Case -1
        sql = sql & " and S.Status=-1"
    Case -2
        sql = sql & " and S.Status=-2"
    End Select
    If ClassID > 0 Then
        If Child > 0 Then
            sql = sql & " and S.ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and S.ClassID=" & ClassID
        End If
    End If

    If Keyword <> "" Then
        Select Case strField
        Case "SoftName"
            sql = sql & " and S.SoftName like '%" & Keyword & "%' "
        Case "SoftIntro"
            sql = sql & " and S.SoftIntro like '%" & Keyword & "%' "
        Case "Author"
            sql = sql & " and S.Author like '%" & Keyword & "%' "
        Case "Inputer"
            sql = sql & " and S.Inputer='" & Keyword & "' "
        Case Else
            sql = sql & " and S.SoftName like '%" & Keyword & "%' "
        End Select
    End If
    sql = sql & " order by S.SoftID desc"

    Set rsSoftList = Server.CreateObject("ADODB.Recordset")
    rsSoftList.Open sql, Conn, 1, 1
    If rsSoftList.BOF And rsSoftList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>" & GetStrNoItem(ClassID, Status) & "<br><br></td></tr>"
    Else
        totalPut = rsSoftList.RecordCount
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
                rsSoftList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim SoftNum
        SoftNum = 0
        Do While Not rsSoftList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'><input name='SoftID' type='checkbox' onclick='unselectall()' id='SoftID' value='" & rsSoftList("SoftID") & "'></td>"
            Response.Write "        <td width='25' align='center'>" & rsSoftList("SoftID") & "</td>"
            Response.Write "        <td>"
            If rsSoftList("ClassID") <> ClassID Then
                Response.Write "<a href='" & FileName & "&ClassID=" & rsSoftList("ClassID") & "'>[" & rsSoftList("ClassName") & "]</a>&nbsp;"
            End If
            Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rsSoftList("SoftID") & "' title='" & GetLinkTips(rsSoftList("SoftName"), rsSoftList("SoftVersion"), rsSoftList("Author"), rsSoftList("UpdateTime"), rsSoftList("Hits"), rsSoftList("Keyword"), rsSoftList("Stars"), rsSoftList("InfoPoint")) & "'>" & rsSoftList("SoftName")
            If rsSoftList("SoftVersion") <> "" Then
                Response.Write "&nbsp;&nbsp;" & rsSoftList("SoftVersion")
            End If
            Response.Write "</a></td>"
            Response.Write "            <td width='60' align='center'><a href='" & FileName & "&field=Inputer&keyword=" & rsSoftList("Inputer") & "' title='点击将查看此用户录入的所有" & ChannelShortName & "'>" & rsSoftList("Inputer") & "</a></td>"
            Response.Write "            <td width='40' align='center'>" & rsSoftList("Hits") & "</td>"
            Response.Write "            <td width='80' align='center'>" & GetInfoProperty(rsSoftList("OnTop"), rsSoftList("Hits"), rsSoftList("Elite")) & "</td>"
            Response.Write "            <td width='60' align='center'>" & GetInfoStatus(rsSoftList("Status")) & "</td>"
            Response.Write "    <td width='80' align='center'>"
            If rsSoftList("Inputer") = UserName And (rsSoftList("Status") <= 0 Or EnableModifyDelete = 1) Then
                Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & rsSoftList("SoftID") & "'>修改</a>&nbsp;"
                Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Del&SoftID=" & rsSoftList("SoftID") & "' onclick=""return confirm('确定要删除此" & ChannelShortName & "吗？一旦删除将不能恢复！');"">删除</a>"
            End If
            Response.Write "</td>"
            Response.Write "</tr>"

            SoftNum = SoftNum + 1
            If SoftNum >= MaxPerPage Then Exit Do
            rsSoftList.MoveNext
        Loop
    End If
    rsSoftList.Close
    Set rsSoftList = Nothing
    Response.Write "</table>"

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中本页显示的所有" & ChannelShortName & "</td><td>"
    Response.Write "<input name='submit1' type='submit' value='删除选定的" & ChannelShortName & "' onClick=""document.myform.Action.value='Del'"" >"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName & "", True)
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>" & ChannelShortName & "搜索：</strong></td>"
    Response.Write "   <td>"
    Response.Write "<table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='SoftName' selected>" & ChannelShortName & "名称</option>"
    Response.Write "<option value='SoftIntro'>" & ChannelShortName & "简介</option>"
    Response.Write "<option value='Author'>" & ChannelShortName & "作者</option>"
    Response.Write "</select>"
    Response.Write "<select name='ClassID'><option value=''>所有栏目</option>" & User_GetClass_Option(1, 0) & "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='搜索'>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></form></table>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "属性中的各项含义：<font color=blue>顶</font>----固顶" & ChannelShortName & "，<font color=red>热</font>----热门" & ChannelShortName & "，<font color=green>荐</font>----推荐" & ChannelShortName & "<br><br>"
End Sub

Function GetStrNoItem(iClassID, iStatus)
    Dim strNoItem
    strNoItem = ""
    If ClassID > 0 Then
        strNoItem = strNoItem & "此栏目及其子栏目中没有任何"
    Else
        strNoItem = strNoItem & "没有任何"
    End If
    Select Case Status
    Case -2
        strNoItem = strNoItem & "未被采用的" & ChannelShortName
    Case -1
        strNoItem = strNoItem & "草稿"
    Case 0
        strNoItem = strNoItem & "<font color=blue>待审核</font>的" & ChannelShortName & "！"
    Case 3
        strNoItem = strNoItem & "<font color=green>已审核</font>的" & ChannelShortName & "！"
    Case Else
        strNoItem = strNoItem & "" & ChannelShortName & "！"
    End Select
    GetStrNoItem = strNoItem
End Function

Function GetLinkTips(SoftName, SoftVersion, Author, UpdateTime, Hits, Keyword, Stars, InfoPoint)
    Dim strLinkTips
    strLinkTips = ""
    strLinkTips = strLinkTips & "名&nbsp;&nbsp;&nbsp;&nbsp;称：" & SoftName & vbCrLf
    strLinkTips = strLinkTips & "版&nbsp;&nbsp;&nbsp;&nbsp;本：" & SoftVersion & vbCrLf
    strLinkTips = strLinkTips & "作&nbsp;&nbsp;&nbsp;&nbsp;者：" & Author & vbCrLf
    strLinkTips = strLinkTips & "更新时间：" & UpdateTime & vbCrLf
    strLinkTips = strLinkTips & "下载次数：" & Hits & vbCrLf
    strLinkTips = strLinkTips & "关 键 字：" & Mid(Keyword, 2, Len(Keyword) - 2) & vbCrLf
    strLinkTips = strLinkTips & "推荐等级："
    If Stars = 0 Then
        strLinkTips = strLinkTips & "无"
    Else
        strLinkTips = strLinkTips & String(Stars, "★")
    End If
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "下载点数：" & InfoPoint
    GetLinkTips = strLinkTips
End Function

Function GetInfoStatus(iStatus)
    Dim strInfoStatus
    Select Case iStatus
    Case -2
        strInfoStatus = "<font color=gray>退稿</font>"
    Case -1
        strInfoStatus = "<font color=gray>草稿</font>"
    Case 0
        strInfoStatus = "<font color=red>待审核</font>"
    Case 1
        strInfoStatus = "<font color=red>一审通过</font>"
    Case 2
        strInfoStatus = "<font color=red>二审通过</font>"
    Case 3
        strInfoStatus = "<font color=black>终审通过</font>"
    End Select
    GetInfoStatus = strInfoStatus
End Function

Function GetInfoProperty(OnTop, Hits, Elite)
    Dim strInfoProperty
    strInfoProperty = ""
    If OnTop = True Then
        strInfoProperty = strInfoProperty & "<font color=blue>顶</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    strInfoProperty = strInfoProperty & "&nbsp;"
    If Hits >= HitsOfHot Then
        strInfoProperty = strInfoProperty & "<font color=red>热</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    strInfoProperty = strInfoProperty & "&nbsp;"
    If Elite = True Then
        strInfoProperty = strInfoProperty & "<font color=green>荐</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    GetInfoProperty = strInfoProperty
End Function

Sub ShowJS_Soft()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddUrl(){" & vbCrLf
    Response.Write "  var thisurl='" & XmlText("Soft", "DownloadUrlTip", "下载地址") & "'+(document.myform.DownloadUrl.length+1)+'|http://'; " & vbCrLf
    Response.Write "  var url=prompt('请输入下载地址名称和链接，中间用“|”隔开：',thisurl);" & vbCrLf
    Response.Write "  if(url!=null&&url!=''){document.myform.DownloadUrl.options[document.myform.DownloadUrl.length]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ModifyUrl(){" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.DownloadUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('请先选择一个下载地址，再点修改按钮！');return false;}" & vbCrLf
    Response.Write "  var url=prompt('请输入下载地址名称和链接，中间用“|”隔开：',thisurl);" & vbCrLf
    Response.Write "  if(url!=thisurl&&url!=null&&url!=''){document.myform.DownloadUrl.options[document.myform.DownloadUrl.selectedIndex]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function DelUrl(){" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.DownloadUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('请先选择一个下载地址，再点删除按钮！');return false;}" & vbCrLf
    Response.Write "  document.myform.DownloadUrl.options[document.myform.DownloadUrl.selectedIndex]=null;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if (document.myform.ClassID.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "所属栏目不能指定为外部栏目！');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.ClassID.value=='0'){" & vbCrLf
    Response.Write "    alert('指定的栏目不允许添加" & ChannelShortName & "！只允许在其子栏目中添加" & ChannelShortName & "。');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.ClassID.value=='-1'){" & vbCrLf
    Response.Write "    alert('您没有在此栏目发表" & ChannelShortName & "的权限，请选择其他栏目！');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.SoftName.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "名称不能为空！');" & vbCrLf
    Response.Write "    document.myform.SoftName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    alert('关键字不能为空！');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.SoftIntro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.SoftIntro.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "简介不能为空！');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "下载地址不能为空！');" & vbCrLf
    Response.Write "    document.myform.DownloadUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.DownloadUrls.value=''" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.DownloadUrl.length;i++){" & vbCrLf
    Response.Write "    if (document.myform.DownloadUrls.value=='') document.myform.DownloadUrls.value=document.myform.DownloadUrl.options[i].value;" & vbCrLf
    Response.Write "    else document.myform.DownloadUrls.value+='$$$'+document.myform.DownloadUrl.options[i].value;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Add()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起！您没有在" & ChannelName & "添加" & ChannelShortName & "的权限！</li><br><br>"
        Exit Sub
    End If
    Dim trs
    If MaxPerDay > 0 Then
        Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)=0")
        If trs(0) >= MaxPerDay Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您今天发表的" & ChannelShortName & "已经达到了上限！</li>"
        End If
        Set trs = Nothing
        If FoundErr = True Then Exit Sub
    End If
    
    Call ShowJS_Soft
    
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Soft.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>添加" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属栏目：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <select name='ClassID'>" & User_GetClass_Option(4, ClassID) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属专题：</strong></td>"
    Response.Write "            <td><select name='SpecialID'><option value='0'>不属于任何专题</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "名称：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftName' type='text' value='' size='50' maxlength='255'> <font color='#FF0000'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>关键字：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Keyword' type='text' id='Keyword' size='50' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("User", ChannelID)
    Response.Write "              <br><font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>作者/开发商：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='50' maxlength='30'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = True Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "版本：</strong></td>"
        Response.Write "            <td><input name='SoftVersion' type='text' size='15' maxlength='100'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "来源：</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "类别：</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <select name='SoftType' id='SoftType'>" & GetSoftType(0) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>" & ChannelShortName & "语言：</strong> <select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(0) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>授权形式：</strong> <select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(0) & "</select>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "平台：</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='OperatingSystem' type='text' value='" & XmlText("Soft", "OperatingSystem", "Win9x/NT/2000/XP/") & "' size='80' maxlength='200'> <br>" & GetOperatingSystemList
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "演示地址：</strong></td>"
        Response.Write "            <td><input name='DemoUrl' type='text' value='http://' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "注册地址：</strong></td>"
        Response.Write "            <td><input name='RegUrl' type='text' value='http://' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>解压密码：</strong></td>"
        Response.Write "            <td><input name='DecompressPassword' type='text' id='DecompressPassword' size='30' maxlength='30'></td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "图片：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftPicUrl' type='text' id='SoftPicUrl' size='80' maxlength='200'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'></td>"
    Response.Write "            <td><table><tr><td>上传" & ChannelShortName & "图片：</td><td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=softpic' frameborder=0 scrolling=no width='450' height='25'></iframe></td></tr></table></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5' valign='middle'><strong>" & ChannelShortName & "简介：</strong></td>"
    Response.Write "            <td><textarea name='SoftIntro' cols='80' rows='10' id='SoftIntro' style='display:none'></textarea>"
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=SoftIntro' frameborder='1' scrolling='no' width='650' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>上传" & ChannelShortName & "：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <iframe style='top:2px' ID='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=soft' frameborder=0 scrolling=no width='450' height='25'></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"	
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "地址：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='hidden' name='DownloadUrls' value=''>"
    Response.Write "                    <select name='DownloadUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'></select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='addurl' value='添加外部地址' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='修改当前地址' onclick='return ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='删除当前地址' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              <tr><td  colspan='3'>系统提供的上传功能只适合上传比较小的" & ChannelShortName & "（如ASP源代码压缩包）。如果" & ChannelShortName & "比较大（" & MaxFileSize \ 1024 & "M以上），请先使用FTP上传，而不要使用系统提供的上传功能，以免上传出错或过度占用服务器的CPU资源。FTP上传后请将地址复制到下面的地址框中。</td></tr>"		
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "大小：</strong></td>"
    Response.Write "            <td><input name='SoftSize' type='text' id='SoftSize' size='10' maxlength='10'> K</strong></td>"
    Response.Write "          </tr>"
    
    '自定义字段
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-2")
    Do While Not rsField.EOF
        IF rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsField("DefaultValue"), rsField("Options"), rsField("EnableNull"))
        End If
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg' class='tdbg5'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "状态：</strong></td>"
    Response.Write "            <td><input name='Status' type='radio' id='Status' value='-1'>草稿&nbsp;&nbsp;<input Name='Status' Type='Radio' Id='Status' Value='0' checked>投稿</td>"
    Response.Write "          </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 添 加 ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' 预 览 ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub Modify()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起！您没有在" & ChannelName & "添加" & ChannelShortName & "的权限！</li><br><br>"
        Exit Sub
    End If
    Dim rsSoft, sql, tmpAuthor, tmpCopyFrom, SpecialID
    
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        SoftID = PE_CLng(SoftID)
    End If
    sql = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and SoftID=" & SoftID & ""
    Set rsSoft = Server.CreateObject("ADODB.Recordset")
    rsSoft.Open sql, Conn, 1, 1
    If rsSoft.BOF And rsSoft.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "</li>"
    Else
        If rsSoft("Status") > 0 And EnableModifyDelete = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "已经被审核通过或在审核中，您不能再进行修改！</li>"
        End If
    End If
    If FoundErr = True Then
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    End If
    SpecialID = PE_CLng(Conn.Execute("select top 1 SpecialID from PE_InfoS where ModuleType=2 and ItemID=" & SoftID & "")(0))

    If Right(rsSoft("Author"), 1) = "|" Then
        tmpAuthor = Left(rsSoft("Author"), Len(rsSoft("Author")) - 1)
    Else
        tmpAuthor = rsSoft("Author")
    End If
    If Right(rsSoft("CopyFrom"), 1) = "|" Then
        tmpCopyFrom = Left(rsSoft("CopyFrom"), Len(rsSoft("CopyFrom")) - 1)
    Else
        tmpCopyFrom = rsSoft("CopyFrom")
    End If
    Call ShowJS_Soft

    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Soft.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>修改" & ChannelShortName & "</b></td>"
    Response.Write "    </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属栏目：</strong></td>"
    Response.Write "            <td colspan='2'>"
    Response.Write "              <select name='ClassID'>" & User_GetClass_Option(4, rsSoft("ClassID")) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属专题：</strong></td>"
    Response.Write "            <td><select name='SpecialID'><option value='0'>不属于任何专题</option>" & GetSpecial_Option(SpecialID) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "名称：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftName' type='text' value='" & rsSoft("SoftName") & "' size='50' maxlength='255'> <font color='#FF0000'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>关键字：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Keyword' type='text' id='Keyword' value='" & Mid(rsSoft("Keyword"), 2, Len(rsSoft("Keyword")) - 2) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("User", ChannelID)
    Response.Write "              <br><font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>作者/开发商：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & tmpAuthor & "' size='50' maxlength='30'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = True Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "版本：</strong></td>"
        Response.Write "            <td><input name='SoftVersion' type='text' size='15' maxlength='100' value='" & rsSoft("SoftVersion") & "'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "来源：</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "类别：</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <select name='SoftType' id='SoftType'>" & GetSoftType(rsSoft("SoftType")) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>" & ChannelShortName & "语言：</strong> <select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(rsSoft("SoftLanguage")) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>授权形式：</strong> <select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(rsSoft("CopyrightType")) & "</select>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "平台：</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='OperatingSystem' type='text' value='" & rsSoft("OperatingSystem") & "' size='80' maxlength='200'> <br>" & GetOperatingSystemList
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "演示地址：</strong></td>"
        Response.Write "            <td><input name='DemoUrl' type='text' value='" & rsSoft("DemoUrl") & "' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "注册地址：</strong></td>"
        Response.Write "            <td><input name='RegUrl' type='text' value='" & rsSoft("RegUrl") & "' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>解压密码：</strong></td>"
        Response.Write "            <td><input name='DecompressPassword' type='text' id='DecompressPassword' value='" & rsSoft("DecompressPassword") & "' size='30' maxlength='30'></td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "图片：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftPicUrl' type='text' id='SoftPicUrl' value='" & rsSoft("SoftPicUrl") & "' size='80' maxlength='200'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'></td>"
    Response.Write "            <td><table><tr><td>上传" & ChannelShortName & "图片：</td><td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=softpic' frameborder=0 scrolling=no width='450' height='25'></iframe></td></tr></table></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5' valign='middle'><strong>" & ChannelShortName & "简介：</strong></td>"
    Response.Write "            <td><textarea name='SoftIntro' cols='80' rows='10' id='SoftIntro' style='display:none'>" & Server.HTMLEncode(FilterJS(rsSoft("SoftIntro"))) & "</textarea>"
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=SoftIntro' frameborder='1' scrolling='no' width='650' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>上传" & ChannelShortName & "：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <iframe style='top:2px' ID='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=soft' frameborder=0 scrolling=no width='450' height='25'></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"	
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "地址：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='hidden' name='DownloadUrls' value=''>"
    Response.Write "                    <select name='DownloadUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'>"
    Dim DownloadUrls, arrDownloadUrls, iTemp
    DownloadUrls = rsSoft("DownloadUrl")
    If InStr(DownloadUrls, "$$$") > 1 Then
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        For iTemp = 0 To UBound(arrDownloadUrls)
            Response.Write "<option value='" & arrDownloadUrls(iTemp) & "'>" & arrDownloadUrls(iTemp) & "</option>"
        Next
    Else
        Response.Write "<option value='" & DownloadUrls & "'>" & DownloadUrls & "</option>"
    End If
    Response.Write "                    </select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='addurl' value='添加外部地址' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='修改当前地址' onclick='return ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='删除当前地址' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
        Response.Write "              <tr><td  colspan='3'>系统提供的上传功能只适合上传比较小的" & ChannelShortName & "（如ASP源代码压缩包）。如果" & ChannelShortName & "比较大（" & MaxFileSize \ 1024 & "M以上），请先使用FTP上传，而不要使用系统提供的上传功能，以免上传出错或过度占用服务器的CPU资源。FTP上传后请将地址复制到下面的地址框中。</td></tr>"		
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "大小：</strong></td>"
    Response.Write "            <td><input name='SoftSize' type='text' id='SoftSize' value='" & rsSoft("SoftSize") & "' size='10' maxlength='10'> K</td>"
    Response.Write "          </tr>"
    '自定义字段
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-2")
    Do While Not rsField.EOF
        If rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsSoft(Trim(rsField("FieldName"))), rsField("Options"), rsField("EnableNull"))
        End If
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "状态：</td>"
    Response.Write "            <td>"
    If rsSoft("Status") <= 0 Then
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='-1'"
        If rsSoft("Status") = -1 Then
            Response.Write " checked"
        End If
        Response.Write "> 草稿&nbsp;&nbsp;"
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='0'"
        If rsSoft("Status") = 0 Then
            Response.Write "checked"
        End If
        Response.Write "> 投稿"
    Else
        If rsSoft("Status") < 3 Then
            Response.Write "审核中"
        Else
            Response.Write "已经发布"
        End If
    End If
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "   <input name='SoftID' type='hidden' id='SoftID' value='" & rsSoft("SoftID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='保存修改结果' style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsSoft.Close
    Set rsSoft = Nothing

End Sub

Sub WriteFieldHTML(FieldName, Title, Tips, FieldType, strValue, Options, EnableNull)
    Dim FieldUpload, ChannelUpload, UserUpload,rsFieldUpload,sqlFieldUpload   
    Select Case FieldType
    Case 4,5
        FieldUpload = True		
        ChannelUpload = Conn.Execute("Select EnableUploadFile from PE_Channel where ChannelID="&ChannelID)(0) 
        If  ChannelUpload = False Then FieldUpload = False
        If UserName<>"" Then   
            sqlFieldUpload = "SELECT U.UserID,U.SpecialPermission,U.UserSetting,G.GroupSetting FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
            sqlFieldUpload = sqlFieldUpload & " UserName='" & UserName & "'" 
            Set rsFieldUpload = Conn.Execute(sqlFieldUpload)
            If rsFieldUpload.BOF And rsFieldUpload.EOF Then
                FieldUpload = False
            Else
                If rsFieldUpload("SpecialPermission") = True Then
                    UserSetting = Split(Trim(rsFieldUpload("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                Else
                    UserSetting = Split(Trim(rsFieldUpload("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                End If
                If CBool(PE_CLng(UserSetting(9))) = False Then
                    FieldUpload = False
                End If
            End If
            Set rsFieldUpload = nothing			 
        End If			               			
    End Select	
    Dim strEnableNull
    If EnableNull = False Then
        strEnableNull = " <font color='#FF0000'>*</font>"
    End If
    Response.Write "<tr class='tdbg'><td width='120' align='right' class='tdbg5'><b>" & Title & "：</b></td><td colspan='5'>"
    Select Case FieldType
    Case 1 ,8   '单行文本框
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2 ,9   '多行文本框
        Response.Write "<textarea name='" & FieldName & "' cols='80' rows='10'>" & strValue & "</textarea>" & strEnableNull
    Case 3    '下拉列表
        Response.Write "<select name='" & FieldName & "'>"
        Dim arrOptions, i
        arrOptions = Split(Options, vbCrLf)
        For i = 0 To UBound(arrOptions)
            Response.Write "<option value='" & arrOptions(i) & "'"
            If arrOptions(i) = strValue Then Response.Write " selected"
            Response.Write ">" & arrOptions(i) & "</option>"
        Next
        Response.Write "</select>" & strEnableNull
    Case 4   '图片  					
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
        End If
        If PE_CBool(FieldUpload) = True Then		
            Response.Write "<iframe style='top:2px;' id='uploadPhoto' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldpic&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"
        End If				
    Case 5   '文件
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
        End If
        If PE_CBool(FieldUpload) = True Then			
            Response.Write "<iframe style='top:2px' id='uploadsoft' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldsoft&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"	
        End If
    Case 6    '日期
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & Now() & "'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & strValue & "'>" & strEnableNull
        End If
    Case 7    '数字
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "'  onkeyup=""value=value.replace(/[^\d]/g,'')"" size='20' maxlength='20' value='0'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' onkeyup=""value=value.replace(/[^\d]/g,'')"" size='20' maxlength='20' value='" & PE_Clng(strValue) & "'>" & strEnableNull
        End If		
    End Select
    If IsNull(Tips) = False And Tips <> "" Then
        Response.Write "<br>" & PE_HTMLEncode(Tips)
    End If
    Response.Write "</td></tr>"
End Sub

Sub SaveSoft()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起！您没有在" & ChannelName & "添加" & ChannelShortName & "的权限！</li><br><br>"
        Exit Sub
    End If
    Dim rsSoft, sql
    Dim trs, tAuthor
    Dim SoftID, ClassID, SpecialID, SoftName, SoftVersion, SoftType, SoftLanguage, CopyrightType, OperatingSystem, Author, CopyFrom
    Dim DemoUrl, RegUrl, SoftPicUrl, SoftIntro, Keyword, DecompressPassword, SoftSize, DownloadUrls, Inputer
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent

    
    SoftID = PE_CLng(Trim(Request.Form("SoftID")))
    ClassID = PE_CLng(Trim(Request.Form("ClassID")))
    SpecialID = PE_CLng(Trim(Request.Form("SpecialID")))
    SoftName = Trim(Request.Form("SoftName"))
    SoftVersion = Trim(Request.Form("SoftVersion"))
    Keyword = Trim(Request.Form("Keyword"))
    SoftType = PE_HTMLEncode(Trim(Request.Form("SoftType")))
    SoftLanguage = PE_HTMLEncode(Trim(Request.Form("SoftLanguage")))
    CopyrightType = PE_HTMLEncode(Trim(Request.Form("CopyrightType")))
    OperatingSystem = PE_HTMLEncode(Trim(Request.Form("OperatingSystem")))
    Author = PE_HTMLEncode(Trim(Request.Form("Author")))
    CopyFrom = PE_HTMLEncode(Trim(Request.Form("CopyFrom")))
    DemoUrl = PE_HTMLEncode(Trim(Request.Form("DemoUrl")))
    RegUrl = PE_HTMLEncode(Trim(Request.Form("RegUrl")))
    SoftPicUrl = PE_HTMLEncode(Trim(Request.Form("SoftPicUrl")))
    SoftIntro = ReplaceBadUrl(FilterJS(Trim(Request.Form("SoftIntro"))))
    DecompressPassword = PE_HTMLEncode(Trim(Request.Form("DecompressPassword")))
    SoftSize = PE_CLng(Trim(Request.Form("SoftSize")))
    DownloadUrls = PE_HTMLEncode(Trim(Request.Form("DownloadUrls")))
    Inputer = UserName
    Status = PE_CLng(Trim(Request.Form("Status")))

    If ClassID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>未指定所属栏目，或者指定的栏目不允许此操作！</li>"
    Else
        Dim tClass
        Set tClass = Conn.Execute("select ClassName,ClassType,Depth,ParentID,ParentPath,Child,EnableAdd,PresentExp,DefaultItemPoint,DefaultItemChargeType,DefaultItemPitchTime,DefaultItemReadTimes,DefaultItemDividePercent from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
        Else
            ClassName = tClass("ClassName")
            Depth = tClass("Depth")
            ParentPath = tClass("ParentPath")
            ParentID = tClass("ParentID")
            Child = tClass("Child")
            PresentExp = tClass("PresentExp")
            DefaultItemPoint = tClass("DefaultItemPoint")
            DefaultItemChargeType = tClass("DefaultItemChargeType")
            DefaultItemPitchTime = tClass("DefaultItemPitchTime")
            DefaultItemReadTimes = tClass("DefaultItemReadTimes")
            DefaultItemDividePercent = tClass("DefaultItemDividePercent")

            If Child > 0 And tClass("EnableAdd") = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>指定的栏目不允许添加" & ChannelShortName & "</li>"
            End If
            If tClass("ClassType") = 2 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>不能指定为外部栏目</li>"
            End If
            Dim CheckParentPath
            If ParentID > 0 Then
                CheckParentPath = ChannelDir & "all," & ParentPath & "," & ClassID
            Else
                CheckParentPath = ChannelDir & "all," & ClassID
            End If
            If CheckPurview_Class(arrClass_Input, CheckParentPath) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对不起，你没有此栏目的相应操作权限！</li>"
            End If
        End If
        Set tClass = Nothing
    End If
    
    If SoftName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "名称不能为空</li>"
    End If

    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "佚名")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "本站原创")
    Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入" & ChannelShortName & "关键字</li>"
    Else
        Call SaveKeyword(Keyword)
    End If
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = True Then
        If SoftType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "类别不能为空</li>"
        End If
        If SoftLanguage = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "语言不能为空</li>"
        End If
        If CopyrightType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>授权形式不能为空</li>"
        End If
        If OperatingSystem = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "平台不能为空</li>"
        End If
    End If
    If DownloadUrls = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "下载地址不能为空</li>"
    End If

    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-2")
    Do While Not rsField.EOF
        If rsField("EnableNull") = False Then
            If Trim(Request(rsField("FieldName"))) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请输入" & rsField("Title") & "！</li>"
            End If
        End If
        rsField.MoveNext
    Loop
    
    If FoundErr = True Then
        Exit Sub
    End If

    If Status < 0 Then
        Status = -1
    Else
        If CheckLevel = 0 Or NeedlessCheck = 1 Then
            Status = 3
        Else
            Status = 0
        End If
    End If
    
    SoftName = PE_HTMLEncode(SoftName)
    SoftVersion = PE_HTMLEncode(SoftVersion)
    SoftType = PE_HTMLEncode(SoftType)
    SoftLanguage = PE_HTMLEncode(SoftLanguage)
    CopyrightType = PE_HTMLEncode(CopyrightType)
    OperatingSystem = PE_HTMLEncode(OperatingSystem)
    DemoUrl = PE_HTMLEncode(DemoUrl)
    RegUrl = PE_HTMLEncode(RegUrl)
    SoftPicUrl = PE_HTMLEncode(SoftPicUrl)
    Keyword = "|" & ReplaceBadChar(Keyword) & "|"
    DecompressPassword = PE_HTMLEncode(DecompressPassword)

    Set rsSoft = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("SoftName") = SoftName And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>请不要重复添加同一" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("SoftName") = SoftName
            Session("AddTime") = Now()
            If MaxPerDay > 0 Then
                Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Inputer='" & UserName & "' and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")=0")
                If trs(0) >= MaxPerDay Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>您今天发表的" & ChannelShortName & "已经达到了上限！</li>"
                End If
                Set trs = Nothing
                If FoundErr = True Then Exit Sub
            End If
            
            sql = "select top 1 * from PE_Soft"
            rsSoft.Open sql, Conn, 1, 3
            rsSoft.addnew
            SoftID = PE_CLng(Conn.Execute("select max(SoftID) from PE_Soft")(0)) + 1
            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (2," & SoftID & "," & SpecialID & ")")
            rsSoft("SoftID") = SoftID
            rsSoft("ChannelID") = ChannelID
            rsSoft("ClassID") = ClassID
            rsSoft("SoftName") = SoftName
            rsSoft("SoftVersion") = SoftVersion
            rsSoft("SoftType") = SoftType
            rsSoft("SoftLanguage") = SoftLanguage
            rsSoft("CopyrightType") = CopyrightType
            rsSoft("OperatingSystem") = OperatingSystem
            rsSoft("Author") = Author
            rsSoft("CopyFrom") = CopyFrom
            rsSoft("DemoUrl") = DemoUrl
            rsSoft("RegUrl") = RegUrl
            rsSoft("SoftPicUrl") = SoftPicUrl
            rsSoft("SoftIntro") = SoftIntro
            rsSoft("Keyword") = Keyword
            rsSoft("Hits") = 0
            rsSoft("DayHits") = 0
            rsSoft("WeekHits") = 0
            rsSoft("MonthHits") = 0
            rsSoft("Stars") = 0
            rsSoft("UpdateTime") = Now()
            rsSoft("Status") = Status
            rsSoft("OnTop") = False
            rsSoft("Elite") = False
            rsSoft("DecompressPassword") = DecompressPassword
            rsSoft("SoftSize") = SoftSize
            rsSoft("DownloadUrl") = DownloadUrls
            rsSoft("Inputer") = Inputer
            rsSoft("Editor") = Inputer
            rsSoft("SkinID") = 0
            rsSoft("TemplateID") = 0
            rsSoft("Deleted") = False
            PresentExp = CLng(PresentExp * PresentExpTimes)
            rsSoft("PresentExp") = PresentExp
            rsSoft("InfoPoint") = DefaultItemPoint
            rsSoft("VoteID") = 0
            rsSoft("InfoPurview") = 0
            rsSoft("arrGroupID") = ""
            rsSoft("ChargeType") = DefaultItemChargeType
            rsSoft("PitchTime") = DefaultItemPitchTime
            rsSoft("ReadTimes") = DefaultItemReadTimes
            rsSoft("DividePercent") = DefaultItemDividePercent
            
            If Not (rsField.BOF And rsField.EOF) Then
                rsField.MoveFirst
                Do While Not rsField.EOF
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rsSoft(Trim(rsField("FieldName"))) = FilterJS(Trim(Request(rsField("FieldName"))))
                    End If
                    rsField.MoveNext
                Loop
            End If
            Set rsField = Nothing

            If BlogFlag = True Then '写入BLOGID
                Dim blogid
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsSoft("BlogID") = 0
                Else
                    rsSoft("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If

            rsSoft.Update
            If CheckLevel = 0 Or NeedlessCheck = 1 Then
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1,ItemChecked=ItemChecked+1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_Class set ItemCount=ItemCount+1 where ClassID=" & ClassID & "")
                If rsSoft("Status") = 3 Then
                    Conn.Execute ("update PE_User set PostItems=PostItems+1,PassedItems=PassedItems+1,UserExp=UserExp+" & PresentExp & " where UserName='" & UserName & "'")
                End If
            Else
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_User set PostItems=PostItems+1 where UserName='" & UserName & "'")
            End If
        End If
    ElseIf Action = "SaveModify" Then
        If SoftID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定" & ChannelShortName & "ID的值</li>"
        Else
            sql = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and SoftID=" & SoftID
            rsSoft.Open sql, Conn, 1, 3
            If rsSoft.BOF And rsSoft.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到此" & ChannelShortName & "，可能已经被其他人删除。</li>"
            Else
                If rsSoft("Status") > 0 And EnableModifyDelete = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & ChannelShortName & "已经被审核通过，您不能再进行修改！</li>"
                Else
                    Conn.Execute ("delete from PE_InfoS where ModuleType=2 and ItemID=" & SoftID)
                    Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (2," & SoftID & "," & SpecialID & ")")
                    rsSoft("ClassID") = ClassID
                    rsSoft("SoftName") = SoftName
                    rsSoft("SoftVersion") = SoftVersion
                    rsSoft("SoftType") = SoftType
                    rsSoft("SoftLanguage") = SoftLanguage
                    rsSoft("CopyrightType") = CopyrightType
                    rsSoft("OperatingSystem") = OperatingSystem
                    rsSoft("Author") = Author
                    rsSoft("CopyFrom") = CopyFrom
                    rsSoft("DemoUrl") = DemoUrl
                    rsSoft("RegUrl") = RegUrl
                    rsSoft("SoftPicUrl") = SoftPicUrl
                    rsSoft("SoftIntro") = SoftIntro
                    rsSoft("Keyword") = Keyword
                    rsSoft("UpdateTime") = Now()
                    rsSoft("DecompressPassword") = DecompressPassword
                    rsSoft("SoftSize") = SoftSize
                    rsSoft("DownloadUrl") = DownloadUrls
                    rsSoft("Status") = Status
                    
                    If Not (rsField.BOF And rsField.EOF) Then
                        rsField.MoveFirst
                        Do While Not rsField.EOF
                            If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                                rsSoft(Trim(rsField("FieldName"))) = PE_HTMLEncode(FilterJS(Trim(Request(rsField("FieldName")))))
                            End If
                            rsField.MoveNext
                        Loop
                    End If
                    Set rsField = Nothing

                    rsSoft.Update
                End If
            End If
        End If
    End If
    rsSoft.Close
    Set rsSoft = Nothing
    
    If FoundErr = True Then Exit Sub
    
    Response.Write "<br><br>"
    Response.Write "<table class='border' align=center width='400' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
    Response.Write "  <tr align=center> "
    Response.Write "    <td  height='22' align='center' class='title'> "
    If Action = "SaveAdd" Then
        Response.Write "<b>添加" & ChannelShortName & "成功</b>"
    Else
        Response.Write "<b>修改" & ChannelShortName & "成功</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    If Status = 0 Then
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td height='60'><font color='#0000FF'>注意：</font><br>&nbsp;&nbsp;&nbsp;&nbsp;您的" & ChannelShortName & "尚未真正发表！只有等管理员审核并通过了您的" & ChannelShortName & "后，您所添加的" & ChannelShortName & "才会发表。</td>"
        Response.Write "  </tr>"
    End If
    Response.Write "  <tr>"
    Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>所属栏目：</strong></td>"
    Response.Write "          <td>" & ShowClassPath() & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "名称：</strong></td>"
    Response.Write "          <td>" & SoftName & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "版本：</strong></td>"
    Response.Write "          <td>" & SoftVersion & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "作者：</strong></td>"
    Response.Write "          <td>" & Author & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'> "
    Response.Write "          <td width='100' align='right'><strong>关 键 字：</strong></td>"
    Response.Write "          <td>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='30' align='center'>"
    Response.Write "【<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & SoftID & "'>修改此" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>继续添加" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "管理</a>】&nbsp;"
    Response.Write "【<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & SoftID & "'>预览" & ChannelShortName & "内容</a>】"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Call ClearSiteCache(0)
    Call CreateAllJS_User
End Sub

Sub Del()
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If

    Dim sqlDel, rsDel, NeedUpdateCache
    NeedUpdateCache = False

    sqlDel = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and "
    If InStr(SoftID, ",") > 0 Then
        sqlDel = sqlDel & " SoftID in (" & SoftID & ") order by SoftID"
    Else
        sqlDel = sqlDel & " SoftID=" & SoftID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If rsDel("Status") > 0 Then
            If EnableModifyDelete = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>删除" & ChannelShortName & "“" & rsDel("SoftName") & "”失败。原因：此" & ChannelShortName & "已经被审核通过，您不能再删除！</li>"
            Else
                Conn.Execute ("update PE_User set PostItems=PostItems-1,PassedItems=PassedItems-1,UserExp=UserExp-" & rsDel("PresentExp") & " where UserName='" & UserName & "'")
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount-1,ItemChecked=ItemChecked-1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_Class set ItemCount=ItemCount-1 where ClassID=" & rsDel("ClassID") & "")
                rsDel("Deleted") = True
                rsDel.Update
                NeedUpdateCache = True
            End If
        Else
            Conn.Execute ("update PE_Channel set ItemCount=ItemCount-1 where ChannelID=" & ChannelID & "")
            Conn.Execute ("update PE_User set PostItems=PostItems-1 where UserName='" & UserName & "'")
            rsDel("Deleted") = True
            rsDel.Update
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    
    If NeedUpdateCache = True Then
        Call ClearSiteCache(0)
        Call CreateAllJS_User
    End If

    Call CloseConn
    If FoundErr = False Then
        Response.Redirect ComeUrl
    End If
End Sub

Sub Show()
    Dim rs, sql
    SoftID = PE_CLng(SoftID)
    sql = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and SoftID=" & SoftID & ""
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "</li>"
    Else
        ClassID = rs("ClassID")
        Call GetClass
    End If
    If FoundErr = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "<tr class='title'>"
    Response.Write "  <td height='22' colspan='4'>"
    Response.Write "您现在的位置：&nbsp;<a href='User_Soft.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "管理</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rs("SoftID") & "'>" & PE_HTMLEncode(rs("SoftName")) & "</a>"
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "名称：</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(rs("SoftName")) & "&nbsp;&nbsp;" & PE_HTMLEncode(rs("SoftVersion")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>文件大小：</td>"
    Response.Write "  <td width='200'>" & rs("SoftSize") & " K" & "</td>"
    Response.Write "  <td colspan='2' rowspan='7' align=center valign='middle'>"
    If rs("SoftPicUrl") = "" Then
        Response.Write "相关图片"
    Else
        Response.Write "<img src='" & GetSoftPicUrl(rs("SoftPicUrl")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>运行环境：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("OperatingSystem")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "类别：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("SoftType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "语言：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("SoftLanguage")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>授权方式：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("CopyrightType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>评分等级：</td>"
    Response.Write "  <td width='200'>" & String(rs("Stars"), "★") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>解压密码：</td>"
    Response.Write "  <td width='200'>" & rs("DecompressPassword") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>添加时间：</td>"
    Response.Write "  <td width='200'>" & rs("UpdateTime") & "</td>"
    Response.Write "  <td width='100' align=right valign='middle'>开 发 商：</td>"
    Response.Write "  <td valign='middle'>" & PE_HTMLEncode(rs("Author")) & ""
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align=right valign='middle'>下载点数：</td>"
    Response.Write "  <td width='200'><font color=red> " & rs("InfoPoint") & "</font> 点</td>"
    Response.Write "  <td></td>"
    Response.Write "  <td></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "添加：</td>"
    Response.Write "  <td width='200'>" & rs("Inputer") & "</td>"
    Response.Write "  <td width='100' align=right valign='middle'>责任编辑：</td>"
    Response.Write "  <td valign='middle'>"
    If rs("Status") > 0 Then
        Response.Write rs("Editor")
    Else
        Response.Write "无"
    End If
    Response.Write "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>相关链接：</td>"
    Response.Write "  <td colspan='3'><a href='" & rs("DemoUrl") & "' target='_blank'>" & ChannelShortName & "演示地址</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & rs("RegUrl") & "' target='_blank'>" & ChannelShortName & "注册地址</a></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>下载次数：</td>"
    Response.Write "  <td colspan='3'>本日：" & rs("DayHits") & "&nbsp;&nbsp;&nbsp;&nbsp;本周：" & rs("WeekHits") & "&nbsp;&nbsp;&nbsp;&nbsp;本月：" & rs("MonthHits") & "&nbsp;&nbsp;&nbsp;&nbsp;总计：" & rs("Hits")
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>下载地址：</td>"
    Response.Write "  <td colspan='3'>" & ShowDownloadUrls(rs("DownloadUrl")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td align='right'>&nbsp;</td>"
    Response.Write "  <td colspan='3' align='right'>"
    Response.Write "<strong>可用操作：</strong>"
    If rs("Inputer") = UserName And (rs("Status") <= 0 Or UserSetting(2) = 1) Then
        Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & rs("SoftID") & "'>修改</a>&nbsp;&nbsp;"
        Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Del&SoftID=" & rs("SoftID") & "' onclick=""return confirm('确定要删除此" & ChannelShortName & "吗？');"">删除</a>&nbsp;&nbsp;"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "简介：</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterJS(rs("SoftIntro")) & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub Preview()
    Response.Write "<br><table width='100%' border=0 align=center cellPadding=2 cellSpacing=1 bgcolor='#FFFFFF' class='border' style='WORD-BREAK: break-all'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4'>"

    If ClassID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定所属栏目</li>"
        Exit Sub
    End If

    Call GetClass
    If FoundErr = True Then Exit Sub

    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write ClassName & "&nbsp;&gt;&gt;&nbsp;"

    Response.Write PE_HTMLEncode(Request("SoftName"))
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "名称：</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(Request("SoftName")) & "&nbsp;&nbsp;" & PE_HTMLEncode(Request("SoftVersion")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>文件大小：</td>"
    Response.Write "  <td width='200'>" & Request("SoftSize") & " K" & "</td>"
    Response.Write "  <td colspan='2' rowspan='7' align=center valign='middle'>"
    If Request("SoftPicUrl") = "" Then
        Response.Write "相关图片"
    Else
        Response.Write "<img src='" & GetSoftPicUrl(Request("SoftPicUrl")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>运行环境：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("OperatingSystem")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "类别：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("SoftType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "语言：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("SoftLanguage")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>授权方式：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("CopyrightType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>评分等级：</td>"
    Response.Write "  <td width='200'>" & String(Request("Stars"), "★") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>解压密码：</td>"
    Response.Write "  <td width='200'>" & Request("DecompressPassword") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>添加时间：</td>"
    Response.Write "  <td width='200'>" & Now() & "</td>"
    Response.Write "  <td width='100' align=right valign='middle'>开 发 商：</td>"
    Response.Write "  <td valign='middle'>" & PE_HTMLEncode(Request("Author")) & ""
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>相关链接：</td>"
    Response.Write "  <td colspan='3'><a href='" & Request("DemoUrl") & "' target='_blank'>" & ChannelShortName & "演示地址</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & Request("RegUrl") & "' target='_blank'>" & ChannelShortName & "注册地址</a></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "地址：</td>"
    Response.Write "  <td colspan='3'>" & ShowDownloadUrls(Request("DownloadUrls")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "简介：</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterJS(Request("SoftIntro")) & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>【<a href='javascript:window.close();'>关闭窗口</a>】</p>"
End Sub

Function GetSoftPicUrl(SoftPicUrl)
    If Left(SoftPicUrl, Len("UploadSoftPic")) = "UploadSoftPic" Then
        GetSoftPicUrl = strInstallDir & ChannelDir & "/" & SoftPicUrl
    Else
        GetSoftPicUrl = SoftPicUrl
    End If
End Function

Function ShowDownloadUrls(DownloadUrls)
    Dim arrDownloadUrls, arrUrls, iTemp, strUrls
    strUrls = ""
    arrDownloadUrls = Split(DownloadUrls, "$$$")
    For iTemp = 0 To UBound(arrDownloadUrls)
        arrUrls = Split(arrDownloadUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                strUrls = strUrls & arrUrls(0) & "：<a href='" & strInstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            Else
                strUrls = strUrls & arrUrls(0) & "：<a href='" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            End If
        End If
    Next
    ShowDownloadUrls = strUrls
End Function

Function GetSoftType(SoftType)
    If IsArray(arrSoftType) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrSoftType)
        If Trim(arrSoftType(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrSoftType(i) & "'"
            If Trim(SoftType) = arrSoftType(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrSoftType(i) & "</option>"
        End If
    Next
    GetSoftType = strTemp
End Function

Function GetSoftLanguage(SoftLanguage)
    If IsArray(arrSoftLanguage) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrSoftLanguage)
        If Trim(arrSoftLanguage(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrSoftLanguage(i) & "'"
            If Trim(SoftLanguage) = arrSoftLanguage(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrSoftLanguage(i) & "</option>"
        End If
    Next
    GetSoftLanguage = strTemp
End Function

Function GetCopyrightType(CopyrightType)
    If IsArray(arrCopyrightType) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrCopyrightType)
        If Trim(arrCopyrightType(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrCopyrightType(i) & "'"
            If Trim(CopyrightType) = arrCopyrightType(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrCopyrightType(i) & "</option>"
        End If
    Next
    GetCopyrightType = strTemp
End Function

Function GetOperatingSystemList()
    Dim strOperatingSystemList, i
    
    strOperatingSystemList = "<script language = 'JavaScript'>" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "function ToSystem(addTitle){" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    var str=document.myform.OperatingSystem.value;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    if (document.myform.OperatingSystem.value=="""") {" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    }else{" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        if (str.substr(str.length-1,1)==""" & XmlText("Soft", "OperatingSystemEmblem", "/") & """){" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "            document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        }else{" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "            document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+""" & XmlText("Soft", "OperatingSystemEmblem", "/") & """+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        }" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    }" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    document.myform.OperatingSystem.focus();" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "}" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "</script>" & vbCrLf

    strOperatingSystemList = strOperatingSystemList & "<font color='#808080'>平台选择："
    If IsArray(arrOperatingSystem) Then
        For i = 0 To UBound(arrOperatingSystem)
            If Trim(arrOperatingSystem(i)) <> "" Then
                strOperatingSystemList = strOperatingSystemList & "<a href=""javascript:ToSystem('" & arrOperatingSystem(i) & "')"">" & arrOperatingSystem(i) & "</a>" & XmlText("Soft", "OperatingSystemEmblem", "/")
            End If
        Next
    End If
    strOperatingSystemList = strOperatingSystemList & "</font>"
    GetOperatingSystemList = strOperatingSystemList
End Function
%>
