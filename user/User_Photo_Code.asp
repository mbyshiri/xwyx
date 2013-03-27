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

Dim PhotoID, AuthorName, Status, ManageType
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created

Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview

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
    PhotoID = Trim(Request("PhotoID"))
    ClassID = PE_CLng(Trim(Request("ClassID")))
    Status = Trim(Request("Status"))
    AuthorName = Trim(Request("AuthorName"))
    strField = Trim(Request("Field"))
    If Status = "" Then
        Status = 9
    Else
        Status = PE_CLng(Status)
    End If
    If IsValidID(PhotoID) = False Then
        PhotoID = ""
    End If

    If Action = "" Then Action = "Manage"
    FileName = "User_Photo.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
    strFileName = FileName & "&ClassID=" & ClassID & "&Field=" & strField & "&keyword=" & Keyword
    If AuthorName <> "" Then
        AuthorName = ReplaceBadChar(AuthorName)
        strFileName = strFileName & "&AuthorName=" & AuthorName
    End If

    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

    Response.Write "<table align='center'><tr align='center' valign='top'>"
    If CheckUser_ChannelInput() Then
        Response.Write "<td width='90'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Add'><img src='images/Photo_add.gif' border='0' align='absmiddle'><br>添加" & ChannelShortName & "</a></td>"
    End If
    Response.Write "<td width='90'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Status=9'><img src='images/Photo_all.gif' border='0' align='absmiddle'><br>所有" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Status=-1'><img src='images/Photo_draft.gif' border='0' align='absmiddle'><br>草 稿</a></td>"
    Response.Write "<td width='90'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Status=0'><img src='images/Photo_unpassed.gif' border='0' align='absmiddle'><br>待审核的" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Status=3'><img src='images/Photo_passed.gif' border='0' align='absmiddle'><br>已审核的" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Status=-2'><img src='images/Photo_reject.gif' border='0' align='absmiddle'><br>未被采用的" & ChannelShortName & "</a></td>"
    Response.Write "</tr></table>" & vbCrLf

    Select Case Action
    Case "Add"
        Call Add
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SavePhoto
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

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='2' class='border'>"
    Response.Write "<form name='myform' method='Post' action='User_Photo.asp' onsubmit='return ConfirmDel();'><tr>"

    Dim rsPhotoList, sql
    sql = "select P.PhotoID,P.ClassID,C.ClassName,C.ParentDir,C.ClassDir,P.PhotoName,P.Keyword,P.Author,P.UpdateTime,P.Inputer,P.Editor,P.Hits,P.OnTop,P.Elite,P.Status,P.Stars,P.InfoPoint,P.PhotoThumb from PE_Photo P"
    sql = sql & " left join PE_Class C on P.ClassID=C.ClassID where P.ChannelID=" & ChannelID & " and P.Deleted=" & PE_False & " and P.Inputer='" & UserName & "' "
    
    If AuthorName <> "" Then
        sql = sql & " and P.Author='" & AuthorName & "|' "
    End If
    Select Case Status
    Case 3
        sql = sql & " and P.Status=3"
    Case 0
        sql = sql & " and (P.Status=0 Or P.Status=1 Or P.Status=2)"
    Case -1
        sql = sql & " and P.Status=-1"
    Case -2
        sql = sql & " and P.Status=-2"
    End Select
    If ClassID > 0 Then
        If Child > 0 Then
            sql = sql & " and P.ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and P.ClassID=" & ClassID
        End If
    End If

    If Keyword <> "" Then
        Select Case strField
        Case "PhotoName"
            sql = sql & " and P.PhotoName like '%" & Keyword & "%' "
        Case "PhotoIntro"
            sql = sql & " and P.PhotoIntro like '%" & Keyword & "%' "
        Case "Author"
            sql = sql & " and P.Author like '%" & Keyword & "%' "
        Case "Inputer"
            sql = sql & " and P.Inputer='" & Keyword & "' "
        Case Else
            sql = sql & " and P.PhotoName like '%" & Keyword & "%' "
        End Select
    End If
    sql = sql & " order by P.PhotoID desc"

    Set rsPhotoList = Server.CreateObject("ADODB.Recordset")
    rsPhotoList.Open sql, Conn, 1, 1
    If rsPhotoList.BOF And rsPhotoList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>" & GetStrNoItem(ClassID, Status) & "<br><br></td></tr>"
    Else
        totalPut = rsPhotoList.RecordCount
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
                rsPhotoList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim PhotoNum, PhotoPath
        PhotoNum = 0
        Do While Not rsPhotoList.EOF
            Response.Write "<td class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""><table width='100%'  cellpadding='0' cellspacing='0'>"
            Response.Write "<tr><td colspan='2' align='center'><a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rsPhotoList("PhotoID") & "'><img src='" & GetPhotoThumb(rsPhotoList("PhotoThumb")) & "' width='130' height='90' border='0'></a></td></tr>"
            If rsPhotoList("ClassID") <> ClassID Then
                Response.Write "<tr><td align='right'>栏目名称：</td><td><a href='" & FileName & "&ClassID=" & rsPhotoList("ClassID") & "'>[" & rsPhotoList("ClassName") & "]</a></td></tr>"
            End If
            Response.Write "<tr><td align='right'>" & ChannelShortName & "名称：</td><td>"
            Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rsPhotoList("PhotoID") & "' title='" & GetLinkTips(rsPhotoList("PhotoName"), rsPhotoList("Author"), rsPhotoList("UpdateTime"), rsPhotoList("Hits"), rsPhotoList("Keyword"), rsPhotoList("Stars"), rsPhotoList("InfoPoint")) & "'>" & rsPhotoList("PhotoName") & "</a>"
            Response.Write "</td></tr>"
            Response.Write "<tr><td align='right'>添 加 者：</td><td><a href='" & FileName & "&field=Inputer&keyword=" & rsPhotoList("Inputer") & "' title='点击将查看此用户录入的所有" & ChannelShortName & "'>" & rsPhotoList("Inputer") & "</a></td></tr>"
            Response.Write "<tr><td align='right'>点 击 数：</td><td>" & rsPhotoList("Hits") & "</td></tr>"
            Response.Write "<tr><td align='right'>" & ChannelShortName & "属性：</td><td>" & GetInfoProperty(rsPhotoList("OnTop"), rsPhotoList("Hits"), rsPhotoList("Elite")) & "</td></tr>"
            Response.Write "<tr><td align='right'>审核状态：</td><td>" & GetInfoStatus(rsPhotoList("Status")) & "</td></tr>"
            Response.Write "<tr><td align='right'>操作选项：</td><td><input name='PhotoID' type='checkbox' onclick='unselectall()' id='PhotoID' value='" & rsPhotoList("PhotoID") & "'>"
            Response.Write "</td></tr>"
            Response.Write "<tr><td colspan='2' align='center'>"
            If rsPhotoList("Inputer") = UserName And (rsPhotoList("Status") <= 0 Or EnableModifyDelete = 1) Then
                Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Modify&PhotoID=" & rsPhotoList("PhotoID") & "'>修改</a>&nbsp;"
                Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Del&PhotoID=" & rsPhotoList("PhotoID") & "' onclick=""return confirm('确定要删除此" & ChannelShortName & "吗？删除后你还可以从回收站中还原。');"">删除</a>&nbsp;"
            End If
            Response.Write "</td></tr>"
            Response.Write "</table></td>"

            PhotoNum = PhotoNum + 1
            If PhotoNum Mod 4 = 0 Then
                Response.Write "</tr><tr>"
            End If
            If PhotoNum >= MaxPerPage Then Exit Do
            rsPhotoList.MoveNext
        Loop
    End If
    rsPhotoList.Close
    Set rsPhotoList = Nothing
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
    Response.Write "<option value='PhotoName' selected>" & ChannelShortName & "名称</option>"
    Response.Write "<option value='PhotoIntro'>" & ChannelShortName & "简介</option>"
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

Function GetLinkTips(PhotoName, Author, UpdateTime, Hits, Keyword, Stars, InfoPoint)
    Dim strLinkTips
    strLinkTips = ""
    strLinkTips = strLinkTips & "名&nbsp;&nbsp;&nbsp;&nbsp;称：" & PhotoName & vbCrLf
    strLinkTips = strLinkTips & "作&nbsp;&nbsp;&nbsp;&nbsp;者：" & Author & vbCrLf
    strLinkTips = strLinkTips & "更新时间：" & UpdateTime & vbCrLf
    strLinkTips = strLinkTips & "查看次数：" & Hits & vbCrLf
    strLinkTips = strLinkTips & "关 键 字：" & Mid(Keyword, 2, Len(Keyword) - 2) & vbCrLf
    strLinkTips = strLinkTips & "推荐等级："
    If Stars = 0 Then
        strLinkTips = strLinkTips & "无"
    Else
        strLinkTips = strLinkTips & String(Stars, "★")
    End If
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "查看点数：" & InfoPoint
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

Sub ShowJS_Photo()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddUrl(){" & vbCrLf
    Response.Write "  var thisurl='" & ChannelShortName & "地址'+(document.myform.PhotoUrl.length+1)+'|http://'; " & vbCrLf
    Response.Write "  var url=prompt('请输入" & ChannelShortName & "地址名称和链接，中间用“|”隔开：',thisurl);" & vbCrLf
    Response.Write "  if(url!=null&&url!=''){document.myform.PhotoUrl.options[document.myform.PhotoUrl.length]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ModifyUrl(){" & vbCrLf
    Response.Write "  if(document.myform.PhotoUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.PhotoUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('请先选择一个" & ChannelShortName & "地址，再点修改按钮！');return false;}" & vbCrLf
    Response.Write "  var url=prompt('请输入" & ChannelShortName & "地址名称和链接，中间用“|”隔开：',thisurl);" & vbCrLf
    Response.Write "  if(url!=thisurl&&url!=null&&url!=''){document.myform.PhotoUrl.options[document.myform.PhotoUrl.selectedIndex]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function DelUrl(){" & vbCrLf
    Response.Write "  if(document.myform.PhotoUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.PhotoUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('请先选择一个" & ChannelShortName & "地址，再点删除按钮！');return false;}" & vbCrLf
    Response.Write "  document.myform.PhotoUrl.options[document.myform.PhotoUrl.selectedIndex]=null;" & vbCrLf
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
    Response.Write "  if (document.myform.PhotoName.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "名称不能为空！');" & vbCrLf
    Response.Write "    document.myform.PhotoName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    alert('关键字不能为空！');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.PhotoThumb.value==''){" & vbCrLf
    Response.Write "    alert('缩略图地址不能为空！');" & vbCrLf
    Response.Write "    document.myform.PhotoThumb.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.PhotoUrl.length==0){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "" & ChannelShortName & "地址不能为空！');" & vbCrLf
    Response.Write "    document.myform.PhotoUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.PhotoIntro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  document.myform.PhotoUrls.value=''" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.PhotoUrl.length;i++){" & vbCrLf
    Response.Write "    if (document.myform.PhotoUrls.value=='') document.myform.PhotoUrls.value=document.myform.PhotoUrl.options[i].value;" & vbCrLf
    Response.Write "    else document.myform.PhotoUrls.value+='$$$'+document.myform.PhotoUrl.options[i].value;" & vbCrLf
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
        Set trs = Conn.Execute("select count(PhotoID) from PE_Photo where Inputer='" & UserName & "' and Deleted=" & PE_False & " and DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)=0")
        If trs(0) >= MaxPerDay Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您今天发表的" & ChannelShortName & "已经达到了上限！</li>"
        End If
        Set trs = Nothing
        If FoundErr = True Then Exit Sub
    End If
    
    Call ShowJS_Photo
    
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Photo.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>添加" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属栏目：</strong></td>"
    Response.Write "            <td><select name='ClassID'>" & User_GetClass_Option(4, ClassID) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属专题：</strong></td>"
    Response.Write "            <td><select name='SpecialID'><option value='0'>不属于任何专题</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "名称：</strong></td>"
    Response.Write "            <td><input name='PhotoName' type='text' value='' size='50' maxlength='255'> <font color='#FF0000'>*</font></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>关键字：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Keyword' type='text' id='Keyword' value='" & Trim(Session("Keyword")) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("User", ChannelID)
    Response.Write "              <br><font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "作者：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='50' maxlength='30'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "来源：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5' valign='middle'><strong>" & ChannelShortName & "简介：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='PhotoIntro' cols='67' rows='5' id='PhotoIntro' style='display:none'></textarea>"
    Response.Write "               <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=PhotoIntro' frameborder='1' scrolling='no' width='650' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>缩略图：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='PhotoThumb' type='text' id='PhotoThumb' size='80' maxlength='200'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "地址：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='410'>"
    Response.Write "                    <input type='hidden' name='PhotoUrls' value=''>"
    Response.Write "                    <select name='PhotoUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'></select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='addurl' value='添加外部地址' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='修改当前地址' onclick='return ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='删除当前地址' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>上传" & ChannelShortName & "：</strong></td>"
    Response.Write "            <td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=photos' frameborder=0 scrolling=no width='650' height='150'></iframe></td>"
    Response.Write "          </tr>"
    '自定义字段
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-3")
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
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub Modify()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起！您没有在" & ChannelName & "添加" & ChannelShortName & "的权限！</li><br><br>"
        Exit Sub
    End If
    Dim rsPhoto, sql, tmpAuthor, tmpCopyFrom, SpecialID
    
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        PhotoID = PE_CLng(PhotoID)
    End If
    sql = "select * from PE_Photo where Inputer='" & UserName & "' and Deleted=" & PE_False & " and PhotoID=" & PhotoID & ""
    Set rsPhoto = Server.CreateObject("ADODB.Recordset")
    rsPhoto.Open sql, Conn, 1, 1
    If rsPhoto.BOF And rsPhoto.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "</li>"
    Else
        If rsPhoto("Status") > 0 And EnableModifyDelete = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "已经被审核通过或在审核中，您不能再进行修改！</li>"
        End If
    End If
    If FoundErr = True Then
        rsPhoto.Close
        Set rsPhoto = Nothing
        Exit Sub
    End If
    SpecialID = PE_CLng(Conn.Execute("select top 1 SpecialID from PE_InfoS where ModuleType=3 and ItemID=" & PhotoID & "")(0))

    If Right(rsPhoto("Author"), 1) = "|" Then
        tmpAuthor = Left(rsPhoto("Author"), Len(rsPhoto("Author")) - 1)
    Else
        tmpAuthor = rsPhoto("Author")
    End If
    If Right(rsPhoto("CopyFrom"), 1) = "|" Then
        tmpCopyFrom = Left(rsPhoto("CopyFrom"), Len(rsPhoto("CopyFrom")) - 1)
    Else
        tmpCopyFrom = rsPhoto("CopyFrom")
    End If

    Call ShowJS_Photo

    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Photo.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>修改" & ChannelShortName & "</b></td>"
    Response.Write "    </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属栏目：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <select name='ClassID'>" & User_GetClass_Option(4, rsPhoto("ClassID")) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>所属专题：</strong></td>"
    Response.Write "            <td><select name='SpecialID'><option value='0'>不属于任何专题</option>" & GetSpecial_Option(SpecialID) & "</select></td>"
    Response.Write "          </tr>"

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "名称：</strong></td>"
    Response.Write "            <td><input name='PhotoName' type='text' value='" & rsPhoto("PhotoName") & "' size='50' maxlength='255'><font color='#FF0000'>*</font></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>关键字：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Keyword' type='text' id='Keyword' value='" & Mid(rsPhoto("Keyword"), 2, Len(rsPhoto("Keyword")) - 2) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("User", ChannelID)
    Response.Write "              <br><font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "作者：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & tmpAuthor & "' size='50' maxlength='30'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "来源：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5' valign='middle'><strong>" & ChannelShortName & "简介：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='PhotoIntro' cols='67' rows='5' id='PhotoIntro' style='display:none'>" & Server.HTMLEncode(FilterJS(rsPhoto("PhotoIntro"))) & "</textarea>"
    Response.Write "              <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=PhotoIntro' frameborder='1' scrolling='no' width='650' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>缩略图：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='PhotoThumb' type='text' id='PhotoThumb' size='80' maxlength='200' value='" & rsPhoto("PhotoThumb") & "'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "地址：</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='410'>"
    Response.Write "                    <input type='hidden' name='PhotoUrls' value=''>"
    Response.Write "                    <select name='PhotoUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'>"
    Dim PhotoUrls, arrPhotoUrls, iTemp
    PhotoUrls = rsPhoto("PhotoUrl")
    If InStr(PhotoUrls, "$$$") > 1 Then
        arrPhotoUrls = Split(PhotoUrls, "$$$")
        For iTemp = 0 To UBound(arrPhotoUrls)
            Response.Write "<option value='" & arrPhotoUrls(iTemp) & "'>" & arrPhotoUrls(iTemp) & "</option>"
        Next
    Else
        Response.Write "<option value='" & PhotoUrls & "'>" & PhotoUrls & "</option>"
    End If
    Response.Write "                    </select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='addurl' value='添加外部地址' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='修改当前地址' onclick='ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='删除当前地址' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' class='tdbg5'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>上传" & ChannelShortName & "：</strong></td>"
    Response.Write "            <td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=photos' frameborder=0 scrolling=no width='650' height='150'></iframe></td>"
    Response.Write "          </tr>"
    '自定义字段
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-3")
    Do While Not rsField.EOF
        IF rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsPhoto(Trim(rsField("FieldName"))), rsField("Options"), rsField("EnableNull"))
        End If	
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "状态：</td>"
    Response.Write "            <td>"
    If rsPhoto("Status") <= 0 Then
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='-1'"
        If rsPhoto("Status") = -1 Then
            Response.Write " checked"
        End If
        Response.Write "> 草稿&nbsp;&nbsp;"
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='0'"
        If rsPhoto("Status") = 0 Then
            Response.Write "checked"
        End If
        Response.Write "> 投稿"
    Else
        If rsPhoto("Status") < 3 Then
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
    Response.Write "   <input name='PhotoID' type='hidden' id='PhotoID' value='" & rsPhoto("PhotoID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='保存修改结果' style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsPhoto.Close
    Set rsPhoto = Nothing

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
    Case 1, 8    '单行文本框
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2, 9   '多行文本框
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

Sub SavePhoto()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起！您没有在" & ChannelName & "添加" & ChannelShortName & "的权限！</li><br><br>"
        Exit Sub
    End If
    Dim rsPhoto, sql
    Dim trs, tAuthor
    Dim PhotoID, ClassID, SpecialID, PhotoName, Keyword, Author, CopyFrom, PhotoIntro
    Dim PhotoThumb, PhotoUrl, Inputer
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent

    PhotoID = PE_CLng(Trim(Request.Form("PhotoID")))
    ClassID = PE_CLng(Trim(Request.Form("ClassID")))
    SpecialID = PE_CLng(Trim(Request.Form("SpecialID")))
    PhotoName = Trim(Request.Form("PhotoName"))
    Keyword = Trim(Request.Form("Keyword"))
    Author = Trim(Request.Form("Author"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    PhotoIntro = ReplaceBadUrl(FilterJS(Trim(Request.Form("PhotoIntro"))))
    PhotoThumb = PE_HTMLEncode(Trim(Request.Form("PhotoThumb")))
    PhotoUrl = PE_HTMLEncode(Trim(Request.Form("PhotoUrls")))
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
    
    If PhotoName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "名称不能为空</li>"
    End If
	Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入" & ChannelShortName & "关键字</li>"
    Else
        Call SaveKeyword(Keyword)
    End If
    If PhotoThumb = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>缩略图地址不能为空</li>"
    End If
    If PhotoUrl = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "地址不能为空</li>"
    End If
    
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-3")
    Do While Not rsField.EOF
        If rsField("EnableNull") = False Then
            If Trim(Request(rsField("FieldName"))) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请输入" & rsField("Title") & "！</li>"
            End If
        End If
        rsField.MoveNext
    Loop
    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "佚名")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "本站原创")
    
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
    
    PhotoName = PE_HTMLEncode(PhotoName)
    Keyword = "|" & ReplaceBadChar(Keyword) & "|"
    Author = PE_HTMLEncode(Author)
    CopyFrom = PE_HTMLEncode(CopyFrom)
    
    
    Set rsPhoto = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("PhotoName") = PhotoName And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>请不要重复添加同一" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("PhotoName") = PhotoName
            Session("AddTime") = Now()
            If MaxPerDay > 0 Then
                Set trs = Conn.Execute("select count(PhotoID) from PE_Photo where Inputer='" & UserName & "' and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")=0")
                If trs(0) >= MaxPerDay Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>您今天发表的" & ChannelShortName & "已经达到了上限！</li>"
                End If
                Set trs = Nothing
                If FoundErr = True Then Exit Sub
            End If
            
            sql = "select top 1 * from PE_Photo"
            rsPhoto.Open sql, Conn, 1, 3
            rsPhoto.addnew
            PhotoID = PE_CLng(Conn.Execute("select max(PhotoID) from PE_Photo")(0)) + 1
            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (3," & PhotoID & "," & SpecialID & ")")
            rsPhoto("PhotoID") = PhotoID
            rsPhoto("ChannelID") = ChannelID
            rsPhoto("ClassID") = ClassID
            rsPhoto("PhotoName") = PhotoName
            rsPhoto("Keyword") = Keyword
            rsPhoto("Author") = Author
            rsPhoto("CopyFrom") = CopyFrom
            rsPhoto("PhotoIntro") = PhotoIntro
            rsPhoto("PhotoThumb") = PhotoThumb
            rsPhoto("PhotoUrl") = PhotoUrl
            rsPhoto("Hits") = 0
            rsPhoto("DayHits") = 0
            rsPhoto("WeekHits") = 0
            rsPhoto("MonthHits") = 0
            rsPhoto("Stars") = 0
            rsPhoto("UpdateTime") = Now()
            rsPhoto("Status") = Status
            rsPhoto("OnTop") = False
            rsPhoto("Elite") = False
            rsPhoto("Inputer") = Inputer
            rsPhoto("Editor") = Inputer
            rsPhoto("SkinID") = 0
            rsPhoto("TemplateID") = 0
            rsPhoto("Deleted") = False
            PresentExp = CLng(PresentExp * PresentExpTimes)
            rsPhoto("PresentExp") = PresentExp
            rsPhoto("InfoPoint") = DefaultItemPoint
            rsPhoto("VoteID") = 0
            rsPhoto("InfoPurview") = 0
            rsPhoto("arrGroupID") = ""
            rsPhoto("ChargeType") = DefaultItemChargeType
            rsPhoto("PitchTime") = DefaultItemPitchTime
            rsPhoto("ReadTimes") = DefaultItemReadTimes
            rsPhoto("DividePercent") = DefaultItemDividePercent
            
            If Not (rsField.BOF And rsField.EOF) Then
                rsField.MoveFirst
                Do While Not rsField.EOF
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rsPhoto(Trim(rsField("FieldName"))) = Trim(FilterJS(Request(rsField("FieldName"))))
                    End If
                    rsField.MoveNext
                Loop
            End If
            Set rsField = Nothing

            If BlogFlag = True Then '写入BLOGID
                Dim blogid
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsPhoto("BlogID") = 0
                Else
                    rsPhoto("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If

            rsPhoto.Update
            If CheckLevel = 0 Or NeedlessCheck = 1 Then
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1,ItemChecked=ItemChecked+1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_Class set ItemCount=ItemCount+1 where ClassID=" & ClassID & "")
                If rsPhoto("Status") = 3 Then
                    Conn.Execute ("update PE_User set PostItems=PostItems+1,PassedItems=PassedItems+1,UserExp=UserExp+" & PresentExp & " where UserName='" & UserName & "'")
                End If
            Else
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_User set PostItems=PostItems+1 where UserName='" & UserName & "'")
            End If
        End If
    ElseIf Action = "SaveModify" Then
        If PhotoID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定" & ChannelShortName & "ID的值</li>"
        Else
            sql = "select * from PE_Photo where Inputer='" & UserName & "' and Deleted=" & PE_False & " and PhotoID=" & PhotoID
            rsPhoto.Open sql, Conn, 1, 3
            If rsPhoto.BOF And rsPhoto.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到此" & ChannelShortName & "，可能已经被其他人删除。</li>"
            Else
                If rsPhoto("Status") > 0 And EnableModifyDelete = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & ChannelShortName & "已经被审核通过，您不能再进行修改！</li>"
                Else
                    Conn.Execute ("delete from PE_InfoS where ModuleType=3 and ItemID=" & PhotoID)
                    Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (3," & PhotoID & "," & SpecialID & ")")
                    rsPhoto("ClassID") = ClassID
                    rsPhoto("PhotoName") = PhotoName
                    rsPhoto("Keyword") = Keyword
                    rsPhoto("Author") = Author
                    rsPhoto("CopyFrom") = CopyFrom
                    rsPhoto("PhotoIntro") = PhotoIntro
                    rsPhoto("PhotoThumb") = PhotoThumb
                    rsPhoto("PhotoUrl") = PhotoUrl
                    rsPhoto("Status") = Status
                    
                    If Not (rsField.BOF And rsField.EOF) Then
                        rsField.MoveFirst
                        Do While Not rsField.EOF
                            If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                                rsPhoto(Trim(rsField("FieldName"))) = PE_HTMLEncode(FilterJS(Trim(Request(rsField("FieldName")))))
                            End If
                            rsField.MoveNext
                        Loop
                    End If
                    Set rsField = Nothing

                    rsPhoto.Update
                End If
            End If
        End If
    End If
    rsPhoto.Close
    Set rsPhoto = Nothing
    
    If FoundErr = True Then Exit Sub
    
    Response.Write "<br><br>"
    Response.Write "<table class='border' align=center width='500' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
    Response.Write "  <tr align=center>"
    Response.Write "    <td  height='22' colspan='2' align='center' class='title'>"
    If Action = "SaveAdd" Then
        Response.Write "<b>添加" & ChannelShortName & "成功</b>"
    Else
        Response.Write "<b>修改" & ChannelShortName & "成功</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    If Status = 0 Then
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td height='60' colspan='2'><font color='#0000FF'>注意：</font><br>&nbsp;&nbsp;&nbsp;&nbsp;您的" & ChannelShortName & "尚未真正发表！只有等管理员审核并通过了您的" & ChannelShortName & "后，您所添加的" & ChannelShortName & "才会发表。</td>"
        Response.Write "  </tr>"
    End If
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='5' colspan='2'></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='150' align='center' valign='top'><img src='" & GetPhotoThumb(PhotoThumb) & "' width='150'></td>"
    Response.Write "    <td width='350' valign='top'><table width='100%' border='0' cellpadding='2' cellspacing='0'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>所属栏目：</strong></td>"
    Response.Write "          <td>" & ShowClassPath() & "</td>"
    Response.Write "        </tr>"
    Response.Write "       <tr> "
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "名称：</strong></td>"
    Response.Write "          <td>" & PE_HTMLEncode(PhotoName) & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr> "
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "作者：</strong></td>"
    Response.Write "          <td>" & PE_HTMLEncode(Author) & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr> "
    Response.Write "          <td width='100' align='right'><strong>关 键 字：</strong></td>"
    Response.Write "          <td>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='40' colspan='2' align='center'>"
    Response.Write "【<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Modify&PhotoID=" & PhotoID & "'>修改此" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>继续添加" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "管理</a>】&nbsp;"
    Response.Write "【<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & PhotoID & "'>预览" & ChannelShortName & "内容</a>】"
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
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If

    Dim sqlDel, rsDel, NeedUpdateCache
    NeedUpdateCache = False

    sqlDel = "select * from PE_Photo where Inputer='" & UserName & "' and Deleted=" & PE_False & " and "
    If InStr(PhotoID, ",") > 0 Then
        sqlDel = sqlDel & " PhotoID in (" & PhotoID & ") order by PhotoID"
    Else
        sqlDel = sqlDel & " PhotoID=" & PhotoID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If rsDel("Status") > 0 Then
            If EnableModifyDelete = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>删除" & ChannelShortName & "“" & rsDel("PhotoName") & "”失败。原因：此" & ChannelShortName & "已经被审核通过，您不能再删除！</li>"
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
        Call Class_Terminate
    End If
End Sub

Sub Show()
    Dim rs, sql
    PhotoID = PE_CLng(PhotoID)
    sql = "select * from PE_Photo where Inputer='" & UserName & "' and Deleted=" & PE_False & " and PhotoID=" & PhotoID & ""
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


    Response.Write "<br><table width='100%' border=0 align=center cellPadding=2 cellSpacing=1 bgcolor='#FFFFFF' class='border' style='WORD-BREAK: break-all'>"
    Response.Write "<tr class='title'>"
    Response.Write "  <td height='22' colspan='4'>"
    Response.Write "您现在的位置：&nbsp;<a href='User_Photo.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "管理</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rs("PhotoID") & "'>" & PE_HTMLEncode(rs("PhotoName")) & "</a>"
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "名称：</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(rs("PhotoName")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "作者：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("Author")) & "</td>"
    Response.Write "  <td colspan='2' rowspan='8' align=center valign='middle'>"
    If rs("PhotoThumb") = "" Then
        Response.Write "无缩略图"
    Else
        Response.Write "<img src='" & GetPhotoThumb(rs("PhotoThumb")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>添加时间：</td>"
    Response.Write "  <td width='200'>" & rs("UpdateTime") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>评分等级：</td>"
    Response.Write "  <td width='200'>" & String(rs("Stars"), "★") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>查看点数：</td>"
    Response.Write "  <td width='200'><font color=red> " & rs("InfoPoint") & "</font> 点</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>软件添加：</td>"
    Response.Write "  <td width='200'>" & rs("Inputer") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>责任编辑：</td>"
    Response.Write "  <td width='200'>"
    If rs("Status") > 0 Then
        Response.Write rs("Editor")
    Else
        Response.Write "无"
    End If
    Response.Write "  </td>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>查看次数：</td>"
    Response.Write "  <td colspan='3'>本日：" & rs("DayHits") & "&nbsp;&nbsp;&nbsp;&nbsp;本周：" & rs("WeekHits") & "&nbsp;&nbsp;&nbsp;&nbsp;本月：" & rs("MonthHits") & "&nbsp;&nbsp;&nbsp;&nbsp;总计：" & rs("Hits")
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "地址：</td>"
    Response.Write "  <td colspan='3'>" & ShowPhotoUrls(rs("PhotoUrl")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td align='right'>&nbsp;</td>"
    Response.Write "  <td colspan='3' align='right'>"
    Response.Write "<strong>可用操作：</strong>"
    If rs("Inputer") = UserName And (rs("Status") <= 0 Or UserSetting(2) = 1) Then
        Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Modify&PhotoID=" & rs("PhotoID") & "'>修改</a>&nbsp;&nbsp;"
        Response.Write "<a href='User_Photo.asp?ChannelID=" & ChannelID & "&Action=Del&PhotoID=" & rs("PhotoID") & "' onclick=""return confirm('确定要删除此" & ChannelShortName & "吗？');"">删除</a>&nbsp;&nbsp;"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "简介：</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterJS(rs("PhotoIntro")) & "</td>"
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

    Response.Write PE_HTMLEncode(Request("PhotoName"))
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "名称：</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(Request("PhotoName")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "作者：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("Author")) & "</a></td>"
    Response.Write "  <td colspan='2' rowspan='6' align=center valign='middle'>"
    If Request("PhotoThumb") = "" Then
        Response.Write "无缩略图"
    Else
        Response.Write "<img src='" & GetPhotoThumb(Request("PhotoThumb")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "来源：</td>"
    Response.Write "  <td width='200'>" & Request("CopyFrom") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>添加时间：</td>"
    Response.Write "  <td width='200'>" & Now() & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>评分等级：</td>"
    Response.Write "  <td width='200'>" & String(Request("Stars"), "★") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "地址：</td>"
    Response.Write "  <td colspan='3'>" & ShowPhotoUrls(Request("PhotoUrls")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "简介：</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterJS(Request("PhotoIntro")) & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>【<a href='javascript:window.close();'>关闭窗口</a>】</p>"
End Sub

Function GetPhotoThumb(PhotoThumb)
    If Left(PhotoThumb, 1) <> "/" And InStr(PhotoThumb, "://") <= 0 Then
        GetPhotoThumb = strInstallDir & ChannelDir & "/" & UploadDir & "/" & PhotoThumb
    Else
        GetPhotoThumb = PhotoThumb
    End If
End Function

Function ShowPhotoUrls(PhotoUrls)
    Dim arrPhotoUrls, arrUrls, iTemp, strUrls
    strUrls = ""
    arrPhotoUrls = Split(PhotoUrls, "$$$")
    For iTemp = 0 To UBound(arrPhotoUrls)
        arrUrls = Split(arrPhotoUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                strUrls = strUrls & arrUrls(0) & "：<a href='" & strInstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            Else
                strUrls = strUrls & arrUrls(0) & "：<a href='" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            End If
        End If
    Next
    ShowPhotoUrls = strUrls
End Function
%>
