<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
If ShowAnonymous = False Then
    Call WriteErrMsg("网站未开启匿名投稿功能",ComeUrl)
    response.end
End If
Dim IsUpload
'给用户的相应变量赋值
Sub GetUserAmon()
    If ShowAnonymous = False Then
        Call WriteErrMsg("网站未开启匿名投稿功能",ComeUrl)
        Exit Sub    
    End If	
    Dim  rsGroup
    Set rsGroup = Conn.Execute("select * from PE_UserGroup where GroupID=-1")
    GroupName = rsGroup("GroupName")
    GroupType = rsGroup("GroupType")
    arrClass_Browse = Trim(rsGroup("arrClass_Browse"))
    arrClass_View = Trim(rsGroup("arrClass_View"))
    arrClass_Input = Trim(rsGroup("arrClass_Input"))
    UserSetting = Split(Trim(rsGroup("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
    rsGroup.Close
    Set rsGroup = Nothing
    NeedlessCheck = PE_CLng(UserSetting(1))
    EnableModifyDelete = PE_CLng(UserSetting(2))
    MaxPerDay = PE_CLng(UserSetting(3))
    PresentExpTimes = PE_CDbl(UserSetting(4))
    MaxSendNum = PE_CLng(UserSetting(7))
    MaxFavorite = PE_CLng(UserSetting(8))
    Discount_Member = PE_CDbl(UserSetting(11))
    UserEnableComment = PE_CBool(UserSetting(5))
    UserCheckComment = PE_CBool(UserSetting(6))
    If PE_CBool(PE_CLng(UserSetting(9))) = True and  ShowAnonymous = True Then
        IsUpload = True
    Else
 	    IsUpload = False	
    End If	       
    UserChargeType = PE_CLng(UserSetting(14))
End Sub

Sub GetClass()
	ClassName = ""
	RootID = 0
	ParentID = 0
	Depth = 0
	ParentPath = "0"
	Child = 0
	arrChildID = ""
    If ClassID > 0 Then
        Dim tClass
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
        Else
            ClassName = tClass(0)
            RootID = tClass(1)
            ParentID = tClass(2)
            Depth = tClass(3)
            ParentPath = tClass(4)
            Child = tClass(5)
            arrChildID = tClass(6)
        End If
        Set tClass = Nothing
    End If
End Sub
'**************************************************
'函数名：CheckPurview_Class
'作  用：栏目权限数组检测
'参  数：str1 ---- 要比较数组1
'        str2 ---- 要比较数组2
'返回值：True  ---- 存在
'**************************************************
Function CheckPurview_Class(str1, str2)
    Dim arrTemp, arrTemp2, i, j
    CheckPurview_Class = False
    If IsNull(str1) Or IsNull(str2) Or str1 = "" Or str2 = "" Then
        Exit Function
    End If
    arrTemp = Split(str1 & ",", ",")
    arrTemp2 = Split(str2 & ",", ",")
    For i = 0 To UBound(arrTemp)
        For j = 0 To UBound(arrTemp2)
            If Trim(arrTemp2(j)) <> "" And Trim(arrTemp(i)) <> "" And Trim(arrTemp2(j)) = Trim(arrTemp(i)) Then
                CheckPurview_Class = True
                Exit Function
            End If
        Next
    Next
End Function
'**************************************************
'函数名：User_GetClass_Option
'作  用：显示用户栏目下拉菜单
'参  数：ShowType ----显示类型
'        CurrentID ----当前栏目ID
'返回值：用户栏目下拉菜单
'**************************************************
Function User_GetClass_Option(ShowType, CurrentID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i, ClassNum
    Dim arrShowLine(20)
    Dim CheckParentPath, PurviewChecked
    
    ClassNum = 1
    CurrentID = PE_CLng(CurrentID)
    
    sqlClass = "Select * from PE_Class where ChannelID=" & ChannelID & " And ClassType=1 order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strClass_Option = strClass_Option & "<option value=''>请先添加栏目</option>"
    Else
        Do While Not rsClass.EOF
            ClassNum = ClassNum + 1
            tmpDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            If ShowType = 1 Then
                strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
            Else
                If rsClass("ParentID") > 0 Then
                    CheckParentPath = ChannelDir & "all," & rsClass("ParentPath") & "," & rsClass("ClassID") & "," & rsClass("arrChildID")
                Else
                    CheckParentPath = ChannelDir & "all," & rsClass("ClassID") & "," & rsClass("arrChildID")
                End If

                If CheckPurview_Class(arrClass_Input, CheckParentPath) = True Then
                    PurviewChecked = True
                    If rsClass("Child") > 0 And rsClass("EnableAdd") = False And rsClass("ClassID") <> CurrentID Then
                        strClass_Option = strClass_Option & "<option value='0'"
                    Else
                        strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                    End If
                Else
                    PurviewChecked = False
                End If
            End If
            If ShowType = 1 Or PurviewChecked = True Then
                If CurrentID = 0 Then
                    If ClassNum = 1 Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                Else
                    If rsClass("ClassID") = CurrentID Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                End If
                strClass_Option = strClass_Option & ">"
                
                If tmpDepth > 0 Then
                    For i = 1 To tmpDepth
                        strClass_Option = strClass_Option & "&nbsp;&nbsp;"
                    Next
                End If
                strClass_Option = strClass_Option & rsClass("ClassName")
                strClass_Option = strClass_Option & "</option>"
            
                ClassNum = ClassNum + 1
            End If
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    User_GetClass_Option = strClass_Option	
End Function


'**************************************************
'函数名：User_GetChannel_Option
'作  用：显示用户频道下拉菜单
'参  数：ShowType ----显示类型
'        CurrentID ----当前栏目ID
'返回值：用户栏目下拉菜单
'**************************************************

Function User_GetChannel_Option()
    arrClass_Input = Conn.Execute("SELECT arrClass_Input from PE_UserGroup where GroupID=-1")(0)
    Dim strChannel_Option,rsChannel
    Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType=1 AND Disabled=" & PE_False & " ORDER BY OrderID")	
    If PE_CLng(ChannelID) = 0 Then
        strChannel_Option = strChannel_Option & "<option value='0' selected>请选择频道</option>"
    End If
    Do While not rsChannel.Eof 
	    If FoundInArr(arrClass_Input, rsChannel("ChannelDir") & "none", ",") = True Then
        Else
            strChannel_Option = strChannel_Option & "<option value='" & rsChannel("ChannelID") &"'"
            If rsChannel("ChannelID") = ChannelID then strChannel_Option = strChannel_Option & " selected"			   
            strChannel_Option = strChannel_Option &  ">" &rsChannel(1)&"</option>"		
		End If	
	rsChannel.MoveNext
	Loop
    User_GetChannel_Option = strChannel_Option
End Function 
%>
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
Dim ArticleID, AuthorName, Status, ManageType
Dim IncludePic, UploadFiles, DefaultPicUrl
Dim ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created
Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview
Dim ChannelUpload
Sub Execute()
    ChannelID = PE_CLng(Request("ChannelID"))
    Call GetUserAmon
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
        ChannelUpload = Conn.Execute("Select EnableUploadFile From PE_Channel Where ChannelID = "& ChannelID)(0)
        If ChannelUpload = False Then IsUpload = False		
    'Else
    '   FoundErr = True
    '   ErrMsg = ErrMsg & "<li>请指定要查看的频道ID！</li>"
    '   Response.Write ErrMsg
    '   Exit Sub
    Else
        ChannelShortName = "文章"	
        IsUpload = False			
    End If

    ArticleID = Trim(Request("ArticleID"))
    ClassID = PE_CLng(Trim(Request("ClassID")))
    Status = Trim(Request("Status"))
    AuthorName = Trim(Request("AuthorName"))
    If Status = "" Then
        Status = 9
    Else
        Status = PE_CLng(Status)
    End If
    If IsValidID(ArticleID) = False Then
        ArticleID = ""
    End If
    ManageType = Trim(Request("ManageType"))

    If Action = "" Then Action = "Manage"
    FileName = "User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
    strFileName = FileName & "&ClassID=" & ClassID & "&Field=" & strField & "&keyword=" & Keyword
    If AuthorName <> "" Then
        AuthorName = ReplaceBadChar(AuthorName)
        strFileName = strFileName & "&AuthorName=" & AuthorName
    End If


    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
    ArticlePro1 = XmlText("Article", "ArticlePro1", "[图文]")
    ArticlePro2 = XmlText("Article", "ArticlePro2", "[组图]")
    ArticlePro3 = XmlText("Article", "ArticlePro3", "[推荐]")
    ArticlePro4 = XmlText("Article", "ArticlePro4", "[注意]")


    Select Case Action
    Case "Add"
        Call Add
    Case "SaveAdd"
        Call SaveArticle
    Case "Preview"
        Call Preview
    Case "Del"
        Call Del
    Case "Show"
        Call Show	
    Case Else
        Call Add
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
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

Function GetInfoIncludePic(IncludePic)
    Dim strInfoIncludePic
    Select Case PE_CLng(IncludePic)
        Case 1
            strInfoIncludePic = "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            strInfoIncludePic = "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            strInfoIncludePic = "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            strInfoIncludePic = "<font color=blue>" & ArticlePro4 & "</font>"
    End Select
    GetInfoIncludePic = strInfoIncludePic
End Function

Function GetLinkTips(Title, Author, CopyFrom, UpdateTime, Hits, Keyword, Stars, PaginationType, InfoPoint)
    Dim strLinkTips
    strLinkTips = ""
    strLinkTips = strLinkTips & "标&nbsp;&nbsp;&nbsp;&nbsp;题：" & Title & vbCrLf
    strLinkTips = strLinkTips & "作&nbsp;&nbsp;&nbsp;&nbsp;者：" & Author & vbCrLf
    strLinkTips = strLinkTips & "转 贴 自：" & CopyFrom & vbCrLf
    strLinkTips = strLinkTips & "更新时间：" & UpdateTime & vbCrLf
    strLinkTips = strLinkTips & "点 击 数：" & Hits & vbCrLf
    strLinkTips = strLinkTips & "关 键 字：" & Mid(Keyword, 2, Len(Keyword) - 2) & vbCrLf
    strLinkTips = strLinkTips & "推荐等级："
    If Stars = 0 Then
        strLinkTips = strLinkTips & "无"
    Else
        strLinkTips = strLinkTips & String(Stars, "★")
    End If
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "分页方式："
    Select Case PaginationType
    Case 0
        strLinkTips = strLinkTips & "不分页"
    Case 1
        strLinkTips = strLinkTips & "自动分页"
    Case 2
        strLinkTips = strLinkTips & "手动分页"
    End Select
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "阅读点数：" & InfoPoint
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

Function GetInfoProperty(OnTop, Hits, Elite, DefaultPicUrl)
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
    strInfoProperty = strInfoProperty & "&nbsp;"
    If DefaultPicUrl <> "" Then
        strInfoProperty = strInfoProperty & "<font color=blue>图</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    GetInfoProperty = strInfoProperty
End Function

Sub ShowJS_Article()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddItem(strFileName){" & vbCrLf
    Response.Write "  var arrName=strFileName.split('.');" & vbCrLf
    Response.Write "  var FileExt=arrName[1];" & vbCrLf
    Response.Write "  if (FileExt=='gif'||FileExt=='jpg'||FileExt=='jpeg'||FileExt=='jpe'||FileExt=='bmp'||FileExt=='png'){" & vbCrLf
    
    If PE_CLng(UserSetting(21)) = 1 Then
        Response.Write "      if(document.myform.IncludePic.selectedIndex<2){" & vbCrLf
        Response.Write "        document.myform.IncludePic.selectedIndex+=1;" & vbCrLf
        Response.Write "      }" & vbCrLf
    End If

    Response.Write "  document.myform.DefaultPicUrl.value=strFileName;}" & vbCrLf
    Response.Write "  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);" & vbCrLf
    Response.Write "  document.myform.DefaultPicList.selectedIndex+=1;" & vbCrLf
    Response.Write "  if(document.myform.UploadFiles.value==''){" & vbCrLf
    Response.Write "    document.myform.UploadFiles.value=strFileName;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    document.myform.UploadFiles.value=document.myform.UploadFiles.value+'|'+strFileName;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function selectPaginationType(){" & vbCrLf
    Response.Write "  document.myform.PaginationType.value=2;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function rUseLinkUrl(){" & vbCrLf
    Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=false;" & vbCrLf
    Response.Write "     ArticleContent.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=true;" & vbCrLf
    Response.Write "    ArticleContent.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "   document.myform.Content.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    alert('预览状态不能保存！请先回到编辑状态后再保存');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.curChannelID.value=='0'){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "请指定频道！');" & vbCrLf
    Response.Write "    document.myform.curChannelID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf	
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
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "标题不能为空！');" & vbCrLf
    Response.Write "    document.myform.Title.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    alert('关键字不能为空！');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    If PE_CLng(UserSetting(23)) = 1 Then
        Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
        Response.Write "    if (document.myform.LinkUrl.value==''||document.myform.LinkUrl.value=='http://'){" & vbCrLf
        Response.Write "      alert('请输入转向链接的地址！');" & vbCrLf
        Response.Write "      document.myform.LinkUrl.focus();" & vbCrLf
        Response.Write "      return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  else{" & vbCrLf
        Response.Write "    if (document.myform.Content.value==''){" & vbCrLf
        Response.Write "      alert('" & ChannelShortName & "内容不能为空！');" & vbCrLf
        Response.Write "      editor.HtmlEdit.focus();" & vbCrLf
        Response.Write "      return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
    Else
        Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
        Response.Write "    alert('" & ChannelShortName & "内容不能为空！');" & vbCrLf
        Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
        Response.Write "    return false;" & vbCrLf
        Response.Write "  }" & vbCrLf
    End If
    Response.Write "  return true;  " & vbCrLf
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
        Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Inputer='" & UserName & "' and Deleted=" & PE_False & " and DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)=0")
        If trs(0) >= MaxPerDay Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您今天发表的" & ChannelShortName & "已经达到了上限！</li>"
        End If
        Set trs = Nothing
        If FoundErr = True Then Exit Sub
    End If
    
    Call ShowJS_Article
    
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Anonymous.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>添加" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>所属频道：</strong></td>"
    Response.Write "          <td><select  onchange=""window.location.href='User_Anonymous.asp?ChannelID='+this.options[this.selectedIndex].value"" name='curChannelID'>" & User_GetChannel_Option() & "</select></td>"
    Response.Write "        </tr>"	
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>所属栏目：</strong></td>"
    Response.Write "          <td><select name='ClassID'>" & User_GetClass_Option(4, ClassID) & "</select></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>所属专题：</strong></td>"
    Response.Write "          <td><select name='SpecialID'><option value='0'>不属于任何专题</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "标题：</strong></td>"
    Response.Write "          <td>"
    If PE_CLng(UserSetting(21)) = 1 Then
        Response.Write "<select name='IncludePic'><option  value='0' selected> </option><option value='1'>" & ArticlePro1 & "</option><option value='2'>" & ArticlePro2 & "</option><option value='3'>" & ArticlePro3 & "</option><option value='4'>" & ArticlePro4 & "</option></select>"
    Else
       Response.Write "<Input TYPE='hidden' Name='IncludePic' value=''>"
    End If
    Response.Write "          <input name='Title' type='text' id='Title' value='' size='45' maxlength='255' class='bginput'> <font color='#FF0000'>*</font>"
    If PE_CLng(UserSetting(22)) = 1 Then
        Response.Write "<input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'> 显示" & ChannelShortName & "列表时在标题旁显示评论链接"
    End If
    Response.Write "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>关键字：</strong></td>"
    Response.Write "          <td><input name='Keyword' type='text' id='Keyword' value='" & Trim(Session("Keyword")) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font>" & GetKeywordList("User", ChannelID)
    Response.Write "<br><font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "作者：</strong></td>"
    Response.Write "          <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='50' maxlength='100'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "来源：</strong></td>"
    Response.Write "          <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
    Response.Write "          </td>"
    Response.Write "        </tr>"
    If PE_CLng(UserSetting(23)) = 1 Then
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='120' align='right' class='tdbg5'><strong><font color='#FF0000'>转向链接：</font></strong></td>"
        Response.Write "          <td>"
        Response.Write "            <input name='LinkUrl' type='text' id='LinkUrl' value='http://' size='50' maxlength='255' disabled>"
        Response.Write "            <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'>"
        Response.Write "            <font color='#FF0000'>使用转向链接</font></td>"
        Response.Write "        </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "简介：</strong></td>"
    Response.Write "            <td ><textarea name='Intro' cols='80' rows='4'></textarea></td>"
    Response.Write "          </tr>"
    Response.Write "        <tr class='tdbg' id='ArticleContent' style=""display:''"">"
    Response.Write "          <td width='120' align='right' class='tdbg5'><p><strong>" & ChannelShortName & "内容：</strong></p>"
    Response.Write "<br><br><font color='red'>换行请按Shift+Enter<br><br>另起一段请按Enter</font></div>"
    Response.Write "         </td>"
    Response.Write "         <td><textarea name='Content' style='display:none'>" & XmlText("Article", "DefaultAddTemplate", "") & "</textarea>"
    
    If PE_CLng(UserSetting(24)) = 1 Then
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content&Anonymous=1' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    Else
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=2&tContentid=Content&Anonymous=1' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    End If
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong><font color='#FF0000'>首页图片：</font></strong></td>"
    Response.Write "          <td><input name='DefaultPicUrl' type='text' id='DefaultPicUrl' size='56' maxlength='200'>"
    Response.Write "      用于在首页的图片" & ChannelShortName & "处显示 <br>直接从上传图片中选择："
    Response.Write "      <select name='DefaultPicList' id='select' onChange='DefaultPicUrl.value=this.value;'>"
    Response.Write "        <option selected>不指定首页图片</option>"
    Response.Write "      </select><input name='UploadFiles' type='hidden' id='UploadFiles'>"
    Response.Write "          </td>"
    Response.Write "          </tr>"
    '自定义字段
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-1")
    Do While Not rsField.EOF
        IF rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsField("DefaultValue"), rsField("Options"), rsField("EnableNull"))
        End If
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "状态：</strong></td>"
    Response.Write "            <td><input name='Status' type='radio' id='Status' value='-1'>草稿&nbsp;&nbsp;<input Name='Status' Type='Radio' Id='Status' Value='0' checked>投稿</td>"
    Response.Write "          </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'><input name='PaginationType' type='hidden' value='0'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 添 加 ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' 预 览 ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub SaveArticle()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起！您没有在" & ChannelName & "添加" & ChannelShortName & "的权限！</li><br><br>"
        Exit Sub
    End If
    Dim rsArticle, sql, i
    Dim trs
    Dim ArticleID, ClassID, SpecialID, Title, ShowCommentLink, Keyword, UseLinkUrl, LinkUrl, Content, tAuthor, Intro
    Dim Author, CopyFrom, Inputer
    Dim arrUploadFiles, SaveRemotePic
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent

    ArticleID = PE_CLng(Trim(Request.Form("ArticleID")))
    ClassID = PE_CLng(Trim(Request.Form("ClassID")))
    SpecialID = PE_CLng(Trim(Request.Form("SpecialID")))
    Title = PE_HTMLEncode(Trim(Request.Form("Title")))
    ShowCommentLink = Trim(Request.Form("ShowCommentLink"))
    Keyword = Trim(Request.Form("Keyword"))
    UseLinkUrl = PE_HTMLEncode(Trim(Request.Form("UseLinkUrl")))
    LinkUrl = PE_HTMLEncode(Trim(Request.Form("LinkUrl")))
    Intro = PE_HTMLEncode(Trim(Request.Form("Intro")))
    For i = 1 To Request.Form("Content").Count
        Content = Content & FilterJS(Request.Form("Content")(i))
    Next
    Author = PE_HTMLEncode(Trim(Request.Form("Author")))
    CopyFrom = PE_HTMLEncode(Trim(Request.Form("CopyFrom")))
    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
    DefaultPicUrl = PE_HTMLEncode(Trim(Request.Form("DefaultPicUrl")))
    UploadFiles = PE_HTMLEncode(Trim(Request.Form("UploadFiles")))
    SaveRemotePic = PE_HTMLEncode(Trim(Request.Form("SaveRemotePic")))
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

    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "标题不能为空</li>"
    Else
        Title = ReplaceText(Title, 2)
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
    If UseLinkUrl = "Yes" Then
        If LinkUrl = "" Or LCase(LinkUrl) = "http://" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址不能为空</li>"
        ElseIf Left(LCase(LinkUrl), 7) <> "http://" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址必须以 http:// 开头</li>"
        End If
    Else
        If Content = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "内容不能为空</li>"
        End If
    End If

    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-1")
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

    Keyword = "|" & ReplaceBadChar(Keyword) & "|"

    '将绝对地址转化为相对地址
    Dim strSiteUrl
    strSiteUrl = Request.ServerVariables("HTTP_REFERER")
    strSiteUrl = LCase(Left(strSiteUrl, InStrRev(strSiteUrl, "/") - 1))
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/")) & ChannelDir & "/"
    Content = ReplaceBadUrl(Replace(Content, strSiteUrl, "[InstallDir_ChannelDir]"))
    strSiteUrl = InstallDir & ChannelDir & "/"
    Content = Replace(Content, strSiteUrl, "[InstallDir_ChannelDir]")

    Set rsArticle = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("Title") = Title And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>请不要重复添加同一篇文章</li>"
            Exit Sub
        Else
            Session("Title") = Title
            Session("AddTime") = Now()
            If MaxPerDay > 0 Then
                Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Inputer='" & UserName & "' and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")=0")
                If trs(0) >= MaxPerDay Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>您今天发表的" & ChannelShortName & "已经达到了上限！</li>"
                End If
                Set trs = Nothing
                If FoundErr = True Then Exit Sub
            End If
            
            sql = "select top 1 * from PE_Article"
            rsArticle.Open sql, Conn, 1, 3
            rsArticle.addnew
            ArticleID = PE_CLng(Conn.Execute("select max(ArticleID) from PE_Article")(0)) + 1
            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (1," & ArticleID & "," & SpecialID & ")")
            rsArticle("ArticleID") = ArticleID
            rsArticle("ChannelID") = ChannelID
            rsArticle("ClassID") = ClassID
            rsArticle("Title") = Title
            rsArticle("Intro") = Intro
            rsArticle("Content") = Content
            rsArticle("Keyword") = Keyword
            rsArticle("Hits") = 0
            rsArticle("Author") = Author
            rsArticle("CopyFrom") = CopyFrom
            rsArticle("LinkUrl") = LinkUrl
            rsArticle("Inputer") = "匿名投稿"
            rsArticle("Editor") = "匿名投稿"
            rsArticle("IncludePic") = IncludePic
            If ShowCommentLink = "Yes" Then
                rsArticle("ShowCommentLink") = True
            Else
                rsArticle("ShowCommentLink") = False
            End If
            rsArticle("Status") = Status
            rsArticle("OnTop") = False
            'rsArticle("Hot") = False
            rsArticle("Elite") = False
            rsArticle("Stars") = 0
            rsArticle("UpdateTime") = Now()
            rsArticle("PaginationType") = 0
            rsArticle("MaxCharPerPage") = 0
            rsArticle("SkinID") = 0
            rsArticle("TemplateID") = 0
            rsArticle("DefaultPicUrl") = DefaultPicUrl
            rsArticle("UploadFiles") = UploadFiles
            rsArticle("Deleted") = False
            PresentExp = CLng(PresentExp * PresentExpTimes)
            rsArticle("PresentExp") = PresentExp
            rsArticle("InfoPoint") = DefaultItemPoint
            rsArticle("VoteID") = 0
            rsArticle("InfoPurview") = 0
            rsArticle("arrGroupID") = ""
            rsArticle("ChargeType") = DefaultItemChargeType
            rsArticle("PitchTime") = DefaultItemPitchTime
            rsArticle("ReadTimes") = DefaultItemReadTimes
            rsArticle("DividePercent") = DefaultItemDividePercent
            rsArticle("Copymoney") = 0
            rsArticle("IsPayed") = False
            rsArticle("Beneficiary") = UserName
            
            If Not (rsField.BOF And rsField.EOF) Then
                rsField.MoveFirst
                Do While Not rsField.EOF
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rsArticle(Trim(rsField("FieldName"))) = PE_HTMLEncode(FilterJS(Trim(Request(rsField("FieldName")))))
                    End If
                    rsField.MoveNext
                Loop
            End If
            Set rsField = Nothing         
            rsArticle.Update
        End If
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    
    If FoundErr = True Then Exit Sub

    Call UpdateChannelData(ChannelID)
  '  Call UpdateUserData(0, UserName, 0, 0)
    
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
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "标题：</strong></td>"
    Response.Write "          <td>" & Title & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>作&nbsp;&nbsp;&nbsp;&nbsp;者：</strong></td>"
    Response.Write "          <td>" & Author & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>转 贴 自：</strong></td>"
    Response.Write "          <td>" & CopyFrom & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'> "
    Response.Write "          <td width='100' align='right'><strong>关 键 字：</strong></td>"
    Response.Write "          <td>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='30' align='center'>"
    Response.Write "【<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & ArticleID & "'>修改本文</a>】&nbsp;"
    Response.Write "【<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>继续添加" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "'>预览" & ChannelShortName & "内容</a>】"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Call ClearSiteCache(0)
   ' Call CreateAllJS_User
End Sub


Sub Show()
    Dim rsArticle, sql, i
    ArticleID = PE_CLng(ArticleID)
    sql = "select * from PE_Article where Deleted=" & PE_False & " and ArticleID=" & ArticleID & ""
    Set rsArticle = Server.CreateObject("ADODB.Recordset")
    rsArticle.Open sql, Conn, 1, 1
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "</li>"
    Else
        If rsArticle("Inputer") <> UserName And FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = False And rsArticle("Inputer")<>"匿名投稿" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>查看" & ChannelShortName & "失败，此" & ChannelShortName & "是其他人添加的。</li>"
        End If
        ClassID = rsArticle("ClassID")
        Call GetClass
    End If
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function resizepic(thispic){" & vbCrLf
    Response.Write "  if(thispic.width>600) thispic.width=600;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function bbimg(o){" & vbCrLf
    Response.Write "  var zoom=parseInt(o.style.zoom, 10)||100;" & vbCrLf
    Response.Write "  zoom+=event.wheelDelta/12;" & vbCrLf
    Response.Write "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    Response.Write "  return false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf


    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' width='82%'>"
    Response.Write "您现在的位置：&nbsp;<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "管理</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write "<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsArticle("ArticleID") & "'>"
    
    Select Case rsArticle("IncludePic")
        Case 1
            Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
    End Select

    Response.Write "" & rsArticle("Title") & "</a>"
    Response.Write " </td>"
    Response.Write "    <td width='18%' height='22' align='right'>"

    If rsArticle("OnTop") = True Then
        Response.Write "<font color=blue>顶</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Hits") >= HitsOfHot Then
        Response.Write "<font color=red>热</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Elite") = True Then
        Response.Write "<font color=green>荐</font>"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "&nbsp;&nbsp;<font color='#009900'>" & String(rsArticle("Stars"), "★") & "</font>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td colspan='2' height='40' valign='bottom'>"
    If Trim(rsArticle("TitleIntact")) <> "" Then
        Response.Write "<font size='4'><b>" & rsArticle("TitleIntact") & "</b></font>"
    Else
        Response.Write "<font size='4'><b>" & rsArticle("Title") & "</b></font>"
    End If
    If Trim(rsArticle("Subheading")) <> "" Then
        Response.Write "<br>" & rsArticle("Subheading")
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Dim Author, CopyFrom
    Author = rsArticle("Author")
    CopyFrom = rsArticle("CopyFrom")
    Response.Write "作者：" & Author & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "来源："
    If InStr(CopyFrom, "|") > 0 Then
        Response.Write "<a href='" & Right(CopyFrom, Len(CopyFrom) - InStr(CopyFrom, "|")) & "' target='_blank'>" & Left(CopyFrom, InStr(CopyFrom, "|") - 1) & "</a>"
    Else
        Response.Write "" & CopyFrom
    End If
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;点击数：" & rsArticle("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;更新时间：" & FormatDateTime(rsArticle("UpdateTime"), 2)
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan='2'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='5'>"
    Response.Write "        <tr>"
    Response.Write "          <td height='200' valign='top'>"
    If Trim(rsArticle("LinkUrl")) <> "" Then
        Response.Write "<p align='center'><br><br><br><font color=red>本" & ChannelShortName & "是链接外部" & ChannelShortName & "内容。链接地址为：<a href='" & rsArticle("LinkUrl") & "' target='_blank'>" & rsArticle("LinkUrl") & "</a></font></p>"
    Else
        Response.Write "<p>" & Replace(Replace(FilterJS(rsArticle("Content")), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir) & "</p>"
    End If
    Response.Write "       </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr  align='right' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Response.Write "" & ChannelShortName & "录入：<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Field=Inputer&Keyword=" & rsArticle("Inputer") & "'>" & rsArticle("Inputer") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;责任编辑："
    If rsArticle("Status") > 0 Then
        Response.Write "" & rsArticle("Editor")
    Else
        Response.Write "无"
    End If
    Response.Write " </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf
    Response.Write "<form name='formA' method='get' action='User_Anonymous.asp'><p align='center'> "
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='ArticleID' value='" & ArticleID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf
    rsArticle.Close
    Set rsArticle = Nothing
    Response.Write "</Form></p>"
End Sub


Sub Preview()
    Response.Write "<br><table width='760' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td width='400' height='22'>"

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
            Response.Write "" & rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "" & ClassName & "&nbsp;&gt;&gt;&nbsp;" & GetInfoIncludePic(Trim(Request("IncludePic"))) & PE_HTMLEncode(Request("Title"))
    Response.Write " </td>"
    Response.Write "    <td width='50' height='22' align='right'>"
    If LCase(Request("OnTop")) = "yes" Then
        Response.Write "顶&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
  
    If LCase(Request("Elite")) = "yes" Then
        Response.Write "荐"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'><font size='4'>"
    If Trim(Request("TitleIntact")) <> "" Then
        Response.Write "<font size='4'><b>" & PE_HTMLEncode(Request("TitleIntact")) & "</b></font>"
    Else
        Response.Write "<font size='4'><b>" & PE_HTMLEncode(Request("Title")) & "</b></font>"
    End If
    If Trim(Request("Subheading")) <> "" Then
        Response.Write "<br>" & PE_HTMLEncode(Request("Subheading"))
    End If

    Response.Write "</font></td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'>作者：" & PE_HTMLEncode(Request("Author")) & "&nbsp;&nbsp;&nbsp;&nbsp;转贴自：" & PE_HTMLEncode(Request("CopyFrom")) & "&nbsp;&nbsp;&nbsp;&nbsp;点击数：0&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "录入：" & UserName & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3'><p>" & FilterJS(Request("Content")) & "</p></td></tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>【<a href='javascript:window.close();'>关闭窗口</a>】</p>"
End Sub


Sub WriteFieldHTML(FieldName, Title, Tips, FieldType, strValue, Options, EnableNull)
    Dim strEnableNull
    If EnableNull = False Then
        strEnableNull = " <font color='#FF0000'>*</font>"
    End If
    Response.Write "<tr class='tdbg'><td width='120' align='right' class='tdbg5'>" & Title & "：</b><td colspan='5'>"
    Select Case FieldType
    Case 1,8    '单行文本框
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2,9    '多行文本框
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
           ' Response.Write "<input type='text' name='" & FieldName & "' size='40' maxlength='255' value='" & strValue & "'>" & strEnableNull
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
		End If
        If IsUpload = True Then		
            Response.Write "<iframe style='top:2px;' id='uploadPhoto' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldpic&Anonymous=1&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"		
        End If		
    Case 5   '文件
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
		End If
        If IsUpload = True Then				
            Response.Write "            <iframe style='top:2px' id='uploadsoft' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldsoft&Anonymous=1&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"	
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
            Response.Write "<input type='text' name='" & FieldName & "'  onkeyup=""value=value.replace(/[^\d]/g,'')""  size='20' maxlength='20' value='" & PE_Clng(strValue) & "'>" & strEnableNull
        End If		
    End Select
    If IsNull(Tips) = False And Tips <> ""  and (FieldType <> 4 and FieldType <> 5) Then
        Response.Write "<br>" & PE_HTMLEncode(Tips)
    End If
    Response.Write "</td></tr>"
End Sub

%>
