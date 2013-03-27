<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'**************************************************
'函数名：GetNumber_Option
'作  用：显示数字下拉菜单
'参  数：MinNum ---- 初始数
'        MaxNum ---- 最大数
'        CurrentNum ----selected 默认数
'返回值：下拉菜单数据
'**************************************************
Public Function GetNumber_Option(MinNum, MaxNum, CurrentNum)
    Dim strNumber, i
    For i = MinNum To MaxNum
        If i = CurrentNum Then
            strNumber = strNumber & "<option value='" & i & "' selected>&nbsp;&nbsp;" & i & "&nbsp;&nbsp;</option>"
        Else
            strNumber = strNumber & "<option value='" & i & "'>&nbsp;&nbsp;" & i & "&nbsp;&nbsp;</option>"
        End If
    Next
    GetNumber_Option = strNumber
End Function

'**************************************************
'函数名：IsStyleDisplay
'作  用：是否显示层
'参  数：Compare1-----比较值1
'参  数：Compare2-----比较值2
'返回值：替换后字符串
'**************************************************
Public Function IsStyleDisplay(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsStyleDisplay = " style='display:'"
    Else
        IsStyleDisplay = " style='display:none'"
    End If
End Function
'**************************************************
'函数名：RadioValue
'作  用：显示单选框或者多选框的值并判断是否选中
'参  数：compvalue ---- 选项的目前实际值
'        showvalue ---- 选项的显示值
'**************************************************
Public Function RadioValue(compvalue, showvalue)
    If compvalue = showvalue Then
        RadioValue = "value='" & showvalue & "' checked"
    Else
        RadioValue = "value='" & showvalue & "'"
    End If
End Function

'**************************************************
'函数名：OptionValue
'作  用：显示下拉列表的值并判断是否选中
'参  数：compvalue ---- 选项的目前实际值
'        showvalue ---- 选项的显示值
'**************************************************
Public Function OptionValue(compvalue, showvalue)
    If compvalue = showvalue Then
        OptionValue = "value='" & showvalue & "' selected"
    Else
        OptionValue = "value='" & showvalue & "'"
    End If
End Function

'**************************************************
'函数名：GetPayOnlineProviderName
'作  用：显示在线支付的名称
'参  数：PayOnlineProviderID ---- 系统在线支付的ID
'返回值：在线支付的名称
'**************************************************
Public Function GetPayOnlineProviderName(PayOnlineProviderID)
    Select Case PayOnlineProviderID
    Case 1, 10
        GetPayOnlineProviderName = "网银在线"
    Case 2
        GetPayOnlineProviderName = "中国在线支付网"
    Case 3
        GetPayOnlineProviderName = "上海环迅"
    Case 4
        GetPayOnlineProviderName = "广东银联"
    Case 5
        GetPayOnlineProviderName = "西部支付"
    Case 6
        GetPayOnlineProviderName = "易付通"
    Case 7
        GetPayOnlineProviderName = "云网在线"
    Case 8, 12
        GetPayOnlineProviderName = "支付宝"
    Case 9
        GetPayOnlineProviderName = "快钱"
    Case 11
        GetPayOnlineProviderName = "快钱神州行"
    Case 13
        GetPayOnlineProviderName = "财付通"
    End Select
End Function

'**************************************************
'函数名：GetArrItem
'作  用：得到数组中某个元素的值
'参  数：arrTemp ---- 要取的数组
'        ItemIndex ---- 第几位数
'返回值：所属位数的元素
'**************************************************
Public Function GetArrItem(ByVal arrTemp, ByVal ItemIndex)
    If Not IsArray(arrTemp) Then
        GetArrItem = ""
        Exit Function
    End If
    ItemIndex = PE_CLng(ItemIndex)
    If ItemIndex < 0 Or ItemIndex > UBound(arrTemp) Then
        GetArrItem = ""
        Exit Function
    End If
    Dim strTemp
    strTemp = arrTemp(ItemIndex)
    If InStr(strTemp, "|") > 0 Then
        GetArrItem = Left(strTemp, InStr(strTemp, "|") - 1)
    Else
        GetArrItem = strTemp
    End If
End Function

'**************************************************
'函数名：Array2Option
'作  用：把数组变成下拉列表项目
'参  数：arrTemp ---- 数组
'        ItemIndex ---- 数组中默认的数字
'返回值：下拉菜单
'**************************************************
Public Function Array2Option(ByVal arrTemp, ByVal ID)
    Dim strOption, i, arrValue
    strOption = "<option value='-1'> </option>"
    ID = PE_CLng(ID)
    For i = 0 To UBound(arrTemp)
        arrValue = Split(arrTemp(i), "|")
        If CLng(arrValue(1)) = 1 Then
            If ID > -1 Then
                If i = ID Then
                    strOption = strOption & "<option value='" & i & "' selected>" & arrValue(0) & "</option>"
                Else
                    strOption = strOption & "<option value='" & i & "'>" & arrValue(0) & "</option>"
                End If
            Else
                If CLng(arrValue(2)) = 1 Then
                    strOption = strOption & "<option value='" & i & "' selected>" & arrValue(0) & "</option>"
                Else
                    strOption = strOption & "<option value='" & i & "'>" & arrValue(0) & "</option>"
                End If
            End If
        End If
    Next
    Array2Option = strOption
End Function

'**************************************************
'函数名：GetArrFromDictionary
'作  用：从字典表获得区域值
'参  数：strTableName ---- 表名称
'        strFieldName ---- 区域名称
'返回值：查询区域值
'**************************************************
Public Function GetArrFromDictionary(strTableName, strFieldName)
    Dim rsDictionary, strTemp
    Set rsDictionary = Conn.Execute("select FieldValue from PE_Dictionary where TableName='" & strTableName & "' and FieldName='" & strFieldName & "'")
    If rsDictionary.BOF And rsDictionary.EOF Then
        strTemp = ""
    Else
        strTemp = rsDictionary(0) & ""
    End If
    Set rsDictionary = Nothing
    GetArrFromDictionary = Split(strTemp, "$$$")
End Function

'**************************************************
'方法名：PopCalendarInit
'作  用：调用日期js
'**************************************************
Public Sub PopCalendarInit()
    Response.Write "<script language='JavaScript' src='PopCalendar.js'></script>" & vbCrLf
    Response.Write "<script language='JavaScript'>" & vbCrLf
    Response.Write "    PopCalendar = getCalendarInstance()" & vbCrLf
    Response.Write "    PopCalendar.startAt = 0 // 0 - sunday ; 1 - monday" & vbCrLf
    Response.Write "    PopCalendar.showWeekNumber = 0 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showTime = 0 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showToday = 0 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showWeekend = 1 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showHolidays = 1 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showSpecialDay = 1 // 0 - don't show, 1 - show" & vbCrLf
    Response.Write "    PopCalendar.selectWeekend = 0 // 0 - don't Select; 1 - Select" & vbCrLf
    Response.Write "    PopCalendar.selectHoliday = 0 // 0 - don't Select; 1 - Select" & vbCrLf
    Response.Write "    PopCalendar.addCarnival = 0 // 0 - don't Add; 1- Add to Holiday" & vbCrLf
    Response.Write "    PopCalendar.addGoodFriday = 0 // 0 - don't Add; 1- Add to Holiday" & vbCrLf
    Response.Write "    PopCalendar.language = 0 // 0 - Chinese; 1 - English" & vbCrLf
    Response.Write "    PopCalendar.defaultFormat = 'yyyy-mm-dd' //Default Format dd-mm-yyyy" & vbCrLf
    Response.Write "    PopCalendar.fixedX = -1 // x position (-1 if to appear below control)" & vbCrLf
    Response.Write "    PopCalendar.fixedY = -1 // y position (-1 if to appear below control)" & vbCrLf
    Response.Write "    PopCalendar.fade = .5 // 0 - don't fade; .1 to 1 - fade (Only IE) " & vbCrLf
    Response.Write "    PopCalendar.shadow = 1 // 0  - don't shadow, 1 - shadow" & vbCrLf
    Response.Write "    PopCalendar.move = 1 // 0  - don't move, 1 - move (Only IE)" & vbCrLf
    Response.Write "    PopCalendar.saveMovePos = 1  // 0  - don't save, 1 - save" & vbCrLf
    Response.Write "    PopCalendar.centuryLimit = 40 // 1940 - 2039" & vbCrLf
    Response.Write "    PopCalendar.initCalendar()" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub



'**************************************************
'函数名：ShowJS_Main
'作  用：页面管理js(多项诓全选,删除提示)
'参  数：ItemName ---- 项目名称
'返回值：javascript 验证
'**************************************************
Public Sub ShowJS_Main(ItemName)
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
    Response.Write "     if(confirm('确定要删除选中的" & ItemName & "吗？'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

'**************************************************
'函数名：ShowJS_Manage
'作  用：通用频道管理js验证(多项诓全选,删除提示,移动)
'参  数：ItemName ---- 项目名称
'返回值：javascript 验证
'**************************************************
Public Sub ShowJS_Manage(ItemName)
    Dim strJS
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function CheckItem(CB){" & vbCrLf
    Response.Write "  var tagname=(arguments.length>1)?arguments[1]:'TR';" & vbCrLf
    Response.Write "  if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write "    document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (CB.checked){hL(CB,tagname)};else{dL(CB,tagname)};" & vbCrLf
    Response.Write "  var TB=TO=0;" & vbCrLf
    Response.Write "  for (var i=0;i<myform.elements.length;i++) {" & vbCrLf
    Response.Write "    var e=myform.elements[i];" & vbCrLf
    Response.Write "    if ((e.name != 'chkAll') && (e.type=='checkbox')) {" & vbCrLf
    Response.Write "      TB++;" & vbCrLf
    Response.Write "      if (e.checked) TO++;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  myform.chkAll.checked=(TO==TB)?true:false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(form){" & vbCrLf
    Response.Write "  var tagname=(arguments.length>1)?arguments[1]:'TR';" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name != 'chkAll' && e.disabled == false && e.type == 'checkbox') {" & vbCrLf
    Response.Write "      e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "      if (e.checked){hL(e,tagname)};else{dL(e,tagname)};" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function hL(E,tagname){" & vbCrLf
    Response.Write "  while (E.tagName!=tagname) {E=E.parentElement;}" & vbCrLf
    Response.Write "  E.className='tdbg2';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function dL(E,tagname){" & vbCrLf
    Response.Write "  while (E.tagName!=tagname) {E=E.parentElement;}" & vbCrLf
    Response.Write "  E.className='tdbg';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ConfirmDel(){" & vbCrLf
    Response.Write " if(document.myform.Action.value=='Del'){" & vbCrLf
    Response.Write "     if(confirm('确定要删除选中的" & ItemName & "吗？本操作将把选中的" & ItemName & "移到回收站中。必要时您可从回收站中恢复！'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='ConfirmDel'){" & vbCrLf
    Response.Write "     if(confirm('确定要彻底删除选中的" & ItemName & "吗？彻底删除后将不能恢复！'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='ClearRecyclebin'){" & vbCrLf
    Response.Write "     if(confirm('确定要清空回收站？一旦清空将不能恢复！'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='DelFromSpecial'){" & vbCrLf
    Response.Write "     if(confirm('确定要将选中的" & ItemName & "从其所属专题中删除吗？操作成功后" & ItemName & "将不属于任何专题。'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub ShowContentManagePath(RootName)
    Response.Write "您现在的位置：&nbsp;" & ChannelName & "管理&nbsp;&gt;&gt;&nbsp;<a href='" & FileName & "'>" & RootName & "</a>&nbsp;&gt;&gt;&nbsp;"
    If ClassID > 0 Then
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Conn.Execute(sqlPath)
            Do While Not rsPath.EOF
                Response.Write "<a href='" & FileName & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        Response.Write "<a href='" & FileName & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If ManageType = "My" Then
        Response.Write "<font color=red>" & AdminName & "</font> 添加的" & ChannelShortName & ""
    Else
        If Keyword = "" Then
            Select Case Status
            Case -2
                Response.Write "退稿"
            Case -1
                Response.Write "草稿"
            Case 0
                Response.Write "待审核的" & ChannelShortName & "！"
            Case 1
                Response.Write "已审核的" & ChannelShortName & "！"
            Case Else
                Response.Write "所有" & ChannelShortName & "！"
            End Select
        Else
            Select Case strField
            Case "ID"
                Response.Write "ID等于" & Keyword & "</font> "
            Case "Title"
                Response.Write "标题中含有 <font color=red>" & Keyword & "</font> "
            Case "Content"
                Response.Write "内容中含有 <font color=red>" & Keyword & "</font> "
            Case "Author"
                Response.Write "作者姓名中含有 <font color=red>" & Keyword & "</font> "
            Case "Inputer"
                Response.Write "<font color=red>" & Keyword & "</font> 添加"
            Case "Editor"
                Response.Write "<font color=red>" & Keyword & "</font> 审核"
            Case "Keyword"
                Response.Write "关键字为 <font color=red>" & Keyword & "</font> "
            Case "UpdateTime"
                Response.Write "更新时间为 <font color=red>" & Keyword & "</font> "
            Case "SoftName", "PhotoName"
                Response.Write "名称中含有 <font color=red>" & Keyword & "</font> "
            Case "SoftIntro", "PhotoIntro"
                Response.Write "内容中含有 <font color=red>" & Keyword & "</font> "
            Case Else
                Response.Write "名称中含有 <font color=red>" & Keyword & "</font> "
            End Select
            Select Case Status
            Case -2
                Response.Write "的退稿"
            Case -1
                Response.Write "的草稿"
            Case 0
                Response.Write "并且未审核的" & ChannelShortName & "！"
            Case 1
                Response.Write "并且已审核的" & ChannelShortName & "！"
            Case Else
                Response.Write "的" & ChannelShortName & "！"
            End Select
        End If
    End If
End Sub

'**************************************************
'函数名：GetRootClass
'作  用：显示栏目标题栏
'参  数：ChannelID ---- 频道ID
'        RootID ---- 根栏目ID
'        FileName ---- 栏目文件名
'返回值：栏目标题栏
'**************************************************
Public Function GetRootClass()
    Dim sqlRoot, rsRoot, strRoot
    sqlRoot = "select ClassID,ClassName,RootID,Child from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and ClassType=1 order by RootID"
    Set rsRoot = Conn.Execute(sqlRoot)
    If rsRoot.BOF And rsRoot.EOF Then
        strRoot = "还没有任何栏目，请首先添加栏目。"
    Else
        strRoot = "|&nbsp;"
        Do While Not rsRoot.EOF
            If rsRoot(2) = RootID Then
                strRoot = strRoot & "<a href='" & FileName & "&ClassID=" & rsRoot(0) & "'><font color=red>" & rsRoot(1) & "</font></a> | "
            Else
                strRoot = strRoot & "<a href='" & FileName & "&ClassID=" & rsRoot(0) & "'>" & rsRoot(1) & "</a> | "
            End If
            rsRoot.MoveNext
        Loop
    End If
    rsRoot.Close
    Set rsRoot = Nothing
    GetRootClass = strRoot
End Function

'**************************************************
'函数名：GetChild_Root
'作  用：显示栏目子栏目标题栏
'参  数：ChannelID ---- 频道ID
'        RootID ---- 根栏目ID
'        ClassID ---- 栏目ID
'        ParentPath ---- 父路径
'        Depth ---- 栏目深度
'        FileName ---- 栏目文件名
'返回值：子栏目标题栏
'**************************************************
Public Function GetChild_Root()
    Dim sqlChild, rsChild, arrParentPath, isCurrent, strChild, i
    If RootID <= 0 Then
        GetChild_Root = ""
        Exit Function
    End If
    sqlChild = "select ClassID,ClassName,Child from PE_Class where ChannelID=" & ChannelID & " and Depth=1 and RootID=" & RootID & " order by OrderID"
    Set rsChild = Conn.Execute(sqlChild)
    If Not (rsChild.BOF And rsChild.EOF) Then
        i = 1
        arrParentPath = Split(ParentPath, ",")
        strChild = "<tr class='tdbg'><td>"
        Do While Not rsChild.EOF
            If Depth <= 1 Then
                If rsChild(0) = ClassID Then
                    isCurrent = True
                Else
                    isCurrent = False
                End If
            Else
                If PE_CLng(arrParentPath(2)) = rsChild(0) Then
                    isCurrent = True
                Else
                    isCurrent = False
                End If
            End If
            If isCurrent = True Then
                strChild = strChild & "&nbsp;&nbsp;<a href='" & FileName & "&ClassID=" & rsChild(0) & "'><font color='red'>" & rsChild(1) & "</font></a>"
            Else
                strChild = strChild & "&nbsp;&nbsp;<a href='" & FileName & "&ClassID=" & rsChild(0) & "'>" & rsChild(1) & "</a>"
            End If
            If rsChild(2) > 0 Then
                strChild = strChild & "(" & rsChild(2) & ")"
            End If
            If i Mod 8 = 0 Then
                strChild = strChild & "<br>"
            Else
                strChild = strChild & "&nbsp;&nbsp;"
            End If
            rsChild.MoveNext
            i = i + 1
        Loop
        strChild = strChild & "</td></tr>"
    End If
    rsChild.Close
    Set rsChild = Nothing
    GetChild_Root = strChild
End Function

Function GetSpecial_Option(SpecialID)
    Dim sqlSpecial, rsSpecial, strOption, strOptionTemp
    sqlSpecial = "select ChannelID,SpecialID,SpecialName,OrderID from PE_Special where ChannelID=0 or ChannelID=" & ChannelID & "   order by ChannelID,OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    Do While Not rsSpecial.EOF
        If rsSpecial("ChannelID") > 0 Then
            strOptionTemp = rsSpecial("SpecialName") & "（本频道）"
        Else
            strOptionTemp = rsSpecial("SpecialName") & "（全站）"
        End If
        If FoundInArr(SpecialID, rsSpecial("SpecialID"), ",") = True Then
            strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "' selected>" & strOptionTemp & "</option>"
        Else
            strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "'>" & strOptionTemp & "</option>"
        End If
        rsSpecial.MoveNext
    Loop
    rsSpecial.Close
    Set rsSpecial = Nothing
    GetSpecial_Option = strOption
End Function

'**************************************************
'函数名：GetStars
'作  用：显示等级★数量
'参  数：Stars ---- 项目名称
'返回值：下拉菜单数据
'**************************************************
Public Function GetStars(Stars)
    Dim strTemp
    strTemp = strTemp & "<option value='5'"
    If Stars = 5 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">★★★★★</option>"
    strTemp = strTemp & "<option value='4'"
    If Stars = 4 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">★★★★</option>"
    strTemp = strTemp & "<option value='3'"
    If Stars = 3 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">★★★</option>"
    strTemp = strTemp & "<option value='2'"
    If Stars = 2 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">★★</option>"
    strTemp = strTemp & "<option value='1'"
    If Stars = 1 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">★</option>"
    strTemp = strTemp & "<option value='0'"
    If Stars = 0 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">无</option>"
    GetStars = strTemp
End Function

'**************************************************
'函数名：GetAuthorList
'作  用：显示作者
'参  数：ChannelID ---- 频道ID
'        UserName ---- 用户名称
'返回值：【未知】【佚名】【管理员】【更多】
'**************************************************
Public Function GetAuthorList(FilePrefix, ChannelID, UserName)
    Dim Author, strAuthorList

    Author = Trim(Session("Author"))
    Dim strDefAuthor, strUnKnowAuthor
    strDefAuthor = XmlText("BaseText", "DefAuthor", "佚名")
    strUnKnowAuthor = XmlText("BaseText", "UnKnowAuthor", "未知")

    strAuthorList = "<font color='blue'><="
    strAuthorList = strAuthorList & "【<font color='green' onclick=""document.myform.Author.value='" & strDefAuthor & "'"" style=""cursor:hand;"">" & strDefAuthor & "</font>】"
    strAuthorList = strAuthorList & "【<font color='green' onclick=""document.myform.Author.value='" & strUnKnowAuthor & "'"" style=""cursor:hand;"">" & strUnKnowAuthor & "</font>】"
    strAuthorList = strAuthorList & "【<font color='green' onclick=""document.myform.Author.value='" & UserName & "'"" style=""cursor:hand;"">" & UserName & "</font>】"
    If Author <> "" And Author <> strDefAuthor And Author <> strUnKnowAuthor And Author <> UserName Then
        strAuthorList = strAuthorList & "【<font color='green' onclick=""document.myform.Author.value='" & FilterJS(Replace(Author, "'", "")) & "'"" style=""cursor:hand;"">" & FilterJS(Author) & "</font>】"
    End If
    strAuthorList = strAuthorList & "【<font color='green' onclick=""window.open('" & FilePrefix & "_SourceList.asp?TypeSelect=AuthorList&ChannelID=" & ChannelID & "', 'AuthorList', 'width=600,height=450,resizable=0,scrollbars=yes');"" style=""cursor:hand;"">更多</font>】"
    strAuthorList = strAuthorList & "</font>"
    GetAuthorList = strAuthorList
End Function

'**************************************************
'函数名：GetCopyFromList
'作  用：显示来源
'参  数：FilePrefix ----访问身份 Admin,User
'        ChannelID ---- 频道ID
'返回值：<=【本站原创】【更多】
'**************************************************
Public Function GetCopyFromList(FilePrefix, ChannelID)
    Dim CopyFrom, strCopyFromList
    CopyFrom = Trim(Session("CopyFrom"))
    Dim strDefCopyFrom
    strDefCopyFrom = XmlText("BaseText", "DefCopyFrom", "本站原创")

    strCopyFromList = "<font color='blue'><="
    strCopyFromList = strCopyFromList & "【<font color='green' onclick=""document.myform.CopyFrom.value='" & strDefCopyFrom & "'"" style=""cursor:hand;"">" & strDefCopyFrom & "</font>】"
    If CopyFrom <> "" And CopyFrom <> strDefCopyFrom Then
        strCopyFromList = strCopyFromList & "【<font color='green' onclick=""document.myform.CopyFrom.value='" & FilterJS(Replace(CopyFrom, "'", "")) & "'"" style=""cursor:hand;"">" & FilterJS(CopyFrom) & "</font>】"
    End If
    strCopyFromList = strCopyFromList & "【<font color='green' onclick=""window.open('" & FilePrefix & "_SourceList.asp?TypeSelect=CopyFromList&ChannelID=" & ChannelID & "', 'CopyFromList', 'width=600,height=450,resizable=0,scrollbars=yes');"" style=""cursor:hand;"">更多</font>】"
    strCopyFromList = strCopyFromList & "</font>"
    GetCopyFromList = strCopyFromList
End Function

'**************************************************
'函数名：GetKeywordList
'作  用：显示关键字
'参  数：FilePrefix ----访问身份 Admin,User
'        ChannelID ---- 频道ID
'返回值：显示频道中前4个关键字 +【更多】
'**************************************************
Public Function GetKeywordList(FilePrefix, ChannelID)
    Dim sqlGetKey, rsGetKey, strKeywordList
    strKeywordList = "<font color='blue'><="
    sqlGetKey = "select top 4 * from PE_NewKeys where ChannelID=" & ChannelID & " or ChannelID=0 order by LastUseTime Desc"
    Set rsGetKey = Conn.Execute(sqlGetKey)
    If rsGetKey.BOF And rsGetKey.EOF Then
        strKeywordList = strKeywordList & "【<font color='green'>无常用关键字</font>】"
    Else
        Do While Not rsGetKey.EOF
            strKeywordList = strKeywordList & "【<font color='green' onclick=""document.myform.Keyword.value+=(document.myform.Keyword.value==''?'':'|')+'" & FilterJS(Replace(rsGetKey("KeyText"), "'", "")) & "'"" style=""cursor:hand;"">" & FilterJS(rsGetKey("KeyText")) & "</font>】"
            rsGetKey.MoveNext
        Loop
    End If
    rsGetKey.Close
    Set rsGetKey = Nothing
    strKeywordList = strKeywordList & "【<font color='green' onclick=""window.open('" & FilePrefix & "_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=KeyList', 'KeyList', 'width=600,height=450,resizable=0,scrollbars=yes');"" style=""cursor:hand;"">更多</font>】"
    strKeywordList = strKeywordList & "</font>"
    GetKeywordList = strKeywordList
End Function

'**************************************************
'方法名：SaveKeyword
'作  用：保存关键字
'**************************************************
Public Sub SaveKeyword(strKeyword)
    Dim rsKeyword, sqlKeyword, arrKeyword, i
    strKeyword = ReplaceBadChar(strKeyword) 
    If strKeyword = "" Then Exit Sub
    arrKeyword = Split(strKeyword, "|")
    Set rsKeyword = Server.CreateObject("adodb.recordset")
    For i = 0 To UBound(arrKeyword)
        sqlKeyword = "Select ChannelID,KeyText,Hits,LastUseTime from PE_NewKeys Where ChannelID=" & ChannelID & " and KeyText='" & arrKeyword(i) & "'"
        rsKeyword.Open sqlKeyword, Conn, 1, 3
        If rsKeyword.BOF And rsKeyword.EOF Then
            If 	arrKeyword(i)<>"" then	
                rsKeyword.addnew
                rsKeyword("ChannelID") = ChannelID
                rsKeyword("KeyText") = arrKeyword(i)
                rsKeyword("Hits") = 0
                rsKeyword("LastUseTime") = Now()
                rsKeyword.Update
            End If				
        Else
            Do While Not rsKeyword.EOF
                If arrKeyword(i)<>"" then				
                    rsKeyword("Hits") = rsKeyword("Hits") + 1
                    rsKeyword("LastUseTime") = Now()
                    rsKeyword.Update
                End If 
                rsKeyword.MoveNext				
            Loop
        End If
        rsKeyword.Close
    Next
    Set rsKeyword = Nothing
End Sub

'**************************************************
'函数名：ReplaceText
'作  用：过滤非法字符串
'参  数：iText-----输入字符串
'返回值：替换后字符串
'**************************************************
Function ReplaceText(iText, iType)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol
    If PE_Cache.GetValue("Site_ReplaceText") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=1 order by Priority"
        rsKey.Open sqlKey, Conn, 1, 1
        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_ReplaceText", rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceText = iText
            Exit Function
        End If
    End If
    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_ReplaceText"), "@@@")
    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        If Int(Keycol(3)) = 0 Or Int(Keycol(3)) = iType Then rText = PE_Replace(rText, Keycol(0), Keycol(1))
    Next
    ReplaceText = rText
End Function

'**************************************************
'函数名：ShowClassPath
'作  用：显示栏目路径
'参  数：无
'返回值：显示栏目
'**************************************************
Public Function ShowClassPath()
    If ParentPath = "" Or IsNull(ParentPath) Then
        ShowClassPath = "不属于任何栏目"
        Exit Function
    End If
    Dim strPath
    If Depth > 0 Then
        Dim rsPath
        Set rsPath = Conn.Execute("select * from PE_Class where ClassID in (" & ParentPath & ") order by Depth")
        Do While Not rsPath.EOF
            strPath = strPath & rsPath("ClassName") & " >> "
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    strPath = strPath & ClassName
    ShowClassPath = strPath
End Function



'**************************************************
'函数名：GetUserGroup
'作  用：显示用户组导航
'参  数：arrGroupID ---- 指定默认用户组
'返回值：用户组表格导航
'**************************************************
Function GetUserGroup(arrGroupID, strDisabled)
    If IsNull(arrGroupID) Then Exit Function
    Dim rsGroup, strGroup, i
    strGroup = "<table width='95%' align='right'><tr>"
    Set rsGroup = Conn.Execute("select GroupID,GroupName from PE_UserGroup where GroupID<>-1 order by GroupType asc,GroupID asc")
    Do While Not rsGroup.EOF
        strGroup = strGroup & "<td><input type='checkbox' name='GroupID' value='" & rsGroup(0) & "'" & strDisabled
        If FoundInArr(arrGroupID, rsGroup(0), ",") = True Then
            strGroup = strGroup & " checked"
        End If
        strGroup = strGroup & ">" & rsGroup(1) & "</td>"
        i = i + 1
        rsGroup.MoveNext
        If i Mod 5 = 0 And Not rsGroup.EOF Then
            strGroup = strGroup & "</tr><tr>"
        End If
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
    strGroup = strGroup & "</table>"
    GetUserGroup = strGroup
End Function

Sub UpdateChannelData(ByVal iChannelID)
    Dim rsChannel, sqlChannel, trs, ModuleName
    Dim ItemCount, ItemChecked, CommentCount, SpecialCount
    sqlChannel = "select ChannelID,ModuleType,ItemCount,ItemChecked,CommentCount,SpecialCount from PE_Channel"
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    If IsValidID(iChannelID) = False Then
        iChannelID = ""
    End If
    
    If InStr(iChannelID, ",") > 0 Then
        sqlChannel = sqlChannel & " where ChannelID in (" & iChannelID & ")"
    ElseIf PE_CLng(iChannelID) > 0 Then
        sqlChannel = sqlChannel & " where ChannelID=" & iChannelID & ""
    Else
        sqlChannel = sqlChannel & " where ChannelType<=1 order by ChannelID"
    End If
    rsChannel.Open sqlChannel, Conn, 1, 3
    Do While Not rsChannel.EOF
        Select Case rsChannel("ModuleType")
        Case 7
            Dim i, HouseTableName
            For i = 1 To 5
                Select Case i
                Case 1
                    HouseTableName = "PE_HouseCS"
                Case 2
                    HouseTableName = "PE_HouseCZ"
                Case 3
                    HouseTableName = "PE_HouseQG"
                Case 4
                    HouseTableName = "PE_HouseQZ"
                Case 5
                    HouseTableName = "PE_HouseHZ"
                End Select
                Set trs = Conn.Execute("select Count(HouseID) from " & HouseTableName & " where Deleted=" & PE_False & "")
                ItemCount = ItemCount + trs(0)
                Set trs = Nothing
                Set trs = Conn.Execute("select Count(HouseID) from " & HouseTableName & " where Passed=" & PE_True & " and Deleted=" & PE_False & "")
                ItemChecked = ItemChecked + trs(0)
                Set trs = Nothing
            Next
            rsChannel("ItemCount") = ItemCount
            rsChannel("ItemChecked") = ItemChecked
        Case 8
            Set trs = Conn.Execute("select Count(PositionID) from PE_Position ")
            ItemCount = ItemCount + trs(0)
            Set trs = Nothing
            rsChannel("ItemCount") = ItemCount
        Case Else
            Select Case rsChannel("ModuleType")
            Case 1
                ModuleName = "Article"
            Case 2
                ModuleName = "Soft"
            Case 3
                ModuleName = "Photo"
            Case 5
                ModuleName = "Product"
            Case 6
                ModuleName = "Supply"
            End Select
            Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=" & PE_False & "")
            ItemCount = trs(0)
            Set trs = Nothing

            If ModuleName = "Product" Then
                Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & rsChannel("ChannelID") & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & "")
            Else
                Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & rsChannel("ChannelID") & " and Status=3 and Deleted=" & PE_False & "")
            End If
            ItemChecked = trs(0)
            Set trs = Nothing

            Set trs = Conn.Execute("select Count(CommentID) from PE_Comment C inner join PE_" & ModuleName & " I on C.InfoID=I." & ModuleName & "ID where I.ChannelID=" & rsChannel("ChannelID") & "")
            CommentCount = trs(0)
            Set trs = Nothing

            Set trs = Conn.Execute("select Count(SpecialID) from PE_Special where ChannelID=" & rsChannel("ChannelID") & "")
            SpecialCount = trs(0)
            Set trs = Nothing

            rsChannel("ItemCount") = ItemCount
            rsChannel("ItemChecked") = ItemChecked
            rsChannel("CommentCount") = CommentCount
            rsChannel("SpecialCount") = SpecialCount
        End Select
        rsChannel.Update
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
End Sub

Sub UpdateUserData(UserType, UserName, BeginID, EndID)
    Dim sqlUser, rsUser, trs, PostItems, PassedItems
    If UserType = 0 Then
        If InStr(UserName, ",") > 0 Then
            sqlUser = "select * from PE_User where UserName in ('" & Replace(UserName, ",", "','") & "')"
        Else
            sqlUser = "select * from PE_User where UserName='" & UserName & "'"
        End If
    Else
        sqlUser = "select * from PE_User where UserID>=" & BeginID & " and UserID<=" & EndID
    End If
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    rsUser.Open sqlUser, Conn, 1, 3
    Do While Not rsUser.EOF
        Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Deleted=" & PE_False & " and  Inputer='" & rsUser("UserName") & "'")
        If IsNull(trs(0)) Then
            PostItems = 0
        Else
            PostItems = trs(0)
        End If
        Set trs = Nothing
        Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Deleted=" & PE_False & " and Status=3 and Inputer='" & rsUser("UserName") & "'")
        If IsNull(trs(0)) Then
            PassedItems = 0
        Else
            PassedItems = trs(0)
        End If
        Set trs = Nothing
        Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Deleted=" & PE_False & " and  Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PostItems = PostItems + trs(0)
        Set trs = Nothing
        Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Deleted=" & PE_False & " and  Status=3 and Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PassedItems = PassedItems + trs(0)
        Set trs = Nothing
        Set trs = Conn.Execute("select count(PhotoID) from PE_Photo where Deleted=" & PE_False & " and  Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PostItems = PostItems + trs(0)
        Set trs = Nothing
        Set trs = Conn.Execute("select count(PhotoID) from PE_Photo where Deleted=" & PE_False & " and  Status=3 and Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PassedItems = PassedItems + trs(0)
        Set trs = Nothing
        
        rsUser("PostItems") = PostItems
        rsUser("PassedItems") = PassedItems
        rsUser.Update
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing
End Sub

%>
