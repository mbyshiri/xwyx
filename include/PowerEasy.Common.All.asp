<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'判断当前访问者是否已经登录，若已登录，则读取数据并做必要赋值
Function CheckUserLogined()
    Dim UserPassword, LastPassword
    Dim rsUser, sqlUser
    UserID = 0
    GroupID = 0
    Balance = 0
    UserPoint = 0
    UserExp = 0
    LoginTimes = 0
    UserChargeType = 0

    CheckUserLogined = False
    
    UserName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserName")))
    UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
    LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))
    If (UserName = "" Or UserPassword = "" Or LastPassword = "") Then
        ReDim UserSetting(50)
        CheckUserLogined = False
        Exit Function
    End If
    
    sqlUser = "SELECT UserID,UserName,GroupID,LoginTimes FROM PE_User WHERE UserName='" & UserName & "' AND UserPassword='" & UserPassword & "' AND LastPassword='" & LastPassword & "' and IsLocked=" & PE_False & ""
    Set rsUser = Conn.Execute(sqlUser)
    If rsUser.BOF And rsUser.EOF Then
        ReDim UserSetting(50)
        CheckUserLogined = False
    Else
        UserName = rsUser("UserName")
        CheckUserLogined = True
        UserID = rsUser("UserID")
        GroupID = rsUser("GroupID")
        LoginTimes = rsUser("LoginTimes")
    End If
    Set rsUser = Nothing
End Function

'给用户的相应变量赋值
Sub GetUser(sUserName)
    Dim rsUser, rsGroup
    Set rsUser = Conn.Execute("SELECT * FROM PE_User WHERE UserName='" & sUserName & "'")
    If Not (rsUser.BOF And rsUser.EOF) Then
        UserID = rsUser("UserID")
        GroupID = rsUser("GroupID")
        UserType = rsUser("UserType")
        CompanyID = rsUser("CompanyID")
        ContacterID = rsUser("ContacterID")
        ClientID = rsUser("ClientID")
        Balance = rsUser("Balance")
        UserPoint = rsUser("UserPoint")
        UserExp = rsUser("UserExp")
        ValidNum = rsUser("ValidNum")
        ValidUnit = rsUser("ValidUnit")
        BeginTime = rsUser("BeginTime")
        ValidDays = ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime"))
        email = rsUser("Email")
        UnsignedItems = rsUser("UnsignedItems")
        If PresentExpPerLogin > 0 Then
        If DateDiff("D", rsUser("LastPresentTime"), Now()) > 0 Or IsNull(rsUser("LastPresentTime")) Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExpPerLogin & ",LastPresentTime=" & PE_Now & " where UserID=" & UserID & "")
            End If
        End If
        If PE_CLng(Session("UserID")) = 0 Then
            Conn.Execute ("update PE_User set LastLoginIP='" & UserTrueIP & "',LastLoginTime=" & PE_Now & ",LoginTimes=LoginTimes+1 where UserID=" & UserID & "")
            Session("UserID") = UserID
        End If
        If rsUser("Blog") = True Then
            BlogFlag = True
        Else
            BlogFlag = False
        End If
        Set rsGroup = Conn.Execute("select * from PE_UserGroup where GroupID=" & rsUser("GroupID") & "")
        GroupName = rsGroup("GroupName")
        GroupType = rsGroup("GroupType")
        If rsUser("SpecialPermission") = True Then
            arrClass_Browse = Trim(rsUser("arrClass_Browse"))
            arrClass_View = Trim(rsUser("arrClass_View"))
            arrClass_Input = Trim(rsUser("arrClass_Input"))
            UserSetting = Split(Trim(rsUser("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        Else
            arrClass_Browse = Trim(rsGroup("arrClass_Browse"))
            arrClass_View = Trim(rsGroup("arrClass_View"))
            arrClass_Input = Trim(rsGroup("arrClass_Input"))
            UserSetting = Split(Trim(rsGroup("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        End If
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
        If UserSetting(12) = 1 Then
            IsOffer = "是"
        Else
            IsOffer = "否"
        End If
        UserChargeType = PE_CLng(UserSetting(14))
        Dim Message
        Set Message = Conn.Execute("select Count(0) from PE_Message where Incept = '" & UserName & "' and delR=0 and Flag=0 and IsSend=1")
        If Message.EOF And Message.Bof Then
            UnreadMsg = 0
        Else
            UnreadMsg = Message(0)
        End If
        Set Message = Nothing
    End If
    Set rsUser = Nothing
End Sub

'**************************************************
'函数名：GetSubStr
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'        strlen ----截取长度
'        bShowPoint ---- 是否显示省略号
'返回值：截取后的字符串
'**************************************************
Function GetSubStr(ByVal str, ByVal strlen, bShowPoint)
    If IsNull(str) Or str = ""  Then
        GetSubStr = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(str)
    t = 0
    strTemp = str
    strlen = PE_CLng(strlen)
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next
    str = Replace(Replace(Replace(Replace(str, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
    strTemp = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
    If strTemp <> str And bShowPoint = True Then
        strTemp = strTemp & "…"
    End If
    GetSubStr = strTemp
End Function

'**************************************************
'函数名：GetStrLen
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
Function GetStrLen(str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("中国") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(str)
        t = l
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        GetStrLen = t
    Else
        GetStrLen = Len(str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

Function Charlong(ByVal str)
    If str = "" Then
        Charlong = 0
        Exit Function
    End If
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    
    Charlong = GetStrLen(str)
End Function

'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
Function JoinChar(ByVal strUrl)
    If strUrl = "" Then
        JoinChar = ""
        Exit Function
    End If
    If InStr(strUrl, "?") < Len(strUrl) Then
        If InStr(strUrl, "?") > 1 Then
            If InStr(strUrl, "&") < Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function

'**************************************************
'函数名：ShowPage
'作  用：显示“上一页 下一页”等信息
'参  数：sFileName  ----链接地址
'        TotalNumber ----总数量
'        MaxPerPage  ----每页数量
'        CurrentPage ----当前页
'        ShowTotal   ----是否显示总数量
'        ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
'        strUnit     ----计数单位
'        ShowMaxPerPage  ----是否显示每页信息量选项框
'返回值：“上一页 下一页”等信息的HTML代码
'**************************************************
Function ShowPage(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
    strTemp = "<div class=""show_page"">"
    If ShowTotal = True Then
        strTemp = strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;&nbsp;"
    End If
    
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "首页 | 上一页 |"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a> |"
        strTemp = strTemp & "  <a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a> | "
    End If
    strTemp = strTemp & " "
    If ShowAllPages = True Then
        Dim Jmaxpages
        If (CurrentPage - 4) <= 0 Or TotalPage < 10 Then
            Jmaxpages = 1
            Do While (Jmaxpages < 10)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<font color=""FF0000"">" & Jmaxpages & "</font> "
                Else
                    If strUrl <> "" Then
                        strTemp = strTemp & "<a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a> "
                    End If
                End If
                If Jmaxpages = TotalPage Then Exit Do
                Jmaxpages = Jmaxpages + 1
            Loop
        ElseIf (CurrentPage + 4) >= TotalPage Then
            Jmaxpages = TotalPage - 8
            Do While (Jmaxpages <= TotalPage)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<font color=""FF0000"">" & Jmaxpages & "</font> "
                Else
                    If strUrl <> "" Then
                        strTemp = strTemp & "<a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a> "
                    End If
                End If
                Jmaxpages = Jmaxpages + 1
            Loop
        Else
            Jmaxpages = CurrentPage - 4
            Do While (Jmaxpages < CurrentPage + 5)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<font color=""FF0000"">" & Jmaxpages & "</font> "
                Else
                    If strUrl <> "" Then
                        strTemp = strTemp & "<a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a> "
                    End If
                End If
                Jmaxpages = Jmaxpages + 1
            Loop
        End If
    End If
    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "| 下一页 | 尾页"
    Else
        strTemp = strTemp & " | <a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a> |"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>  尾页</a>"
    End If
	If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;&nbsp;&nbsp;<Input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"">" & strUnit & "/页"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/页"
    End If
    If ShowAllPages = True Then
            strTemp = strTemp & "&nbsp;&nbsp;转到第<Input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;"">页"
    End If
    strTemp = strTemp & "</div>"
    ShowPage = strTemp
End Function


'**************************************************
'函数名：ShowPage_en
'作  用：显示英文“上一页 下一页”等信息
'参  数：sFileName  ----链接地址
'        TotalNumber ----总数量
'        MaxPerPage  ----每页数量
'        CurrentPage ----当前页
'        ShowTotal   ----是否显示总数量
'        ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
'        strUnit     ----计数单位
'        ShowMaxPerPage  ----是否显示每页信息量选项框
'返回值：“上一页 下一页”等信息的HTML代码
'**************************************************
Function ShowPage_en(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage_en = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
        
    strTemp = "<div class=""show_page"">"
    If ShowTotal = True Then
        strTemp = strTemp & "Total <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
	
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "FirstPage PreviousPage&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>FirstPage</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>PreviousPage</a>&nbsp;"
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "NextPage LastPage"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>NextPage</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>LastPage</a>"
    End If
    strTemp = strTemp & " CurrentPage: <strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong> "
    If ShowMaxPerPage = True Then
        strTemp = strTemp & " <Input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"">" & strUnit & "/Page"
    Else
        strTemp = strTemp & " <b>" & MaxPerPage & "</b>" & strUnit & "/Page"
    End If
    If ShowAllPages = True Then
        If TotalPage > 20 Then
            strTemp = strTemp & "&nbsp;&nbsp;GoTo Page:<Input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;"">"
        Else
            strTemp = strTemp & "&nbsp;GoTo:<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To TotalPage
               strTemp = strTemp & "<option value='" & i & "'"
               If PE_CLng(CurrentPage) = PE_CLng(i) Then strTemp = strTemp & " selected "
               strTemp = strTemp & ">Page" & i & "</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
    End If
    strTemp = strTemp & "</div>"
    ShowPage_en = strTemp
End Function



'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'        False ----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = CreateObject(strClassString)
    If Err.Number = 0 Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function


'**************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'**************************************************
Sub WriteErrMsg(sErrMsg, sComeUrl)
    Response.Write "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='" & strInstallDir & "images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & sErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        Response.Write "<a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a>"
    Else
        Response.Write "<a href='javascript:window.close();'>【关闭】</a>"
    End If
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub

'**************************************************
'过程名：WriteSuccessMsg
'作  用：显示成功提示信息
'参  数：无
'**************************************************
Sub WriteSuccessMsg(sSuccessMsg, sComeUrl)
    Response.Write "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='" & strInstallDir & "images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>恭喜你！</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & sSuccessMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        Response.Write "<a href='" & sComeUrl & "'>&lt;&lt; 返回上一页</a>"
    Else
        Response.Write "<a href='javascript:window.close();'>【关闭】</a>"
    End If
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub

'**************************************************
'函数名：FoundInArr
'作  用：检测数组中是否有指定的数值
'参  数：strArr ----- 调入的数组
'        strItem  ----- 检测的字符
'        strSplit  ----- 分割字符
'返回值：True  ----有
'        False ----没有
'**************************************************
Function FoundInArr(strArr, strItem, strSplit)
    Dim arrTemp, arrTemp2, i, j
    FoundInArr = False
    If IsNull(strArr) Or IsNull(strItem) Or Trim(strArr) = "" Or Trim(strItem) = "" Then
        Exit Function
    End If
    If IsNull(strSplit) Or strSplit = "" Then
        strSplit = ","
    End If
    If InStr(Trim(strArr), strSplit) > 0 Then
        If InStr(Trim(strItem), strSplit) > 0 Then
            arrTemp = Split(strArr, strSplit)
            arrTemp2 = Split(strItem, strSplit)
            For i = 0 To UBound(arrTemp)
                For j = 0 To UBound(arrTemp2)
                    If LCase(Trim(arrTemp2(j))) <> "" And LCase(Trim(arrTemp(i))) <> "" And LCase(Trim(arrTemp2(j))) = LCase(Trim(arrTemp(i))) Then
                        FoundInArr = True
                        Exit Function
                    End If
                Next
            Next
        Else
            arrTemp = Split(strArr, strSplit)
            For i = 0 To UBound(arrTemp)
                If LCase(Trim(arrTemp(i))) = LCase(Trim(strItem)) Then
                    FoundInArr = True
                    Exit Function
                End If
            Next
        End If
    Else
        If LCase(Trim(strArr)) = LCase(Trim(strItem)) Then
            FoundInArr = True
        End If
    End If
End Function

'**************************************************
'函数名：GetRndPassword
'作  用：得到指定位数的随机数密码
'参  数：PasswordLen ---- 位数
'返回值：密码字符串
'**************************************************
Function GetRndPassword(PasswordLen)
    Dim Ran, i, strPassword
    strPassword = ""
    For i = 1 To PasswordLen
        Randomize
        Ran = CInt(Rnd * 2)
        Randomize
        If Ran = 0 Then
            Ran = CInt(Rnd * 25) + 97
            strPassword = strPassword & UCase(Chr(Ran))
        ElseIf Ran = 1 Then
            Ran = CInt(Rnd * 9)
            strPassword = strPassword & Ran
        ElseIf Ran = 2 Then
            Ran = CInt(Rnd * 25) + 97
            strPassword = strPassword & Chr(Ran)
        End If
    Next
    GetRndPassword = strPassword
End Function

'**************************************************
'函数名：GetRndNum
'作  用：产生制定位数的随机数
'参  数：iLength ---- 随即数的位数
'返回值：随机数
'**************************************************
Function GetRndNum(iLength)
    Dim i, str1
    For i = 1 To (iLength \ 5 + 1)
        Randomize
        str1 = str1 & CStr(CLng(Rnd * 90000) + 10000)
    Next
    GetRndNum = Left(str1, iLength)
End Function

'**************************************************
'函数名：GetIDByDefault
'作  用：获取ID值，如果ID为0，则使用缺省值
'参  数：ItemID ---- 项目ID值
'        DefaultID ---- 缺省ID值
'**************************************************
Function GetIDByDefault(ItemID, DefaultID)
    Dim iItemID
    iItemID = ItemID
    If iItemID = 0 Then iItemID = DefaultID
    If IsNull(iItemID) Then iItemID = 0
    GetIDByDefault = iItemID
End Function




'**************************************************
'函数名：FillInArrStr
'作  用：使用一个用逗号分隔的字符串来填充另外一个逗号分隔的字符串，使其达到指定的项目数
'参  数：strSource ---- 原字符串
'        strFill ---- 填充字符串
'        ItemNum ---- 指定填充后的项目数
'返回值：填充后的字符串
'**************************************************
Function FillInArrStr(ByVal strSource, ByVal strFill, ItemNum)
    Dim arrSource, arrFill, SourceItemNum, FillItemNum, i
    If IsNull(strSource) Or IsNull(strFill) Then
        FillInArrStr = ""
        Exit Function
    End If
    arrSource = Split(strSource, ",")
    arrFill = Split(strFill, ",")
    SourceItemNum = UBound(arrSource) + 1
    FillItemNum = UBound(arrFill) + 1
    If SourceItemNum < ItemNum And SourceItemNum + FillItemNum >= ItemNum Then
        For i = 0 To ItemNum - SourceItemNum - 1
            strSource = strSource & "," & arrFill(SourceItemNum + FillItemNum - ItemNum + i)
        Next
    End If
    FillInArrStr = strSource
End Function

'**************************************************
'函数名：XmlText
'作  用：从语言包中读取指定节点的值
'参  数：iBigNode ---- 大节点
'        iSmallNode ---- 小节点
'        DefChar ---- 默认值
'返回值：语言包中指定节点的值
'**************************************************
Function XmlText(ByVal iBigNode, ByVal iSmallNode, ByVal DefChar)
    Dim LangRoot, LangSub
    If IsNull(iBigNode) Or IsNull(iSmallNode) Then
        XmlText = DefChar
    Else
        Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
        If LangRoot.Length = 0 Then
            XmlText = DefChar
        Else
            Set LangSub = LangRoot(0).getElementsByTagName(iSmallNode)
            If LangSub.Length = 0 Then
                XmlText = DefChar
            Else
                XmlText = LangSub(0).text
            End If
        End If
        Set LangRoot = Nothing
    End If
End Function


'**************************************************
'函数名：GetFirstSeparatorToEnd
'作  用：截取从第一个分隔符到结尾的字符串
'参  数：str   ----原字符串
'        separator ----分隔符
'返回值：截取后的字符串
'**************************************************
Function GetFirstSeparatorToEnd(ByVal str, separator)
    GetFirstSeparatorToEnd = Right(str, Len(str) - InStr(str, separator))
End Function

'**************************************************
'函数名：ChkValidDays
'作  用：有效期的函数
'参  数：iValidNum ----有效期
'        iValidUnit ----有效期单位
'        iBeginTime ---- 开始计算日期
'返回值：剩余的有效天数
'**************************************************
Function ChkValidDays(iValidNum, iValidUnit, iBeginTime)
    If (iValidNum = "" Or IsNumeric(iValidNum) = False Or iValidUnit = "" Or IsNumeric(iValidUnit) = False Or iBeginTime = "" Or IsDate(iBeginTime) = False) Then
        ChkValidDays = 0
        Exit Function
    End If
    Dim tmpDate, arrInterval
    arrInterval = Array("h", "D", "m", "yyyy")
    If iValidNum = -1 Then
        ChkValidDays = 99999
    Else
        tmpDate = DateAdd(arrInterval(iValidUnit), iValidNum, iBeginTime)
        ChkValidDays = DateDiff("D", Date, tmpDate)
    End If
End Function

'**************************************************
'函数名：GetNumString
'作  用：获得项目随即数
'返回值：随机无重复的数字(用于上传,生成)
'**************************************************
Function GetNumString()
    Dim v_ymd, v_hms, v_mmm
    v_ymd = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
    v_hms = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    Randomize
    v_mmm = Right("0" & CStr(CLng(99 * Rnd) + 1), 2)
    GetNumString = v_ymd & v_hms & v_mmm
End Function

'**************************************************
'函数名：GetMinID
'作  用：取某一表某一字段中的最大值
'参  数：SheetName ----查询表
'        FieldName ----查询字段
'返回值：该字段最小值
'**************************************************
Function GetMinID(SheetName, FieldName)
    Dim mrs
    Set mrs = Conn.Execute("select min(" & FieldName & ") from " & SheetName & "")
    If IsNull(mrs(0)) Then
        GetMinID = 1
    Else
        GetMinID = mrs(0)
    End If
    Set mrs = Nothing
End Function

'**************************************************
'函数名：GetNewID
'作  用：取某一表某一字段中的最大值+1
'参  数：SheetName ----查询表
'        FieldName ----查询字段
'返回值：该字段最大值+1
'**************************************************
Function GetNewID(SheetName, FieldName)
    Dim mrs
    Set mrs = Conn.Execute("select max(" & FieldName & ") from " & SheetName & "")
    If IsNull(mrs(0)) Then
        GetNewID = 1
    Else
        GetNewID = mrs(0) + 1
    End If
    Set mrs = Nothing
End Function

'**************************************************
'函数名：PE_Replace
'作  用：容错替换
'参  数：expression ---- 主数据
'        find ---- 被替换的字符
'        replacewith ---- 替换后的字符
'返回值：容错后的替换字符串,如果 replacewith 空字符,被替换的字符 替换成空
'**************************************************
Function PE_Replace(ByVal expression, ByVal find, ByVal replacewith)
    If IsNull(expression) Or IsNull(find) Then
        PE_Replace = expression
    ElseIf IsNull(replacewith) Then
        PE_Replace = Replace(expression, find, "")
    Else
        PE_Replace = Replace(expression, find, replacewith)
    End If
End Function

'**************************************************
'函数名：IsExists
'作  用：判断数据库中的数据表的字段是否存在
'参  数：fieldName ---- 字段名称
'        tableName ---- 数据表名称
'返回值：如果改数据表存在改字段,则返回True,否则返回False
'**************************************************
Function IsExists(fieldName, tableName)
    On Error Resume Next
    IsExists = True
    CONN.execute ("select " & fieldName & " from " & tableName)

    If Err Then
        IsExists = False
    End If
    Err.Clear
End Function

'**************************************************
'函数名：Refresh
'作  用：等待特定时间后跳转到指定的网址
'参  数：url ---- 跳转网址
'        refreshTime ---- 等待跳转时间
'**************************************************
Sub Refresh(url,refreshTime)
        Response.Write "<a Name='rsfreshurl' ID='rsfreshurl' href='"& url &"'></a>" & vbCrLf
        Response.Write "<script language=""javascript""> " & vbCrLf
        Response.Write "  function nextpage(){" & vbCrLf
        Response.Write "    var url = document.getElementById('rsfreshurl');" & vbCrLf
        Response.Write "    if (document.all) {" & vbCrLf
        Response.Write "      url.click();" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "   else if (document.createEvent) {" & vbCrLf
        Response.Write "     var ev = document.createEvent('HTMLEvents');" & vbCrLf
        Response.Write "       ev.initEvent('click', false, true);" & vbCrLf
        Response.Write "       url.dispatchEvent(ev);" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  setTimeout(""nextpage();"","&refreshTime*1000&");" & vbCrLf
        Response.Write "</script>" & vbCrLf
End Sub

%>
