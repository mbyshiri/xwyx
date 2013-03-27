<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'�жϵ�ǰ�������Ƿ��Ѿ���¼�����ѵ�¼�����ȡ���ݲ�����Ҫ��ֵ
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

'���û�����Ӧ������ֵ
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
            IsOffer = "��"
        Else
            IsOffer = "��"
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
'��������GetSubStr
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'        strlen ----��ȡ����
'        bShowPoint ---- �Ƿ���ʾʡ�Ժ�
'����ֵ����ȡ����ַ���
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
        strTemp = strTemp & "��"
    End If
    GetSubStr = strTemp
End Function

'**************************************************
'��������GetStrLen
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
Function GetStrLen(str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("�й�") = 2)
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
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
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
'��������ShowPage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sFileName  ----���ӵ�ַ
'        TotalNumber ----������
'        MaxPerPage  ----ÿҳ����
'        CurrentPage ----��ǰҳ
'        ShowTotal   ----�Ƿ���ʾ������
'        ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת��
'        strUnit     ----������λ
'        ShowMaxPerPage  ----�Ƿ���ʾÿҳ��Ϣ��ѡ���
'����ֵ������һҳ ��һҳ������Ϣ��HTML����
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
        strTemp = strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;&nbsp;"
    End If
    
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "��ҳ | ��һҳ |"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a> |"
        strTemp = strTemp & "  <a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>��һҳ</a> | "
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
        strTemp = strTemp & "| ��һҳ | βҳ"
    Else
        strTemp = strTemp & " | <a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>��һҳ</a> |"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>  βҳ</a>"
    End If
	If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;&nbsp;&nbsp;<Input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"">" & strUnit & "/ҳ"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/ҳ"
    End If
    If ShowAllPages = True Then
            strTemp = strTemp & "&nbsp;&nbsp;ת����<Input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;"">ҳ"
    End If
    strTemp = strTemp & "</div>"
    ShowPage = strTemp
End Function


'**************************************************
'��������ShowPage_en
'��  �ã���ʾӢ�ġ���һҳ ��һҳ������Ϣ
'��  ����sFileName  ----���ӵ�ַ
'        TotalNumber ----������
'        MaxPerPage  ----ÿҳ����
'        CurrentPage ----��ǰҳ
'        ShowTotal   ----�Ƿ���ʾ������
'        ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת��
'        strUnit     ----������λ
'        ShowMaxPerPage  ----�Ƿ���ʾÿҳ��Ϣ��ѡ���
'����ֵ������һҳ ��һҳ������Ϣ��HTML����
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
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'        False ----û�а�װ
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
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
Sub WriteErrMsg(sErrMsg, sComeUrl)
    Response.Write "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='" & strInstallDir & "images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & sErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        Response.Write "<a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a>"
    Else
        Response.Write "<a href='javascript:window.close();'>���رա�</a>"
    End If
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub

'**************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
Sub WriteSuccessMsg(sSuccessMsg, sComeUrl)
    Response.Write "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='" & strInstallDir & "images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & sSuccessMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        Response.Write "<a href='" & sComeUrl & "'>&lt;&lt; ������һҳ</a>"
    Else
        Response.Write "<a href='javascript:window.close();'>���رա�</a>"
    End If
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub

'**************************************************
'��������FoundInArr
'��  �ã�����������Ƿ���ָ������ֵ
'��  ����strArr ----- ���������
'        strItem  ----- �����ַ�
'        strSplit  ----- �ָ��ַ�
'����ֵ��True  ----��
'        False ----û��
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
'��������GetRndPassword
'��  �ã��õ�ָ��λ�������������
'��  ����PasswordLen ---- λ��
'����ֵ�������ַ���
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
'��������GetRndNum
'��  �ã������ƶ�λ���������
'��  ����iLength ---- �漴����λ��
'����ֵ�������
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
'��������GetIDByDefault
'��  �ã���ȡIDֵ�����IDΪ0����ʹ��ȱʡֵ
'��  ����ItemID ---- ��ĿIDֵ
'        DefaultID ---- ȱʡIDֵ
'**************************************************
Function GetIDByDefault(ItemID, DefaultID)
    Dim iItemID
    iItemID = ItemID
    If iItemID = 0 Then iItemID = DefaultID
    If IsNull(iItemID) Then iItemID = 0
    GetIDByDefault = iItemID
End Function




'**************************************************
'��������FillInArrStr
'��  �ã�ʹ��һ���ö��ŷָ����ַ������������һ�����ŷָ����ַ�����ʹ��ﵽָ������Ŀ��
'��  ����strSource ---- ԭ�ַ���
'        strFill ---- ����ַ���
'        ItemNum ---- ָ���������Ŀ��
'����ֵ��������ַ���
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
'��������XmlText
'��  �ã������԰��ж�ȡָ���ڵ��ֵ
'��  ����iBigNode ---- ��ڵ�
'        iSmallNode ---- С�ڵ�
'        DefChar ---- Ĭ��ֵ
'����ֵ�����԰���ָ���ڵ��ֵ
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
'��������GetFirstSeparatorToEnd
'��  �ã���ȡ�ӵ�һ���ָ�������β���ַ���
'��  ����str   ----ԭ�ַ���
'        separator ----�ָ���
'����ֵ����ȡ����ַ���
'**************************************************
Function GetFirstSeparatorToEnd(ByVal str, separator)
    GetFirstSeparatorToEnd = Right(str, Len(str) - InStr(str, separator))
End Function

'**************************************************
'��������ChkValidDays
'��  �ã���Ч�ڵĺ���
'��  ����iValidNum ----��Ч��
'        iValidUnit ----��Ч�ڵ�λ
'        iBeginTime ---- ��ʼ��������
'����ֵ��ʣ�����Ч����
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
'��������GetNumString
'��  �ã������Ŀ�漴��
'����ֵ��������ظ�������(�����ϴ�,����)
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
'��������GetMinID
'��  �ã�ȡĳһ��ĳһ�ֶ��е����ֵ
'��  ����SheetName ----��ѯ��
'        FieldName ----��ѯ�ֶ�
'����ֵ�����ֶ���Сֵ
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
'��������GetNewID
'��  �ã�ȡĳһ��ĳһ�ֶ��е����ֵ+1
'��  ����SheetName ----��ѯ��
'        FieldName ----��ѯ�ֶ�
'����ֵ�����ֶ����ֵ+1
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
'��������PE_Replace
'��  �ã��ݴ��滻
'��  ����expression ---- ������
'        find ---- ���滻���ַ�
'        replacewith ---- �滻����ַ�
'����ֵ���ݴ����滻�ַ���,��� replacewith ���ַ�,���滻���ַ� �滻�ɿ�
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
'��������IsExists
'��  �ã��ж����ݿ��е����ݱ���ֶ��Ƿ����
'��  ����fieldName ---- �ֶ�����
'        tableName ---- ���ݱ�����
'����ֵ����������ݱ���ڸ��ֶ�,�򷵻�True,���򷵻�False
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
'��������Refresh
'��  �ã��ȴ��ض�ʱ�����ת��ָ������ַ
'��  ����url ---- ��ת��ַ
'        refreshTime ---- �ȴ���תʱ��
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
