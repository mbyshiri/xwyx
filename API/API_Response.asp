<!-- #Include File = "../Start.asp" -->
<!-- #Include File = "../Include/PowerEasy.Md5.asp" -->
<!-- #Include File = "API_Config.asp"-->
<!-- #Include File = "API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
If API_Enable = False Then Response.End

Dim recXml
sPE_Items(conSyskey, 1) = Trim(Request.QueryString(sPE_Items(conSyskey, 0)))
sPE_Items(conUsername, 1) = Trim(Request.QueryString(sPE_Items(conUsername, 0)))
sPE_Items(conPassword, 1) = Trim(Request.QueryString(sPE_Items(conPassword, 0)))
sPE_Items(conSavecookie, 1) = Trim(Request.QueryString(sPE_Items(conSavecookie, 0)))
If sPE_Items(conSyskey, 1) <> "" Then
    If sPE_Items(conUsername, 1) <> "" Then
        If sPE_Items(conPassword, 1) <> "" Then
            WriteCookies
        Else
            CleanCookies
        End If
    End If
Else
    DealResponse
End If

Sub WriteCookies()
    Dim strRndPass
    If Not CheckSysKey(sPE_Items(conUsername, 1), sPE_Items(conSyskey, 1)) Then
        Exit Sub
    End If
    strRndPass = GetRndPassword(16)
    If sPE_Items(conSavecookie, 1) <> "" Then
        sPE_Items(conSavecookie, 1) = PE_CLng(sPE_Items(conSavecookie, 1))
    End If
    Dim iSubDomainIndex, strSite_Sn
    If Enable_SubDomain Then
        For iSubDomainIndex = 0 To UBound(arrSubDomains)
            strSite_Sn = LCase(arrSubDomains(iSubDomainIndex) & Replace(Replace(DomainRoot & InstallDir, "/", ""), ".", ""))
            Response.Cookies(strSite_Sn).Domain = DomainRoot
            Select Case sPE_Items(conSavecookie, 1)
            Case 0
                'not save
            Case 1
                Response.Cookies(strSite_Sn).Expires = Date + 1
            Case 2
                Response.Cookies(strSite_Sn).Expires = Date + 31
            Case 3
                Response.Cookies(strSite_Sn).Expires = Date + 365
            End Select
            Response.Cookies(strSite_Sn)("UserName") = sPE_Items(conUsername, 1)
            Response.Cookies(strSite_Sn)("UserPassword") = sPE_Items(conPassword, 1)
            Response.Cookies(strSite_Sn)("LastPassword") = strRndPass
            Response.Cookies(strSite_Sn)("CookieDate") = sPE_Items(conSavecookie, 1)
        Next
    Else
        Select Case sPE_Items(conSavecookie, 1)
            Case 0
                'not save
            Case 1
                Response.Cookies(Site_Sn).Expires = Date + 1
            Case 2
                Response.Cookies(Site_Sn).Expires = Date + 31
            Case 3
                Response.Cookies(Site_Sn).Expires = Date + 365
        End Select
        Response.Cookies(Site_Sn)("UserName") = sPE_Items(conUsername, 1)
        Response.Cookies(Site_Sn)("UserPassword") = sPE_Items(conPassword, 1)
        Response.Cookies(Site_Sn)("LastPassword") = strRndPass
        Response.Cookies(Site_Sn)("CookieDate") = sPE_Items(conSavecookie, 1)
    End If
    sPE_Items(conUsername, 1) = ReplaceBadChar(sPE_Items(conUsername, 1))
    UserTrueIP = ReplaceBadChar(UserTrueIP)
    Conn.Execute ("UPDATE PE_User Set LastPassword='" & strRndPass & "',LastLoginIP='" & UserTrueIP & "',LastLoginTime=" & PE_Now & ",LoginTimes=LoginTimes+1 WHERE UserName='" & sPE_Items(conUsername, 1) & "'")
    Session("UserID") = UserID
End Sub

Sub CleanCookies()
    If Not CheckSysKey(sPE_Items(conUsername, 1), sPE_Items(conSyskey, 1)) Then
        Exit Sub
    End If
    Dim iSubDomainIndex, strSite_Sn, iItem
    If Enable_SubDomain Then
        For iSubDomainIndex = 0 To UBound(arrSubDomains)
            strSite_Sn = LCase(arrSubDomains(iSubDomainIndex) & Replace(Replace(DomainRoot & InstallDir, "/", ""), ".", ""))
            Response.Cookies(strSite_Sn).Domain = DomainRoot
            Response.Cookies(strSite_Sn)("UserName") = ""
            Response.Cookies(strSite_Sn)("UserPassword") = ""
            Response.Cookies(strSite_Sn)("LastPassword") = ""
        Next
    Else
        Response.Cookies(Site_Sn)("UserName") = ""
        Response.Cookies(Site_Sn)("UserPassword") = ""
        Response.Cookies(Site_Sn)("LastPassword") = ""
    End If

End Sub

Sub DealResponse()
    'On Error Resume Next
    If createXmlDom Then
        sMyXmlDoc.Load Request
        If sMyXmlDoc.parseError.errorCode <> 0 Then
            FoundErr = True
            ErrMsg = sMyXmlDoc.parseError.reason & "001"
        Else
            sPE_Items(conSyskey, 1) = getNodeText(sPE_Items(conSyskey, 0))
            sPE_Items(conUsername, 1) = getNodeText(sPE_Items(conUsername, 0))
            sPE_Items(conAction, 1) = getNodeText(sPE_Items(conAction, 0))
            
            If sPE_Items(conSyskey, 1) = "" Or sPE_Items(conUsername, 1) = "" Or sPE_Items(conAction, 1) = "" Then
                FoundErr = True
                ErrMsg = "未包含必须元素，数据同步被拒绝!"
            End If
            If Not CheckSysKey(sPE_Items(conUsername, 1), sPE_Items(conSyskey, 1)) Then
                FoundErr = True
                ErrMsg = "安全码不符，数据同步被拒绝!"
            End If
        End If
    Else
        FoundErr = True
        ErrMsg = "服务器不支持MSXML对象。"
    End If
    If Err Then
        FoundErr = True
        ErrMsg = Err.Description
        Err.Clear
        WriteErrXml
        Exit Sub
    End If
    If FoundErr Then
        sPE_Items(conStatus, 1) = "1"
        sPE_Items(conMessage, 1) = ErrMsg
        prepareXML False
        WriteXml
        Exit Sub
    End If
    '已处理的元素：syskey,username
    '错误检测完成，开始处理数据
    sPE_Items(conAction, 1) = getNodeText(sPE_Items(conAction, 0))
    '已处理的元素：syskey,username,action
    Select Case sPE_Items(conAction, 1)
        Case "checkname"
            Call checkUser
        Case "reguser"
            Call createUser
        Case "login"
            Call loginUser
        Case "logout"
            Call CleanCookies
        Case "update"
            Call UpdateUser
        Case "delete"
            Call DeleteUser
        Case "getinfo"
            Call GetUserInfo
    End Select
    If FoundErr Then
        sPE_Items(conStatus, 1) = "1"
        sPE_Items(conMessage, 1) = ErrMsg
        prepareXML (False)
        WriteXml
        Exit Sub
    Else
        sPE_Items(conStatus, 1) = "0"
        prepareXML (False)
        WriteXml
    End If
End Sub

Sub checkUser()
    sPE_Items(conEmail, 1) = getNodeText(sPE_Items(conEmail, 0))
    CheckUserName (sPE_Items(conUsername, 1))
    CheckUserEmail (sPE_Items(conEmail, 1))
End Sub

Sub createUser()
    sPE_Items(conEmail, 1) = getNodeText(sPE_Items(conEmail, 0))
    If CheckUserName(sPE_Items(conUsername, 1)) = False Or CheckUserEmail(sPE_Items(conEmail, 1)) = False Then
        Exit Sub
    End If
    
    prepareData True
    
    Dim sqlReg, rsReg, tRs, RndPassword, CheckNum
    Set tRs = Conn.Execute("select max(UserID) from PE_User")
    If IsNull(tRs(0)) Then
        UserID = 1
    Else
        UserID = tRs(0) + 1
    End If
    Set tRs = Nothing

    RndPassword = GetRndPassword(16)
    Set rsReg = Server.CreateObject("adodb.recordset")
    rsReg.OPEN "SELECT * FROM PE_User WHERE UserID=0", Conn, 1, 3
    rsReg.addnew
    rsReg("UserID") = UserID
    rsReg("ClientID") = 0
    rsReg("ContacterID") = 0
    rsReg("UserType") = 0
    rsReg("UserName") = sPE_Items(conUsername, 1)
    rsReg("UserPassword") = Md5(sPE_Items(conPassword, 1), 16)
    rsReg("LastPassword") = RndPassword
    rsReg("Question") = sPE_Items(conQuestion, 1)
    rsReg("Answer") = Md5(sPE_Items(conAnswer, 1), 16)
    rsReg("Email") = sPE_Items(conEmail, 1)
    rsReg("RegTime") = PE_CDate(sPE_Items(conJointime, 1))
    rsReg("LoginTimes") = 0
    If sPE_Items(conUserstatus, 1) = "1" Then
        rsReg("IsLocked") = True
    Else
        rsReg("IsLocked") = False
    End If
    rsReg("Balance") = PresentMoney
    rsReg("UserExp") = PresentExp
    rsReg("PostItems") = 0
    rsReg("PassedItems") = 0
    rsReg("DelItems") = 0
    rsReg("UnsignedItems") = ""
    rsReg("UnreadMsg") = 0
    rsReg("arrClass_Browse") = ""
    rsReg("arrClass_View") = ""
    rsReg("arrClass_Input") = ""
    rsReg("UserSetting") = ""
    rsReg("UserFriendGroup") = "黑名单$我的好友"
    rsReg("LoginTimes") = 0
    rsReg("LastLoginIP") = ""
    rsReg("LastLoginTime") = Now()
    rsReg("LastPresentTime") = Now()
    If sPE_Items(conUserstatus, 1) = "4" Then
        Set tRs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=1")
    Else
        Set tRs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=2")
    End If
    Dim GroupID, GroupSetting
    GroupID = tRs(0)
    GroupSetting = Split(tRs(1), ",")
    Set tRs = Nothing
    rsReg("GroupID") = GroupID
    rsReg("UserPoint") = PresentPoint
    rsReg("BeginTime") = FormatDateTime(Now(), 2)
    rsReg("ValidNum") = PresentValidNum
    rsReg("ValidUnit") = PresentValidUnit
    Randomize
    CheckNum = CStr(Int(7999 * Rnd + 2000)) & CStr(Int(7999 * Rnd + 2000))
    rsReg("CheckNum") = CheckNum
    rsReg("SpecialPermission") = False
    rsReg.Update
    rsReg.Close
    Set rsReg = Nothing

    If PresentMoney > 0 Then
        Conn.Execute ("insert into PE_BankrollItem (UserName,ClientID,DateAndTime,[Money],MoneyType,CurrencyType,eBankID,Bank,Income_PayOut,OrderFormID,PaymentID,Remark,LogTime,IP,Inputer) values('" & UserName & "',0," & PE_Now & "," & PresentMoney & ",4,1,0,'',1,0,0,'注册新用户，赠送资金'," & PE_Now & ",'" & UserTrueIP & "','System')")
    End If
    If PresentPoint > 0 Then
        Conn.Execute ("insert into PE_ConsumeLog (UserName,ModuleType,InfoID,Point,Income_Payout,Remark,LogTime,Times,IP,Inputer) values ('" & UserName & "',0,0," & PresentPoint & ",1,'注册新会员，赠送" & PointName & "'," & PE_Now & ",1,'" & UserTrueIP & "','System')")
    End If
    If PresentValidNum > 0 Or PresentValidNum = -1 Then
        Conn.Execute ("insert into PE_RechargeLog (UserName,ValidNum,ValidUnit,Income_Payout,Remark,LogTime,IP,Inputer) values ('" & UserName & "'," & PresentValidNum & "," & PresentValidUnit & ",1,'注册新会员，赠送有效期'," & PE_Now & ",'" & UserTrueIP & "','System')")
    End If
    
    Dim intIndex, NeedContacter
    NeedContacter = False
    For intIndex = 11 To 20
        If sPE_Items(intIndex, 1) <> "" Then
            NeedContacter = True
            Exit For
        End If
    Next
    
    If NeedContacter Then
        Dim ContacterID, sqlContacter, rsContacter
        Set tRs = Conn.Execute("select max(ContacterID) from PE_Contacter")
        If IsNull(tRs(0)) Then
            ContacterID = 1
        Else
            ContacterID = tRs(0) + 1
        End If
        Set tRs = Nothing

        sqlContacter = "select top 1 * From PE_Contacter"
        Set rsContacter = Server.CreateObject("adodb.recordset")
        rsContacter.OPEN sqlContacter, Conn, 1, 3
        rsContacter.addnew
        rsContacter("ContacterID") = ContacterID
        rsContacter("ClientID") = 0
        rsContacter("ParentID") = 0
        rsContacter("UserType") = 0
        rsContacter("TrueName") = sPE_Items(conTruename, 1)
        rsContacter("Title") = ""
        rsContacter("Country") = ""
        rsContacter("Province") = ""
        rsContacter("City") = ""
        rsContacter("ZipCode") = sPE_Items(conZipcode, 1)
        rsContacter("Address") = sPE_Items(conAddress, 1)
        rsContacter("Mobile") = sPE_Items(conMobile, 1)
        rsContacter("OfficePhone") = sPE_Items(conTelephone, 1)
        rsContacter("HomePhone") = ""
        rsContacter("PHS") = ""
        rsContacter("Fax") = ""
        rsContacter("Homepage") = sPE_Items(conHomepage, 1)
        rsContacter("Email") = sPE_Items(conEmail, 1)
        rsContacter("QQ") = sPE_Items(conQQ, 1)
        rsContacter("MSN") = sPE_Items(conMsn, 1)
        rsContacter("ICQ") = ""
        rsContacter("Yahoo") = ""
        rsContacter("UC") = ""
        rsContacter("Aim") = ""
        rsContacter("Company") = ""
        rsContacter("Department") = ""
        rsContacter("Position") = ""
        rsContacter("Operation") = ""
        rsContacter("CompanyAddress") = ""
        rsContacter("BirthDay") = PE_CDate(sPE_Items(conBirthday, 1))
        rsContacter("IDCard") = ""
        rsContacter("NativePlace") = ""
        rsContacter("Nation") = ""
        If sPE_Items(conSex, 1) = "0" Then
            sPE_Items(conSex, 1) = 1
        ElseIf sPE_Items(conSex, 1) = "1" Then
            sPE_Items(conSex, 1) = 0
        Else
            sPE_Items(conSex, 1) = 2
        End If
        rsContacter("Sex") = sPE_Items(conSex, 1)
        rsContacter("Marriage") = 0
        rsContacter("Education") = 0
        rsContacter("GraduateFrom") = ""
        rsContacter("InterestsOfLife") = ""
        rsContacter("InterestsOfCulture") = ""
        rsContacter("InterestsOfAmusement") = ""
        rsContacter("InterestsOfSport") = ""
        rsContacter("InterestsOfOther") = ""
        rsContacter("Family") = ""
        rsContacter("Income") = 0
        rsContacter("CreateTime") = Now()
        rsContacter("Owner") = ""
        rsContacter("UpdateTime") = Now()
        rsContacter.Update
        rsContacter.Close
        Set rsContacter = Nothing

        Conn.Execute ("update PE_User set ContacterID=" & ContacterID & " where UserID=" & UserID & "")
    End If
End Sub

Sub loginUser()
    Dim strRndPass
    strRndPass = GetRndPassword(16)
    sPE_Items(conPassword, 1) = getNodeText(sPE_Items(conPassword, 0))
    sPE_Items(conPassword, 1) = Md5(sPE_Items(conPassword, 1), 16)
    Dim tRs
    sPE_Items(conUsername, 1) = ReplaceBadChar(sPE_Items(conUsername, 1))
    Set tRs = Conn.Execute("SELECT UserID FROM PE_User WHERE UserName='" & sPE_Items(conUsername, 1) & "' AND UserPassword='" & sPE_Items(conPassword, 1) & "'")
    If tRs.Bof And tRs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "数据库中没有此用户的记录!"
    End If
    tRs.Close
    Set tRs = Nothing
End Sub

Sub UpdateUser()
    Dim tRs, tUserID
    sPE_Items(conUsername, 1) = ReplaceBadChar(sPE_Items(conUsername, 1))
    Set tRs = Conn.Execute("SELECT UserID FROM PE_User WHERE UserName='" & sPE_Items(conUsername, 1) & "'")
    If tRs.EOF And tRs.Bof Then
        FoundErr = True
        ErrMsg = "数据库中没有此用户的记录!"
    Else
        tUserID = tRs(0)
    End If
    tRs.Close
    Set tRs = Nothing
    If FoundErr Then Exit Sub
    
    prepareData True
    
    Dim tSql
    tSql = "SELECT * FROM PE_User WHERE UserName='" & sPE_Items(conUsername, 1) & "'"
    Set tRs = Server.CreateObject("adodb.recordset")
    tRs.OPEN tSql, Conn, 1, 3
    If sPE_Items(conPassword, 1) <> "" Then
        tRs("UserPassword") = Md5(sPE_Items(conPassword, 1), 16)
    End If
    If sPE_Items(conQuestion, 1) <> "" Then
        tRs("Question") = sPE_Items(conQuestion, 1)
    End If
    If sPE_Items(conAnswer, 1) <> "" Then
        tRs("Answer") = Md5(sPE_Items(conAnswer, 1), 16)
    End If
    If sPE_Items(conEmail, 1) <> "" Then
        tRs("Email") = sPE_Items(conEmail, 1)
    End If
    If sPE_Items(conUserstatus, 1) = "" Then
        sPE_Items(conUserstatus, 1) = "0"
    End If
    Select Case sPE_Items(conUserstatus, 1)
    Case "0"
        tRs("Islocked") = False
    Case "4"
        tRs("Islocked") = True
    Case "1"
        tRs("IsLocked") = True
    Case Else
        tRs("IsLocked") = True
    End Select
    tRs.Update
    tRs.Close
    Dim intIndex, NeedContacter
    NeedContacter = False
    For intIndex = 7 To 20
        If intIndex < 8 Or intIndex > 10 Then
            If sPE_Items(intIndex, 1) <> "" Then
                NeedContacter = True
                Exit For
            End If
        End If
    Next
    If NeedContacter Then
        tSql = "SELECT * FROM PE_Contacter WHERE ContacterID=" & tUserID
        tRs.OPEN tSql, Conn, 1, 3
        If Not (tRs.Bof And tRs.EOF) Then
            If sPE_Items(conEmail, 1) <> "" Then
                tRs("Email") = sPE_Items(conEmail, 1)
            End If
            If sPE_Items(conTruename, 1) <> "" Then
                tRs("TrueName") = sPE_Items(conTruename, 1)
            End If
            If sPE_Items(conZipcode, 1) <> "" Then
                tRs("ZipCode") = sPE_Items(conZipcode, 1)
            End If
            If sPE_Items(conAddress, 1) <> "" Then
                tRs("Address") = sPE_Items(conAddress, 1)
            End If
            If sPE_Items(conMobile, 1) <> "" Then
                tRs("Mobile") = sPE_Items(conMobile, 1)
            End If
            If sPE_Items(conTelephone, 1) <> "" Then
                tRs("OfficePhone") = sPE_Items(conTelephone, 1)
            End If
            If sPE_Items(conHomepage, 1) <> "" Then
                tRs("Homepage") = sPE_Items(conHomepage, 1)
            End If
            If sPE_Items(conQQ, 1) <> "" Then
                tRs("QQ") = sPE_Items(conQQ, 1)
            End If
            If sPE_Items(conMsn, 1) <> "" Then
                tRs("MSN") = sPE_Items(conMsn, 1)
            End If
            If sPE_Items(conBirthday, 1) <> "" Then
                tRs("BirthDay") = PE_CDate(sPE_Items(conBirthday, 1))
            End If
            tRs.Update
        End If
        tRs.Close
        Set tRs = Nothing
    End If
    If Err Then
        Err.Clear
    End If
    
End Sub

Sub DeleteUser()
    Dim arrUserNames, iUserIndex
    Dim rsDel
    Dim delName
    arrUserNames = Split(sPE_Items(conUsername, 1), ",")
    For iUserIndex = 0 To UBound(arrUserNames)
        delName = ReplaceBadChar(arrUserNames(iUserIndex))
        Set rsDel = Conn.Execute("SELECT UserID,ContacterID FROM PE_User WHERE UserName='" & delName & "'")
        If Not (rsDel.EOF And rsDel.Bof) Then
            'On Error Resume Next
            Conn.Execute ("DELETE FROM PE_Favorite WHERE UserID=" & rsDel(0))
            Conn.Execute ("DELETE FROM PE_Contacter WHERE ContacterID=" & rsDel(1))
            Conn.Execute ("DELETE FROM PE_User WHERE UserID=" & rsDel(0))
        End If
        rsDel.Close
        Set rsDel = Nothing
    Next
End Sub

Sub GetUserInfo()
    Dim rsInfo, dsUser, iUserID
    sPE_Items(conUsername, 1) = ReplaceBadChar(sPE_Items(conUsername, 1))
    Set rsInfo = Conn.Execute("SELECT ContacterID,UserName,UserPassword,Email,Question,Answer,RegTime,LastLoginIP,Balance,UserExp,UserPoint,ConsumePoint,PostItems,IsLocked " &_
                 "FROM PE_User WHERE UserName='" & sPE_Items(conUsername,1) & "'")
    If rsInfo.EOF And rsInfo.Bof Then
        FoundErr = True
        ErrMsg = "查询的用户不存在"
        iUserID = "0"
    Else
        iUserID = CStr(rsInfo(0))
        sPE_Items(conPassword, 1) = rsInfo("UserPassword")
        sPE_Items(conEmail, 1) = rsInfo("Email")
        sPE_Items(conQuestion, 1) = rsInfo("Question")
        sPE_Items(conAnswer, 1) = rsInfo("Answer")
        sPE_Items(conJointime, 1) = rsInfo("RegTime")
        sPE_Items(conUserip, 1) = rsInfo("LastLoginIP")
        sPE_Items(conBalance, 1) = rsInfo("Balance")
        sPE_Items(conExperience, 1) = rsInfo("UserExp")
        sPE_Items(conValuation, 1) = rsInfo("UserPoint")
        sPE_Items(conTicket, 1) = rsInfo("ConsumePoint")
        sPE_Items(conPosts, 1) = rsInfo("PostItems")
        sPE_Items(conUserstatus, 1) = rsInfo("IsLocked")
    End If
    
    rsInfo.Close

    If FoundErr Then
        Set rsInfo = Nothing
        Exit Sub
    End If

    If IsNull(iUserID) = False And iUserID <> "" Then
        iUserID = PE_CLng(iUserID)
        If iUserID <> 0 Then
            Set rsInfo = Conn.Execute("SELECT TrueName,Sex,Homepage,QQ,MSN,OfficePhone,Mobile,Province,City,Address,ZipCode,Birthday " &_
                            "WHERE ContacterID=" & iUserID)
            If Not (rsInfo.EOF And rsInfo.Bof) Then
                sPE_Items(conTruename, 1) = rsInfo("TrueName")
                sPE_Items(conSex, 1) = exchangeGender(rsInfo("Sex"))
                sPE_Items(conHomepage, 1) = rsInfo("Homepage")
                sPE_Items(conQQ, 1) = rsInfo("QQ")
                sPE_Items(conMsn, 1) = rsInfo("MSN")
                sPE_Items(conTelephone, 1) = rsInfo("OfficePhone")
                sPE_Items(conMobile, 1) = rsInfo("Mobile")
                sPE_Items(conProvince, 1) = rsInfo("Province")
                sPE_Items(conCity, 1) = rsInfo("City")
                sPE_Items(conAddress, 1) = rsInfo("Address")
                sPE_Items(conZipcode, 1) = rsInfo(Birthday)
            End If
        End If
    End If
End Sub

Function CheckSysKey(iName, iSysKey)
    If IsNull(iName) Or iName = "" Or IsNull(iSysKey) Or iSysKey = "" Then
        CheckSysKey = False
        Exit Function
    End If
    If Len(iSysKey) = 32 Then
        iSysKey = Mid(iSysKey, 9, 16)
    End If
    Dim strPEKey, strPEKeyNew
    strPEKey = Md5(iName&API_Key,16)
    strPEKeyNew = Md5(iName&API_Key,16)
    If LCase(iSysKey) = LCase(strPEKey) Or LCase(iSysKey) = LCase(strPEKeyNew) Then
        CheckSysKey = True
    Else
        CheckSysKey = False
    End If
End Function

Function CheckUserName(iName)
    FoundErr = False
    If CheckUserBadChar(UserName) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "用户名中含有非法字符"
    End If
    
    If FoundErr = True Then Exit Function
    If iName = "" Or GetStrLen(iName) > UserNameMax Or GetStrLen(iName) < UserNameLimit Then
        FoundErr = True
        ErrMsg = ErrMsg & "请输入用户名(不能大于" & UserNameMax & "小于" & UserNameLimit & ")"
    End If

    If FoundInArr(UserName_RegDisabled, iName, "|") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "您输入的用户名为系统禁止注册的用户名！"
    End If
    iName = ReplaceBadChar(iName)
    Dim rsCheckReg
    Set rsCheckReg = Conn.Execute("select UserName from PE_User where UserName='" & iName & "'")
    If Not (rsCheckReg.EOF And rsCheckReg.Bof) Then
        FoundErr = True
        ErrMsg = ErrMsg & "“" & iName & "”已经存在！请换一个用户名再试试！"
    End If
    rsCheckReg.Close
    Set rsCheckReg = Nothing
    If FoundErr = True Then
        CheckUserName = False
    Else
        CheckUserName = True
    End If
End Function

Function CheckUserEmail(iEmail)
    Dim rsCheckUser
    If Not EnableMultiRegPerEmail And iEmail <> "" Then
        iEmail = ReplaceBadChar(iEmail)
        Set rsCheckUser = Conn.Execute("SELECT Email FROM PE_User WHERE Email='" & iEmail & "'")
        If Not (rsCheckUser.EOF And rsCheckUser.Bof) Then
            FoundErr = True
            ErrMsg = ErrMsg & "您所填写的Email已经存在！"
            CheckUserEmail = False
        Else
            CheckUserEmail = True
        End If
        rsCheckUser.Close
        Set rsCheckUser = Nothing
    Else
        CheckUserEmail = True
    End If
End Function

%>
