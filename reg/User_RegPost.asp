<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim CheckCode, CheckNum, CheckUrl

Call Main
Call CloseConn

Sub Main()
    UserTrueIP = ReplaceBadChar(UserTrueIP)

    Dim UserPassword, PwdConfirm, Question, Answer
    Dim i
    UserName = Trim(Request("UserName"))
    UserPassword = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    Question = Trim(Request("Question"))
    Answer = Trim(Request("Answer"))
    Email = Trim(Request("Email"))
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))

    Dim strFields, arrFields, arrTemp, NeedAddContacter
    strFields = "Homepage,��ҳ|QQ,QQ����|ICQ,ICQ����|MSN,MSN�ʺ�|Yahoo,�Ż�ͨ�ʺ�|UC,UC����|Aim,Aim�ʺ�|OfficePhone,�칫�绰|HomePhone,��ͥ�绰|Fax,�������|Mobile,�ֻ�����|PHS,С��ͨ����|Region,����/������ʡ��/�ݿ�������|Address,��ϵ��ַ|ZipCode,��������|TrueName,��ʵ����|Birthday,��������|IDCard,���֤����|Vocation,ְҵ|Company,��˾/��λ����|Department,��������|PosTitle,ְ��|Marriage,����״̬|Income,�������|UserFace,�û�ͷ��|FaceWidth,ͷ����|FaceHeight,ͷ��߶�|Sign,ǩ����|Privacy,��˽�趨"
    arrFields = Split(strFields, "|")

    Randomize
    CheckNum = CStr(Int(7999 * Rnd + 2000)) & CStr(Int(7999 * Rnd + 2000)) '�����֤��
    CheckUrl = Request.ServerVariables("HTTP_REFERER")
    CheckUrl = Left(CheckUrl, InStrRev(CheckUrl, "/")) & "User_RegCheck.asp?Action=Check&UserName=" & UserName & "&Password=" & UserPassword & "&CheckNum=" & CheckNum
    If UserName = "" Or GetStrLen(UserName) > UserNameMax Or GetStrLen(UserName) < UserNameLimit Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������û���(���ܴ���" & UserNameMax & "С��" & UserNameLimit & ")</li>"
    Else
        If CheckUserBadChar(UserName) = False Then
            ErrMsg = ErrMsg & "<li>�û����к��зǷ��ַ�</li>"
            FoundErr = True
        End If
    End If


    If FoundInArr(UserName_RegDisabled, UserName, "|") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������û���Ϊϵͳ��ֹע����û���</li>"
    End If
    If UserPassword = "" Or GetStrLen(UserPassword) > 12 Or GetStrLen(UserPassword) < 6 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����������(���ܴ���12С��6)</li>"
    Else
        If CheckBadChar(UserPassword) = False Then
            ErrMsg = ErrMsg + "<li>�����к��зǷ��ַ�</li>"
            FoundErr = True
        End If
    End If
    If PwdConfirm = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ȷ������(���ܴ���12С��6)</li>"
    Else
        If UserPassword <> PwdConfirm Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����ȷ�����벻һ��</li>"
        End If
    End If
    If Question = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʾ���ⲻ��Ϊ��</li>"
    End If
    If Answer = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����𰸲���Ϊ��</li>"
    End If
    If Email = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Email����Ϊ��</li>"
    Else
        If IsValidEmail(Email) = False Then
            ErrMsg = ErrMsg & "<li>����Email�д���</li>"
            FoundErr = True
        End If
    End If
    If EnableCheckCodeOfReg = True Then
        If Trim(Session("CheckCode")) = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ע��ҳ���ͣ��ʱ�����������ע����֤��ʧЧ�������·���ע��ҳ�����ע�ᡣ</li>"
        End If
        If CheckCode <> Session("CheckCode") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�������ע����֤���ϵͳ�����Ĳ�һ�£����������롣</li>"
        End If
    End If
    If EnableQAofReg = True Then
        Dim arrQAofReg
        arrQAofReg = Split(QAofReg & "", "$$$")
        For i = 0 To 2
            If Trim(arrQAofReg(i * 2)) <> "" Then
                If Trim(Request("RegAnswer" & i)) <> Trim(arrQAofReg(i * 2 + 1)) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����ȷ�ش�ע����֤���⣬������������ע��</li>"
                    Exit For
                End If
            End If
        Next
    End If
    For i = 0 To UBound(arrFields)
        arrTemp = Split(arrFields(i), ",")
        If FoundInArr(RegFields_MustFill, arrTemp(0), ",") Then
            NeedAddContacter = True
            If Trim(Request(arrTemp(0))) = "" Or (i = 1 And LCase(Trim(Request(arrTemp(0)))) = "http://") Then
                If arrTemp(0) <> "Region" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����д��" & arrTemp(1) & "</li>"
                ElseIf Trim(Request("Country")) = "" Or Trim(Request("Province")) = "" Or Trim(Request("City")) = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����д��" & arrTemp(1) & "</li>"
                End If
            End If
        End If
    Next

    If FoundErr = True Then
        Call ShowRegResult
        Exit Sub
    End If


    Dim sqlReg, rsReg, trs, RndPassword
    Set trs = Conn.Execute("select max(UserID) from PE_User")
    If IsNull(trs(0)) Then
        UserID = 1
    Else
        UserID = trs(0) + 1
    End If
    Set trs = Nothing
    sqlReg = "select * from PE_User where UserName='" & UserName & "'"
    Set rsReg = Server.CreateObject("adodb.recordset")
    rsReg.Open sqlReg, Conn, 1, 3
    If Not (rsReg.BOF And rsReg.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
    Else
        If Not EnableMultiRegPerEmail Then
            Dim rsEmailOnce
            Set rsEmailOnce = Conn.Execute("select UserID from PE_User where Email='" & Email & "'")
            If Not (rsEmailOnce.BOF And rsEmailOnce.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��ע���Email�Ѿ����ڣ������Email�����ԣ�</li>"
            End If
            Set rsEmailOnce = Nothing
        End If
    End If
    If FoundErr = True Then
        rsReg.Close
        Set rsReg = Nothing
        Call ShowRegResult
        Exit Sub
    End If

    '��Ӷ����Ͻӿڵ�֧��
    If API_Enable Then
        '���ô���״̬����Ϣ
        FoundErr = False
        ErrMsg = ""
        '��Ҫ���͵���Ϣ��������
        sPE_Items(conSyskey, 1) = MD5(UserName & API_Key, 16)
        sPE_Items(conAction, 1) = "reguser"
        sPE_Items(conUsername, 1) = UserName
        sPE_Items(conPassword, 1) = UserPassword
        sPE_Items(conQuestion, 1) = Question
        sPE_Items(conAnswer, 1) = Answer
        sPE_Items(conEmail, 1) = Email
        sPE_Items(conUserstatus, 1) = 0
        sPE_Items(conJointime, 1) = Now()
        sPE_Items(conUserip, 1) = UserTrueIP
        sPE_Items(conTruename, 1) = PE_HTMLEncode(Trim(Request.Form("TrueName")))
        sPE_Items(conGender, 1) = exchangeGender(Trim(Request.Form("Sex")))
        sPE_Items(conBirthday, 1) = FormatDateTime(PE_CDate(Trim(Request.Form("Birthday"))), vbShortDate)
        sPE_Items(conQQ, 1) = PE_HTMLEncode(Trim(Request.Form("QQ")))
        sPE_Items(conMsn, 1) = PE_HTMLEncode(Trim(Request.Form("MSN")))
        sPE_Items(conMobile, 1) = PE_HTMLEncode(Trim(Request.Form("Mobile")))
        sPE_Items(conTelephone, 1) = PE_HTMLEncode(Trim(Request.Form("OfficePhone")))
        sPE_Items(conProvince, 1) = PE_HTMLEncode(Trim(Request.Form("Province")))
        sPE_Items(conCity, 1) = PE_HTMLEncode(Trim(Request.Form("City")))
        sPE_Items(conAddress, 1) = PE_HTMLEncode(Trim(Request.Form("Address")))
        sPE_Items(conZipcode, 1) = PE_HTMLEncode(Trim(Request.Form("ZipCode")))
        sPE_Items(conHomepage, 1) = PE_HTMLEncode(Trim(Request.Form("HomePage")))
        If createXmlDom Then
            '֧��MSXML�������ݴ���xml��
            prepareXML True
            '�����Ͻӿڷ���ע������
            SendPost
            If FoundErr Then
                ErrMsg = "<li>" & ErrMsg & "</li>"
            End If
        Else
            '��������֧��MSXML
            FoundErr = True
            ErrMsg = "<li>Ŀǰע����񲻿���! [APIError-XmlDom-Runtime]</li>"
        End If
    End If
    '���

    If FoundErr = True Then
        Call ShowRegResult
        Exit Sub
    End If

    RndPassword = GetRndPassword(16)

    rsReg.addnew
    rsReg("UserID") = UserID
    rsReg("ClientID") = 0
    rsReg("ContacterID") = 0
    rsReg("CompanyID") = 0
    rsReg("UserType") = 0
    rsReg("UserName") = UserName
    rsReg("UserPassword") = MD5(UserPassword, 16)
    rsReg("LastPassword") = RndPassword
    rsReg("Question") = Question
    rsReg("Answer") = MD5(Answer, 16)
    rsReg("Email") = Email
    rsReg("RegTime") = Now()
    rsReg("IsLocked") = False
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
    rsReg("UserFriendGroup") = "������$�ҵĺ���"
    rsReg("LoginTimes") = 1
    rsReg("LastLoginIP") = UserTrueIP
    rsReg("LastLoginTime") = Now()
    rsReg("LastPresentTime") = Now()
    rsReg("UserFace") = PE_HTMLEncode(Trim(Request.Form("UserFace")))
    rsReg("FaceWidth") = PE_CLng(Trim(Request.Form("FaceWidth")))
    rsReg("FaceHeight") = PE_CLng(Trim(Request.Form("FaceHeight")))
    rsReg("Sign") = PE_HTMLEncode(Trim(Request.Form("Sign")))
    rsReg("Privacy") = PE_CLng(Trim(Request.Form("Privacy")))
    If EmailCheckReg = True Then
        Dim strMailBody
        strMailBody = Replace(EmailOfRegCheck, "{$CheckNum}", CheckNum)
        strMailBody = Replace(strMailBody, "{$CheckUrl}", CheckUrl)

        Dim PE_Mail
        Set PE_Mail = New SendMail
        ErrMsg = PE_Mail.Send(Email, UserName, "ע��ȷ����", strMailBody, SiteName, WebmasterEmail, 3)
        Set PE_Mail = Nothing
        Set trs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=0")
    Else
        If AdminCheckReg = True Then
            Set trs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=1")
        Else
            Set trs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=2")
        End If
    End If
    Dim GroupID, GroupSetting
    GroupID = trs(0)
    GroupSetting = Split(trs(1), ",")
    Set trs = Nothing
    rsReg("GroupID") = GroupID
    'rsReg("ChargeType") = GroupSetting(14)
    rsReg("UserPoint") = PresentPoint
    rsReg("BeginTime") = FormatDateTime(Now(), 2)
    If PresentValidNum>32767 then
        rsReg("ValidNum") = -1
    Else 
        rsReg("ValidNum") = PresentValidNum
    End If
    rsReg("ValidUnit") = PresentValidUnit
    rsReg("CheckNum") = CheckNum
    rsReg("SpecialPermission") = False
    rsReg.Update
    rsReg.Close
    Set rsReg = Nothing
    Response.Cookies(Site_Sn)("UserName") = UserName
    Response.Cookies(Site_Sn)("UserPassword") = MD5(UserPassword, 16)
    Response.Cookies(Site_Sn)("LastPassword") = RndPassword
    Session("UserID") = UserID
    If PresentMoney > 0 Then
        Conn.Execute ("insert into PE_BankrollItem (UserName,ClientID,DateAndTime,[Money],MoneyType,CurrencyType,eBankID,Bank,Income_PayOut,OrderFormID,PaymentID,Remark,LogTime,IP,Inputer) values('" & UserName & "',0," & PE_Now & "," & PresentMoney & ",4,1,0,'',1,0,0,'ע�����û��������ʽ�'," & PE_Now & ",'" & UserTrueIP & "','System')")
    End If
    If PresentPoint > 0 Then
        Conn.Execute ("insert into PE_ConsumeLog (UserName,ModuleType,InfoID,Point,Income_Payout,Remark,LogTime,Times,IP,Inputer) values ('" & UserName & "',0,0," & PresentPoint & ",1,'ע���»�Ա������" & PointName & "'," & PE_Now & ",1,'" & UserTrueIP & "','System')")
    End If
    If PresentValidNum > 0 Or PresentValidNum = -1 Then
        Conn.Execute ("insert into PE_RechargeLog (UserName,ValidNum,ValidUnit,Income_Payout,Remark,LogTime,IP,Inputer) values ('" & UserName & "'," & PresentValidNum & "," & PresentValidUnit & ",1,'ע���»�Ա��������Ч��'," & PE_Now & ",'" & UserTrueIP & "','System')")
    End If

    If NeedAddContacter = True or PE_CLng(Trim(Request.Form("Sex")))<>"" Then
        Dim ContacterID, sqlContacter, rsContacter
        Set trs = Conn.Execute("select max(ContacterID) from PE_Contacter")
        If IsNull(trs(0)) Then
            ContacterID = 1
        Else
            ContacterID = trs(0) + 1
        End If
        Set trs = Nothing

        sqlContacter = "select top 1 * From PE_Contacter"
        Set rsContacter = Server.CreateObject("adodb.recordset")
        rsContacter.Open sqlContacter, Conn, 1, 3
        rsContacter.addnew
        rsContacter("ContacterID") = ContacterID
        rsContacter("ClientID") = 0
        rsContacter("ParentID") = 0
        rsContacter("UserType") = 0
        rsContacter("TrueName") = PE_HTMLEncode(Trim(Request.Form("TrueName")))
        rsContacter("Title") = PE_HTMLEncode(Trim(Request.Form("Title")))
        rsContacter("Country") = PE_HTMLEncode(Trim(Request.Form("Country")))
        rsContacter("Province") = PE_HTMLEncode(Trim(Request.Form("Province")))
        rsContacter("City") = PE_HTMLEncode(Trim(Request.Form("City")))
        rsContacter("ZipCode") = PE_HTMLEncode(Trim(Request.Form("ZipCode")))
        rsContacter("Address") = PE_HTMLEncode(Trim(Request.Form("Address")))
        rsContacter("Mobile") = PE_HTMLEncode(Trim(Request.Form("Mobile")))
        rsContacter("OfficePhone") = PE_HTMLEncode(Trim(Request.Form("OfficePhone")))
        rsContacter("HomePhone") = PE_HTMLEncode(Trim(Request.Form("HomePhone")))
        rsContacter("PHS") = PE_HTMLEncode(Trim(Request.Form("PHS")))
        rsContacter("Fax") = PE_HTMLEncode(Trim(Request.Form("Fax")))
        rsContacter("Homepage") = PE_HTMLEncode(Trim(Request.Form("Homepage")))
        rsContacter("Email") = Email
        rsContacter("QQ") = PE_HTMLEncode(Trim(Request.Form("QQ")))
        rsContacter("MSN") = PE_HTMLEncode(Trim(Request.Form("MSN")))
        rsContacter("ICQ") = PE_HTMLEncode(Trim(Request.Form("ICQ")))
        rsContacter("Yahoo") = PE_HTMLEncode(Trim(Request.Form("Yahoo")))
        rsContacter("UC") = PE_HTMLEncode(Trim(Request.Form("UC")))
        rsContacter("Aim") = PE_HTMLEncode(Trim(Request.Form("Aim")))
        rsContacter("Company") = PE_HTMLEncode(Trim(Request.Form("Company")))
        rsContacter("Department") = PE_HTMLEncode(Trim(Request.Form("Department")))
        rsContacter("Position") = PE_HTMLEncode(Trim(Request.Form("PosTitle")))
        rsContacter("Operation") = PE_HTMLEncode(Trim(Request.Form("Operation")))
        rsContacter("CompanyAddress") = PE_HTMLEncode(Trim(Request.Form("CompanyAddress")))
        rsContacter("BirthDay") = PE_CDate(Trim(Request.Form("BirthDay")))
        rsContacter("IDCard") = Left(PE_HTMLEncode(Trim(Request.Form("IDCard"))), 20)
        rsContacter("NativePlace") = PE_HTMLEncode(Trim(Request.Form("NativePlace")))
        rsContacter("Nation") = PE_HTMLEncode(Trim(Request.Form("Nation")))
        rsContacter("Sex") = PE_CLng(Trim(Request.Form("Sex")))
        rsContacter("Marriage") = PE_CLng(Trim(Request.Form("Marriage")))
        rsContacter("Education") = PE_CLng(Trim(Request.Form("Education")))
        rsContacter("GraduateFrom") = PE_HTMLEncode(Trim(Request.Form("GraduateFrom")))
        rsContacter("InterestsOfLife") = PE_HTMLEncode(Trim(Request.Form("InterestsOfLife")))
        rsContacter("InterestsOfCulture") = PE_HTMLEncode(Trim(Request.Form("InterestsOfCulture")))
        rsContacter("InterestsOfAmusement") = PE_HTMLEncode(Trim(Request.Form("InterestsOfAmusement")))
        rsContacter("InterestsOfSport") = PE_HTMLEncode(Trim(Request.Form("InterestsOfSport")))
        rsContacter("InterestsOfOther") = PE_HTMLEncode(Trim(Request.Form("InterestsOfOther")))
        rsContacter("Family") = PE_HTMLEncode(Trim(Request.Form("Family")))
        rsContacter("Income") = PE_CLng(Trim(Request.Form("Income")))
        rsContacter("CreateTime") = Now()
        rsContacter("Owner") = ""
        rsContacter("UpdateTime") = Now()
        rsContacter.Update
        rsContacter.Close
        Set rsContacter = Nothing

        Conn.Execute ("update PE_User set ContacterID=" & ContacterID & " where UserID=" & UserID & "")
    End If

    Call ShowRegResult
End Sub

Sub ShowRegResult()

    strHtml = GetTemplate(0, 21, 0)
    Call ReplaceCommonLabel

    Dim strPath
    strPath = "�����ڵ�λ�ã�&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;ע����"

    strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> ע����")
    strHtml = Replace(strHtml, "{$ShowPath}", strPath)


    strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
    strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))

    strHtml = Replace(strHtml, "{$RegResult}", GetRegResult())

    Response.Write strHtml
End Sub

Function GetRegResult()
    Dim strResult
    If FoundErr = True Then
        strResult = strResult & "<br><br><table align='center' width='300' border='0' cellpadding='2' cellspacing='0'>"
        strResult = strResult & "<tr><td align='center' class='main_title_575'>�������µ�ԭ����ע���û���</td></tr>"
        strResult = strResult & "<tr><td align='left' height='100' class='main_tdbg_575'><br>" & ErrMsg & "<p align='center'>��<a href='javascript:onclick=history.go(-1)'>�� ��</a>��<br></p></td></tr>"
        strResult = strResult & "</table>"
    Else
        strResult = strResult & "<br><br><table align='center' width='300' border='0' cellpadding='2' cellspacing='0'>"
        strResult = strResult & "<tr><td align='center' class='main_title_575'>�ɹ�ע���û���</td></tr>"
        strResult = strResult & "<tr><td align='left' height='100' class='main_tdbg_575'><br>��ע����û�����" & UserName & "<br>"
        If EmailCheckReg = True Then
            strResult = strResult & "ϵͳ�Ѿ�������һ��ȷ���ŵ���ע��ʱ��д�������У����������յ�ȷ���Ų�ͨ��ȷ���������ӽ���ȷ�Ϻ���������ʽ��Ϊ��վ��ע���û���"
        Else
            If EnableWap = True And ShowWapShop = True Then
                strResult = strResult & "�����ֻ������룺" & CheckNum & "<br>"
            End If
            If AdminCheckReg = True Then
                strResult = strResult & "��ȴ�����ͨ������ע����������Ϳ�����ʽ��Ϊ��վ��ע���û��ˡ�"
            Else
                If API_Enable Then
                    Dim iIndex, tempAPIScripts
                    sPE_Items(conSyskey, 1) = MD5(UserName & API_Key, 16)
                    For iIndex = 0 To UBound(arrAPIUrls)
                        Dim arrAPIUrl
                        arrAPIUrl = Split(arrAPIUrls(iIndex), "@@")
                        tempAPIScripts = tempAPIScripts & "<script type=""text/javascript"" language=""JavaScript"" src=""" & arrAPIUrl(1) & "?syskey=" & sPE_Items(conSyskey, 1) & "&username=" & UserName & "&password=" & MD5(sPE_Items(conPassword, 1), 16) & """></script>"
                    Next
                    strResult = strResult & tempAPIScripts
                End If
                strResult = strResult & "��ӭ���ļ��룡����<br><br>"
            End If
        End If
        strResult = strResult & "<p align='center'>��<a href='" & InstallDir & "Index.asp'>������ҳ</a>��<br></p></td></tr>"
        strResult = strResult & "</table>"
    End If
    GetRegResult = strResult

End Function
%>
