<%@language="vbscript" codepage="936" %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Buffer = True
Dim BeginTime
BeginTime = Timer
%>
<!--#include file="Conn.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="Include/PowerEasy.Common.All.asp"-->
<!--#include file="Include/PowerEasy.Common.Security.asp"-->
<%
Dim UserTrueIP
Dim ScriptName
Dim Site_Sn
Dim InstallDir, strInstallDir

'网站配置相关的变量
Dim SiteName, SiteTitle, SiteUrl, LogoUrl, BannerUrl, WebmasterName, WebmasterEmail, Copyright, Meta_Keywords, Meta_Description
Dim ShowSiteChannel, ShowAdminLogin, EnableSaveRemote, EnableLinkReg, EnableCountFriendSiteHits, EnableSoftKey, IsCustom_Content, objName_FSO, AdminDir, ADDir, AnnounceCookieTime, SiteHitsOfHot
Dim FileExt_SiteIndex, FileExt_SiteSpecial, SiteUrlType, LockIPType, LockIP, AllModules
Dim EnableUserReg, EmailCheckReg, AdminCheckReg, EnableMultiRegPerEmail, EnableCheckCodeOfLogin, EnableCheckCodeOfReg, EnableQAofReg, QAofReg
Dim UserNameLimit, UserNameMax, UserName_RegDisabled, RegFields_MustFill
Dim PresentExp, PresentMoney, PresentPoint, PresentValidNum, PresentValidUnit, MoneyExchangePoint, MoneyExchangeValidDay, UserExpExchangePoint, UserExpExchangeValidDay
Dim PointName, PointUnit, EmailOfRegCheck
Dim MailObject, MailServer, MailServerUserName, MailServerPassWord, MailDomain
Dim PhotoObject, Thumb_DefaultWidth, Thumb_DefaultHeight, Thumb_Arithmetic, PhotoQuality, Thumb_BackgroundColor, Watermark_Type, Watermark_Text, Watermark_Text_FontName, Watermark_Text_FontSize, Watermark_Text_FontColor, Watermark_Text_Bold
Dim Watermark_Images_FileName, Watermark_Images_Transparence, Watermark_Images_BackgroundColor, Watermark_Position_X, Watermark_Position_Y, Watermark_Position
Dim SearchInterval, SearchResultNum, MaxPerPage_SearchResult, SearchContent
Dim EnableGuestBuy, IncludeTax, TaxRate, Prefix_OrderFormNum, Prefix_PaymentNum
Dim MyCountry, MyProvince, MyCity, MyPostcode
Dim EmailOfOrderConfirm, EmailOfSendCard, EmailOfReceiptMoney, EmailOfRefund, EmailOfInvoice, EmailOfDeliver
Dim GuestBook_EnableVisitor, EnableGuestBookCheck, GuestBook_EnableManageRubbish, PresentExpPerLogin, GuestBook_ManageRubbish, GuestBook_ShowIP, GuestBook_IsAssignSort, GuestBook_MaxPerPage
Dim EnableRss, RssCodeType
Dim EnableWap, WapLogo, EnableWapPl, ShowWapShop, ShowWapAppendix, ShowWapManage
Dim EnableSMS, SMSUserName, SMSKey, SendMessageToAdminWhenOrder, SendMessageToMemberWhenPaySuccess, Mobiles, MessageOfOrder, MessageOfOrderConfirm, MessageOfSendCard, MessageOfReceiptMoney, MessageOfRefund, MessageOfInvoice, MessageOfDeliver
Dim MessageOfAddRemit, MessageOfAddIncome, MessageOfAddPayment, MessageOfExchangePoint, MessageOfAddPoint, MessageOfMinusPoint, MessageOfExchangeValid, MessageOfAddValid, MessageOfMinusValid
Dim ObjInstalled_FSO, fso, hf
Dim FileName_SiteIndex
Dim ShowAnonymous

'用户相关的变量
Dim UserLogined, UserID, UserName, GroupID, GroupName, GroupType, Discount_Member, IsOffer, LoginTimes, RegTime, JoinTime, LastLoginTime, LastLoginIP
Dim ClientID, CompanyID, ContacterID, UserType, email
Dim Balance, UserPoint, UserExp, ValidNum, ValidUnit, ValidDays, SpecialPermission, UserSetting, ChargeType, UserChargeType
Dim UnsignedItems, UnreadMsg, NeedlessCheck, EnableModifyDelete, MaxPerDay, PresentExpTimes, MaxSendNum, MaxFavorite, BlogFlag

'用户权限相关的几个变量
Dim arrClass_Browse, arrClass_View, arrClass_Input, arrClass_Check, arrClass_Manage
Dim UserEnableComment, UserCheckComment
Dim ShowUserModel
'分页时所用变量
Dim FileName, strFileName, MaxPerPage, CurrentPage, totalPut

'搜索用变量
Dim SearchType, strField, Keyword

Dim arrSubDomains
Dim Action, FoundErr, ErrMsg, ComeUrl
Dim arrCardUnit, arrUserType
Dim arrFileExt

Dim SkinID, TemplateID

'XML相关的变量
Dim XmlDoc, XMLDOM, Node


'正则表达式相关的变量
Dim regEx, Match, Match2, Matches, Matches2
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.MultiLine = True

ScriptName = Trim(Request.ServerVariables("SCRIPT_NAME"))
UserTrueIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserTrueIP = "" Then UserTrueIP = Request.ServerVariables("REMOTE_ADDR")
UserTrueIP = ReplaceBadChar(UserTrueIP)
If EnableStopInjection = True Then
    If Request.QueryString <> "" Then Call StopInjection(Request.QueryString)
    If Request.Cookies <> "" Then Call StopInjection(Request.Cookies)
    If LCase(Mid(ScriptName, InStrRev(ScriptName, "/") + 1)) <> "upfile.asp" Then
        Call StopInjection2(Request.Form)
    End If
End If
FoundErr = False
ErrMsg = ""


Call OpenConn
Call GetSiteConfig
Call InitVar
Call IsIPlock
FileName_SiteIndex = "Index" & arrFileExt(FileExt_SiteIndex)

Sub InitVar()
    If Request("page") <> "" Then
        CurrentPage = PE_CLng(Request("page"))
    Else
        CurrentPage = 1
    End If
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 20
    SearchType = PE_CLng(Trim(Request("SearchType")))
    strField = ReplaceBadChar(Trim(Request("Field")))
    Keyword = ReplaceBadChar(Trim(Request("keyword")))

    arrSubDomains = Split("|" & strSubDomains, "|")

    ObjInstalled_FSO = IsObjInstalled(objName_FSO)
    If ObjInstalled_FSO = True Then
        Set fso = Server.CreateObject(objName_FSO)
    Else
        Response.Write "<li>FSO组件不可用，各种与FSO相关的功能都将出错！请运行Install.asp或者到后台网站配置处设置好FSO组件名称。</li>"
    End If
        
    ComeUrl = FilterJs(Trim(Request("ComeUrl")))
    If ComeUrl = "" Then
        ComeUrl = FilterJs(Trim(Request.ServerVariables("HTTP_REFERER")))
    End If
    Action = ReplaceBadChar(Trim(Request("Action")))
    FoundErr = False
    ErrMsg = ""

    Site_Sn = Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & InstallDir), "/", ""), ".", "")
    
    arrCardUnit = Array("点", "天", "月", "年", "元", "卡")
    arrUserType = Array("个人会员", "企业会员（创建者）", "企业会员（管理员）", "企业会员（普通成员）", "企业会员（待审核成员）")
    arrFileExt = Array(".html", ".htm", ".shtml", ".shtm", ".asp")

    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    XmlDoc.async = False

End Sub

Sub StopInjection(Values)
    Dim FoundInjection
    regEx.Pattern = "'|;|#|([\s\b+()]+(select|update|insert|delete|declare|@|exec|dbcc|alter|drop|create|backup|if|else|end|and|or|add|set|open|close|use|begin|retun|as|go|exists)[\s\b+]*)"
    Dim sItem, sValue
    For Each sItem In Values
        sValue = Values(sItem)
        If regEx.Test(sValue) Then
            FoundInjection = True
            Response.Write "很抱歉，由于您提交的内容中含有危险的SQL注入代码，致使本次操作无效！ "
            Response.Write "<br>字段名：" & sItem
            Response.Write "<br>字段值：" & sValue
            Response.Write "<br>关键字："
            Set Matches = regEx.Execute(sValue)
            For Each Match In Matches
                Response.Write FilterJS(Match.value)
            Next
            Response.Write "<br><br>如果您是正常提交仍出现上面的提示，请联系站长修改Config.asp文件的第7行，暂时禁用掉防SQL注入功能，操作完成后再打开。"
            
        End If
    Next
    If FoundInjection = True Then
        Response.End
    End If
End Sub

Sub StopInjection2(Values)
    Dim FoundInjection
    regEx.Pattern = "[';#()][\s+()]*(select|update|insert|delete|declare|@|exec|dbcc|alter|drop|create|backup|if|else|end|and|or|add|set|open|close|use|begin|retun|as|go|exists)[\s+]*"
    Dim sItem, sValue
    For Each sItem In Values
        sValue = Values(sItem)
        If regEx.Test(sValue) Then
            FoundInjection = True
            Response.Write "很抱歉，由于您提交的内容中含有危险的SQL注入代码，致使本次操作无效！ "
            Response.Write "<br>字段名：" & sItem
            Response.Write "<br>字段值：" & sValue
            Response.Write "<br>关键字："
            Set Matches = regEx.Execute(sValue)
            For Each Match In Matches
                Response.Write FilterJS(Match.value)
            Next
            Response.Write "<br><br>如果您是正常提交仍出现上面的提示，请联系站长修改Config.asp文件的第7行，暂时禁用掉防SQL注入功能，操作完成后再打开。"
            
        End If
    Next
    If FoundInjection = True Then
        Response.End
    End If
End Sub
    


Sub GetSiteConfig()
    On Error Resume Next
    Dim rsConfig
    Set rsConfig = Conn.Execute("select * from PE_Config")
    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.Close
        Set rsConfig = Nothing
        Response.Write "网站配置数据丢失！系统无法正常运行！"
        Response.End
        Exit Sub
    End If
    SiteName = rsConfig("SiteName")
    SiteTitle = rsConfig("SiteTitle")
    SiteUrl = rsConfig("SiteUrl")
    InstallDir = rsConfig("InstallDir")
    LogoUrl = rsConfig("LogoUrl")
    BannerUrl = rsConfig("BannerUrl")
    WebmasterName = rsConfig("WebmasterName")
    WebmasterEmail = rsConfig("WebmasterEmail")
    Copyright = rsConfig("Copyright")
    Meta_Keywords = rsConfig("Meta_Keywords")
    Meta_Description = rsConfig("Meta_Description")

    ShowSiteChannel = rsConfig("ShowSiteChannel")
    ShowAdminLogin = rsConfig("ShowAdminLogin")
    EnableSaveRemote = rsConfig("EnableSaveRemote")
    EnableLinkReg = rsConfig("EnableLinkReg")
    EnableCountFriendSiteHits = rsConfig("EnableCountFriendSiteHits")
    EnableSoftKey = rsConfig("EnableSoftKey")
    IsCustom_Content = PE_CBool(rsConfig("IsCustom_Content"))
    objName_FSO = rsConfig("objName_FSO")
    AdminDir = rsConfig("AdminDir")
    ADDir = rsConfig("ADDir")
    AnnounceCookieTime = PE_CLng(rsConfig("AnnounceCookieTime"))
    SiteHitsOfHot = rsConfig("HitsOfHot")
    AllModules = rsConfig("Modules")
    FileExt_SiteIndex = rsConfig("FileExt_SiteIndex")
    FileExt_SiteSpecial = rsConfig("FileExt_SiteSpecial")
    SiteUrlType = rsConfig("SiteUrlType")
    LockIPType = rsConfig("LockIPType")
    LockIP = rsConfig("LockIP")
    ShowUserModel = rsConfig("ShowUserModel")
    ShowAnonymous = rsConfig("ShowAnonymous")	
    EnableUserReg = rsConfig("EnableUserReg")
    EmailCheckReg = rsConfig("EmailCheckReg")
    AdminCheckReg = rsConfig("AdminCheckReg")
    EnableMultiRegPerEmail = rsConfig("EnableMultiRegPerEmail")
    EnableCheckCodeOfLogin = rsConfig("EnableCheckCodeOfLogin")
    EnableCheckCodeOfReg = rsConfig("EnableCheckCodeOfReg")
    EnableQAofReg = rsConfig("EnableQAofReg")
    QAofReg = rsConfig("QAofReg")

    UserNameLimit = rsConfig("UserNameLimit")
    UserNameMax = rsConfig("UserNameMax")
    UserName_RegDisabled = rsConfig("UserName_RegDisabled")
    RegFields_MustFill = rsConfig("RegFields_MustFill")
    
    PresentExp = rsConfig("PresentExp")
    PresentMoney = rsConfig("PresentMoney")
    PresentPoint = rsConfig("PresentPoint")
    PresentValidNum = rsConfig("PresentValidNum")
    PresentValidUnit = rsConfig("PresentValidUnit")
    MoneyExchangePoint = rsConfig("MoneyExchangePoint")
    MoneyExchangeValidDay = rsConfig("MoneyExchangeValidDay")
    UserExpExchangePoint = rsConfig("UserExpExchangePoint")
    UserExpExchangeValidDay = rsConfig("UserExpExchangeValidDay")
    If MoneyExchangePoint <= 0 Then MoneyExchangePoint = 1
    If MoneyExchangeValidDay <= 0 Then MoneyExchangeValidDay = 1
    If UserExpExchangePoint <= 0 Then UserExpExchangePoint = 1
    If UserExpExchangeValidDay <= 0 Then UserExpExchangeValidDay = 1
    
    PointName = rsConfig("PointName")
    PointUnit = rsConfig("PointUnit")
    EmailOfRegCheck = rsConfig("EmailOfRegCheck")

    MailObject = rsConfig("MailObject")
    MailServer = rsConfig("MailServer")
    MailServerUserName = rsConfig("MailServerUserName")
    MailServerPassWord = rsConfig("MailServerPassword")
    MailDomain = rsConfig("MailDomain")

    PhotoObject = rsConfig("PhotoObject")
    Thumb_DefaultWidth = rsConfig("Thumb_DefaultWidth")
    Thumb_DefaultHeight = rsConfig("Thumb_DefaultHeight")
    Thumb_Arithmetic = rsConfig("Thumb_Arithmetic")
    PhotoQuality = PE_CLng(rsConfig("PhotoQuality"))
    Thumb_BackgroundColor = rsConfig("Thumb_BackgroundColor")
    If Watermark_Position = "" Then Watermark_Position = "1"
    If PhotoQuality < 50 Then PhotoQuality = 90
    If PhotoQuality > 100 Then PhotoQuality = 90
    If Thumb_BackgroundColor = "" Then Thumb_BackgroundColor = "#CCCCCC"
    Watermark_Images_Transparence = Watermark_Images_Transparence / 100
    Watermark_Text_FontColor = "&H" & Replace(Right(Watermark_Text_FontColor, 6), "#", "")
    Watermark_Images_BackgroundColor = "&H" & Replace(Right(Watermark_Images_BackgroundColor, 6), "#", "")
    Thumb_BackgroundColor = "&H" & Replace(Right(Thumb_BackgroundColor, 6), "#", "")
    Watermark_Type = rsConfig("Watermark_Type")
    Watermark_Text = rsConfig("Watermark_Text")
    Watermark_Text_FontName = rsConfig("Watermark_Text_FontName")
    Watermark_Text_FontSize = rsConfig("Watermark_Text_FontSize")
    Watermark_Text_FontColor = rsConfig("Watermark_Text_FontColor")
    Watermark_Text_Bold = rsConfig("Watermark_Text_Bold")
    Watermark_Images_FileName = rsConfig("Watermark_Images_FileName")
    Watermark_Images_Transparence = rsConfig("Watermark_Images_Transparence")
    Watermark_Images_BackgroundColor = rsConfig("Watermark_Images_BackgroundColor")
    Watermark_Position_X = rsConfig("Watermark_Position_X")
    Watermark_Position_Y = rsConfig("Watermark_Position_Y")
    Watermark_Position = PE_CLng(rsConfig("Watermark_Position"))

    SearchInterval = rsConfig("SearchInterval")
    SearchResultNum = rsConfig("SearchResultNum")
    MaxPerPage_SearchResult = rsConfig("MaxPerPage_SearchResult")
    SearchContent = rsConfig("SearchContent")
    
    EnableGuestBuy = rsConfig("EnableGuestBuy")
    IncludeTax = rsConfig("IncludeTax")
    TaxRate = rsConfig("TaxRate")
    Prefix_OrderFormNum = rsConfig("Prefix_OrderFormNum")
    Prefix_PaymentNum = rsConfig("Prefix_PaymentNum")
    
    MyCountry = rsConfig("Country")
    MyProvince = rsConfig("Province")
    MyCity = rsConfig("City")
    MyPostcode = rsConfig("Postcode")

    EmailOfOrderConfirm = rsConfig("EmailOfOrderConfirm") & ""
    EmailOfSendCard = rsConfig("EmailOfSendCard") & ""
    EmailOfReceiptMoney = rsConfig("EmailOfReceiptMoney") & ""
    EmailOfRefund = rsConfig("EmailOfRefund") & ""
    EmailOfInvoice = rsConfig("EmailOfInvoice") & ""
    EmailOfDeliver = rsConfig("EmailOfDeliver") & ""
    
    
    GuestBook_EnableVisitor = rsConfig("GuestBook_EnableVisitor")
    EnableGuestBookCheck = rsConfig("GuestBookCheck")
    GuestBook_EnableManageRubbish = rsConfig("GuestBook_EnableManageRubbish")
    PresentExpPerLogin = rsConfig("PresentExpPerLogin")
    GuestBook_ManageRubbish = rsConfig("GuestBook_ManageRubbish")
    GuestBook_ShowIP = rsConfig("GuestBook_ShowIP")
    GuestBook_IsAssignSort = rsConfig("GuestBook_IsAssignSort")
    GuestBook_MaxPerPage = rsConfig("GuestBook_MaxPerPage")

    EnableRss = rsConfig("EnableRss")
    RssCodeType = rsConfig("RssCodeType")

    EnableWap = rsConfig("EnableWap")
    WapLogo = rsConfig("WapLogo")
    EnableWapPl = rsConfig("EnableWapPl")
    ShowWapShop = rsConfig("ShowWapShop")
    ShowWapAppendix = rsConfig("ShowWapAppendix")
    ShowWapManage = rsConfig("ShowWapManage")
    
    EnableSMS = FoundInArr(rsConfig("Modules"), "SMS", ",")
    SMSUserName = rsConfig("SMSUserName")
    SMSKey = rsConfig("SMSKey")

    SendMessageToAdminWhenOrder = rsConfig("SendMessageToAdminWhenOrder")
    SendMessageToMemberWhenPaySuccess = rsConfig("SendMessageToMemberWhenPaySuccess")
    Mobiles = rsConfig("Mobiles")
    MessageOfOrder = rsConfig("MessageOfOrder")

    MessageOfOrderConfirm = rsConfig("MessageOfOrderConfirm") & ""
    MessageOfSendCard = rsConfig("MessageOfSendCard") & ""
    MessageOfReceiptMoney = rsConfig("MessageOfReceiptMoney") & ""
    MessageOfRefund = rsConfig("MessageOfRefund") & ""
    MessageOfInvoice = rsConfig("MessageOfInvoice") & ""
    MessageOfDeliver = rsConfig("MessageOfDeliver") & ""
    
    MessageOfAddRemit = rsConfig("MessageOfAddRemit")
    MessageOfAddIncome = rsConfig("MessageOfAddIncome")
    MessageOfAddPayment = rsConfig("MessageOfAddPayment")
    MessageOfExchangePoint = rsConfig("MessageOfExchangePoint")
    MessageOfAddPoint = rsConfig("MessageOfAddPoint")
    MessageOfMinusPoint = rsConfig("MessageOfMinusPoint")
    MessageOfExchangeValid = rsConfig("MessageOfExchangeValid")
    MessageOfAddValid = rsConfig("MessageOfAddValid")
    MessageOfMinusValid = rsConfig("MessageOfMinusValid")
    
    If Right(InstallDir, 1) <> "/" Then
        InstallDir = InstallDir & "/"
    End If
    If rsConfig("SiteUrlType") = 1 Then
        If Right(SiteUrl, 1) = "/" Then SiteUrl = Left(SiteUrl, Len(SiteUrl) - 1)
        If Left(SiteUrl, 7) <> "http://" Then SiteUrl = "http://" & SiteUrl
        strInstallDir = SiteUrl & InstallDir
    Else
        strInstallDir = InstallDir
    End If

    rsConfig.Close
    Set rsConfig = Nothing
    
End Sub

Sub IsIPlock()
    If session("IPlock") = "" Then
        session("IPlock") = ChecKIPlock(LockIPType, LockIP, UserTrueIP)
    End If
    If session("IPlock") = True Then
        Response.Write "对不起！您的IP（" & UserTrueIP & "）被系统限定。您可以和站长联系。"
        Response.End
    End If
End Sub
%>
