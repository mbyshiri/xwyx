<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim strPath
If EnableUserReg <> True Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>对不起，本站暂停新用户注册服务！</li>"
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If

strHtml = GetTemplate(0, 19, 0)
Call ReplaceCommonLabel

strPath = "您现在的位置：&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;新会员注册"

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> 新会员注册")
strHtml = Replace(strHtml, "{$ShowPath}", strPath)

strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))

strHtml = Replace(strHtml, "{$Display_Homepage}", IsDisplay(FoundInArr(RegFields_MustFill, "Homepage", ",")))
strHtml = Replace(strHtml, "{$Display_QQ}", IsDisplay(FoundInArr(RegFields_MustFill, "QQ", ",")))
strHtml = Replace(strHtml, "{$Display_ICQ}", IsDisplay(FoundInArr(RegFields_MustFill, "ICQ", ",")))
strHtml = Replace(strHtml, "{$Display_MSN}", IsDisplay(FoundInArr(RegFields_MustFill, "MSN", ",")))
strHtml = Replace(strHtml, "{$Display_Yahoo}", IsDisplay(FoundInArr(RegFields_MustFill, "Yahoo", ",")))
strHtml = Replace(strHtml, "{$Display_UC}", IsDisplay(FoundInArr(RegFields_MustFill, "UC", ",")))
strHtml = Replace(strHtml, "{$Display_Aim}", IsDisplay(FoundInArr(RegFields_MustFill, "Aim", ",")))
strHtml = Replace(strHtml, "{$Display_OfficePhone}", IsDisplay(FoundInArr(RegFields_MustFill, "OfficePhone", ",")))
strHtml = Replace(strHtml, "{$Display_HomePhone}", IsDisplay(FoundInArr(RegFields_MustFill, "HomePhone", ",")))
strHtml = Replace(strHtml, "{$Display_Fax}", IsDisplay(FoundInArr(RegFields_MustFill, "Fax", ",")))
strHtml = Replace(strHtml, "{$Display_Mobile}", IsDisplay(FoundInArr(RegFields_MustFill, "Mobile", ",")))
strHtml = Replace(strHtml, "{$Display_PHS}", IsDisplay(FoundInArr(RegFields_MustFill, "PHS", ",")))
strHtml = Replace(strHtml, "{$Display_Region}", IsDisplay(FoundInArr(RegFields_MustFill, "Region", ",")))
strHtml = Replace(strHtml, "{$Display_Address}", IsDisplay(FoundInArr(RegFields_MustFill, "Address", ",")))
strHtml = Replace(strHtml, "{$Display_ZipCode}", IsDisplay(FoundInArr(RegFields_MustFill, "ZipCode", ",")))
strHtml = Replace(strHtml, "{$Display_TrueName}", IsDisplay(FoundInArr(RegFields_MustFill, "TrueName", ",")))
strHtml = Replace(strHtml, "{$Display_Birthday}", IsDisplay(FoundInArr(RegFields_MustFill, "Birthday", ",")))
strHtml = Replace(strHtml, "{$Display_IDCard}", IsDisplay(FoundInArr(RegFields_MustFill, "IDCard", ",")))
strHtml = Replace(strHtml, "{$Display_Vocation}", IsDisplay(FoundInArr(RegFields_MustFill, "Vocation", ",")))
strHtml = Replace(strHtml, "{$Display_Company}", IsDisplay(FoundInArr(RegFields_MustFill, "Company", ",")))
strHtml = Replace(strHtml, "{$Display_Department}", IsDisplay(FoundInArr(RegFields_MustFill, "Department", ",")))
strHtml = Replace(strHtml, "{$Display_PosTitle}", IsDisplay(FoundInArr(RegFields_MustFill, "PosTitle", ",")))
strHtml = Replace(strHtml, "{$Display_Marriage}", IsDisplay(FoundInArr(RegFields_MustFill, "Marriage", ",")))
strHtml = Replace(strHtml, "{$Display_Income}", IsDisplay(FoundInArr(RegFields_MustFill, "Income", ",")))
strHtml = Replace(strHtml, "{$Display_UserFace}", IsDisplay(FoundInArr(RegFields_MustFill, "UserFace", ",")))
strHtml = Replace(strHtml, "{$Display_FaceWidth}", IsDisplay(FoundInArr(RegFields_MustFill, "FaceWidth", ",")))
strHtml = Replace(strHtml, "{$Display_FaceHeight}", IsDisplay(FoundInArr(RegFields_MustFill, "FaceHeight", ",")))
strHtml = Replace(strHtml, "{$Display_Sign}", IsDisplay(FoundInArr(RegFields_MustFill, "Sign", ",")))
strHtml = Replace(strHtml, "{$Display_SpareEmail}", IsDisplay(False))
strHtml = Replace(strHtml, "{$Display_Privacy}", IsDisplay(FoundInArr(RegFields_MustFill, "Privacy", ",")))
strHtml = Replace(strHtml, "{$Display_CheckCode}", IsDisplay(EnableCheckCodeOfReg))
strHtml = Replace(strHtml, "{$Display_QAofReg}", IsDisplay(EnableQAofReg))
strHtml = Replace(strHtml, "{$QAofReg}", GetQAofReg(QAofReg))

Response.Write strHtml
Call CloseConn


Function IsDisplay(Display)
    If Display = True Then
        IsDisplay = ""
    Else
        IsDisplay = " Style='display:none'"
    End If
End Function
Function GetQAofReg(QAofReg)
    Dim arrQAofReg, i, strTemp
    arrQAofReg = Split(QAofReg & "", "$$$")
    For i = 0 To 2
        If Trim(arrQAofReg(i * 2)) <> "" Then
            strTemp = strTemp & CStr(i + 1) & "、" & arrQAofReg(i * 2) & "<br><input type='text' name='RegAnswer" & i & "' size='30'><br><br>"
        End If
    Next
    GetQAofReg = strTemp
End Function
%>
