<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim sql, rs
Dim CookieDate
Dim UserPassword, RndPassword, CheckCode

UserName = Replace(ReplaceBadChar(Trim(Request("UserName"))), ".", "")
UserPassword = ReplaceBadChar(Trim(Request("UserPassword")))
CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
CookieDate = PE_CLng(Trim(Request("CookieDate")))
If InStr(ComeUrl, "Reg/") > 0 Or InStr(LCase(ComeUrl), "user_login.asp") Or InStr(ComeUrl, "User_ChkLogin.asp") > 0 Then ComeUrl = strInstallDir & "User/"
If ComeUrl = "" Then ComeUrl = strInstallDir
If EnableCheckCodeOfLogin = True Then
    If Trim(Session("CheckCode")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>验证码超时失效。</li>"
    End If
    If CheckCode <> Session("CheckCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>验证码错误，请重新输入。</li>"
    End If
End If
If UserName = "" Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>用户名不能为空！</li>"
End If

If UserPassword = "" Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>密码不能为空！</li>"
End If
If FoundErr = True Then '输出错误结果
    Call WriteErrMsg
    Response.End
End If

Set rs = Conn.Execute("select UserID,UserName,UserPassword,LastPresentTime,LastPresentTime,IsLocked from PE_User where UserName='" & UserName & "'")
If rs.bof And rs.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>用户不存在！！！</li>"
Else
    UserPassword = MD5(UserPassword, 16)
    If UserPassword <> rs("UserPassword") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>密码错误！！！</li>"
    End If
    If rs("IsLocked") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>用户已经被锁住，无法登录。如果您需要解锁，请与管理员联系。</li>"
    End If
    If FoundErr = False Then
        RndPassword = GetRndPassword(16)
        '更新登录状态
        Conn.Execute ("Update PE_User Set LastPassword='" & RndPassword & "',LastLoginIP='" & UserTrueIP & "',LastLoginTime=" & PE_Now & ",LoginTimes=LoginTimes+1 where UserID=" & rs("UserID") & "")
        Session("UserID") = rs("UserID")
        
        If PresentExpPerLogin > 0 Then
            If DateDiff("D", rs("LastPresentTime"), Now()) > 0 Or IsNull(rs("LastPresentTime")) Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExpPerLogin & ",LastPresentTime=" & PE_Now & " where UserID=" & rs("UserID") & "")
                UserExp = UserExp + PresentExpPerLogin
            End If
        End If

        '更新购物车的用户名
        Dim CartID
        CartID = ReplaceBadChar(Trim(Request.Cookies("Cart" & Site_Sn)("CartID")))
        Conn.Execute ("update PE_ShoppingCarts set UserName='" & UserName & "' where CartID='" & CartID & "'")

        If Enable_SubDomain Then
            Dim iSubDomainIndex, strSite_Sn
            For iSubDomainIndex = 0 To UBound(arrSubDomains)
                strSite_Sn = LCase(arrSubDomains(iSubDomainIndex) & Replace(Replace(DomainRoot & InstallDir, "/", ""), ".", ""))
                Response.Cookies(strSite_Sn).Domain = DomainRoot
                Select Case CookieDate
                    Case 0
                        'not save
                    Case 1
                        Response.Cookies(strSite_Sn).Expires = Date + 1
                    Case 2
                        Response.Cookies(strSite_Sn).Expires = Date + 31
                    Case 3
                        Response.Cookies(strSite_Sn).Expires = Date + 365
                End Select
                Response.Cookies(strSite_Sn)("UserName") = UserName
                Response.Cookies(strSite_Sn)("UserPassword") = UserPassword
                Response.Cookies(strSite_Sn)("LastPassword") = RndPassword
                Response.Cookies(strSite_Sn)("CookieDate") = CookieDate
            Next
        Else
            Select Case CookieDate
                Case 0
                    'not save
                Case 1
                    Response.Cookies(Site_Sn).Expires = Date + 1
                Case 2
                    Response.Cookies(Site_Sn).Expires = Date + 31
                Case 3
                    Response.Cookies(Site_Sn).Expires = Date + 365
            End Select
            Response.Cookies(Site_Sn)("UserName") = UserName
            Response.Cookies(Site_Sn)("UserPassword") = UserPassword
            Response.Cookies(Site_Sn)("LastPassword") = RndPassword
            Response.Cookies(Site_Sn)("CookieDate") = CookieDate
        End If

        
        If API_Enable Then
            Call WriteSuccessMsg("您已成功登录，欢迎您的光临!", ComeUrl)
            
            '输出整合登录脚本
            sPE_Items(conSyskey, 1) = MD5(UserName & API_Key, 16)
            sPE_Items(conUsername, 1) = UserName
            sPE_Items(conPassword, 1) = UserPassword
            sPE_Items(conSavecookie, 1) = CookieDate
            Dim iIndex
            For iIndex = 0 To UBound(arrUrlsSP2)
                Response.Write "<iframe frameborder='0' width='1' height='1' src='" & arrUrlsSP2(iIndex) & "?syskey=" & sPE_Items(conSyskey, 1) & "&username=" & sPE_Items(conUsername, 1) & "&password=" & sPE_Items(conPassword, 1) & "&savecookie=" & sPE_Items(conSavecookie, 1) & "'></iframe>" & vbCrLf
            Next
        Else
            Response.Redirect ComeUrl
        End If
    End If
End If
rs.Close
Set rs = Nothing

If FoundErr = True Then '输出错误结果
    Call WriteErrMsg
End If
Call CloseConn

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************

Sub WriteErrMsg()
    Response.Write "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='../Images/style.css' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center'><td height='22' class='title'><strong>错误信息</strong></td></tr>" & vbCrLf
    Response.Write "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & ErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center'><td class='tdbg'><a href=""User_Login.asp?ComeUrl=" & ComeUrl & """>&lt;&lt; 返回登录页面</a></td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub
%>
