<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.Common.Purview.asp"-->
<!--#include file="Admin_CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Server.ScriptTimeOut = 9999999

ChannelID = PE_CLng(Trim(Request("ChannelID")))
If ChannelID > 0 Then
    Call GetChannel(ChannelID)
End If

If NeedCheckComeUrl = True Then
    Call CheckComeUrl
End If

'检查管理员是否登录
Dim AdminID, AdminName, AdminPassword, RndPassword, AdminLoginCode, AdminPurview, PurviewPassed
Dim AdminPurview_Channel, AdminPurview_Others, AdminPurview_GuestBook
Dim arrClass_GuestBook, arrKind_House
Dim rsGetAdmin, sqlGetAdmin
Dim arrPurview(30), PurviewIndex, strThisFile
AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
AdminLoginCode = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminLoginCode")))
If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Or (EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode) Then
    Call WriteEntry(1, "", "管理员未登录")
    Call CloseConn
    Response.redirect "Admin_login.asp"
End If
sqlGetAdmin = "select * from PE_Admin where AdminName='" & AdminName & "' and Password='" & AdminPassword & "'"
Set rsGetAdmin = Server.CreateObject("adodb.recordset")
rsGetAdmin.Open sqlGetAdmin, Conn, 1, 1
If rsGetAdmin.BOF And rsGetAdmin.EOF Then
    Call WriteEntry(4, "", "用户名或密码错误")
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
    Call CloseConn
    Response.redirect "Admin_login.asp"
Else
    If rsGetAdmin("EnableMultiLogin") <> True And Trim(rsGetAdmin("RndPassword")) <> RndPassword Then
        Response.write "<br><p align=center><font color='red'>对不起，为了系统安全，本系统不允许两个人使用同一个管理员帐号进行登录！</font></p><p>因为现在有人已经在其他地方使用此管理员帐号进行登录了，所以你将不能继续进行后台管理操作。</p><p>你可以<a href='Admin_Login.asp' target='_top'>点此重新登录</a>。</p>"
        Call WriteEntry(1, AdminName, "两人使用同一管理员帐号")
        rsGetAdmin.Close
        Set rsGetAdmin = Nothing
        Call CloseConn
        Response.End
    End If
End If
AdminID = rsGetAdmin("ID")
UserName = rsGetAdmin("UserName")
AdminPurview = rsGetAdmin("Purview")
AdminPurview_Others = rsGetAdmin("AdminPurview_Others")
AdminPurview_GuestBook = rsGetAdmin("AdminPurview_GuestBook")
arrClass_View = rsGetAdmin("arrClass_View") & ""
arrClass_Input = rsGetAdmin("arrClass_Input") & ""
arrClass_Check = rsGetAdmin("arrClass_Check") & ""
arrClass_Manage = rsGetAdmin("arrClass_Manage") & ""
arrClass_GuestBook = rsGetAdmin("arrClass_GuestBook") & ""
arrKind_House = Split(rsGetAdmin("arrClass_House") & "", "|||")


PurviewPassed = False   '默认设置为没有权限
If AdminPurview = 1 Then   '如果是超级管理员，直接有所有权限
    PurviewPassed = True
Else
    '如果是普通管理员，根据文件的设置判断是否有相应的权限
    Select Case PurviewLevel
    Case 0       '如果不进行权限检查，直接有权限
        PurviewPassed = True
    Case 1       '如果要求超级管理员，则直接判断没有权限
        PurviewPassed = False
    Case 2    '如果要求普通管理员
        If PurviewLevel_Channel <= 0 Then  '如果不要检查频道权限设置
            PurviewPassed = True
        Else
            AdminPurview_Channel = PE_CLng(rsGetAdmin("AdminPurview_" & ChannelDir))
            If AdminPurview_Channel = 0 Then AdminPurview_Channel = 5
            
            If AdminPurview_Channel <= PurviewLevel_Channel Then
                PurviewPassed = True
            End If
        End If
        If PurviewLevel_Others <> "" Then  '如果还要检查其他权限，则权限以最后检查的为准
            PurviewPassed = CheckPurview_Other(AdminPurview_Others, PurviewLevel_Others)
        End If
    End Select
End If

If PurviewPassed = False Then
    Response.write "<br><p align=center><font color='red'>对不起，你没有此项操作的权限。</font></p>"
    Call WriteEntry(1, AdminName, "越权使用")
    Call CloseConn
    Response.End
End If

%>
