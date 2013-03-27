<!-- #include File="../Start.asp" -->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim ChannelID, ShowType
Dim FilesPath, sql, rs
Dim AdminName,UserPassword,LastPassword
Dim sqlChannel,rsChannel
Dim ModuleType,DialogType,IsUpload,Anonymous


AdminName = ReplaceBadChar(Trim(request.Cookies(Site_Sn)("AdminName")))
UserName = ReplaceBadChar(Trim(request.Cookies(Site_Sn)("UserName")))
UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))

ChannelID = Trim(request("ChannelID"))
ShowType = PE_Clng(Trim(request("ShowType")))
Anonymous = PE_Clng(Trim(request("Anonymous")))


If ChannelID = "" Then
    response.write "频道参数丢失！"
    response.End
Else
    ChannelID = PE_CLng(ChannelID)
End If

If AdminName = "" And UserName = "" And Anonymous = 0  Then
    Response.Write "请先登录后再使用此功能！"
    Response.End
ElseIf Anonymous = 1 Then
    If ShowAnonymous = True Then
        Dim  rsGroup
        Set rsGroup = Conn.Execute("select * from PE_UserGroup where GroupID=-1")
        arrClass_Input = Trim(rsGroup("arrClass_Input"))
        UserSetting = Split(Trim(rsGroup("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        rsGroup.Close
        Set rsGroup = Nothing	
        If PE_CBool(PE_CLng(UserSetting(9))) = True Then
            IsUpload = True
        Else
            IsUpload = False	
        End If
    Else
        IsUpload = False			
    End If 
Else
    If AdminName <> "" And (ShowType=0 or ShowType=4 or ShowType=5) Then
        IsUpload = True
    ElseIf UserName <> "" Then
        If (UserName = "" Or UserPassword = "" Or LastPassword = "") Then
            IsUpload = False
        Else
            sql = "SELECT U.UserID,U.SpecialPermission,U.UserSetting,G.GroupSetting FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
            sql = sql & " UserName='" & UserName & "' AND UserPassword='" & UserPassword & "' AND LastPassword='" & LastPassword & "' and IsLocked=" & PE_False & ""
            Set rs = Conn.Execute(sql)
            If rs.BOF And rs.EOF Then
                IsUpload = False
            Else
                If rs("SpecialPermission") = True Then
                    UserSetting = Split(Trim(rs("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                Else
                    UserSetting = Split(Trim(rs("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                End If
                If CBool(PE_CLng(UserSetting(9))) = True Then
                    IsUpload = True
                End If
            End If
            Set rs = Nothing
        End If
    End If
End If

If PE_Clng(ChannelID) < 0 Then
    Select Case PE_Clng(ChannelID)
    Case -1
        IsUpload = True
    Case -2
        IsUpload = True
    Case -3
        IsUpload = True
    End Select
Else

    sqlChannel = "select ChannelDir,UploadDir,Disabled,EnableUploadFile,ModuleType from PE_Channel where ChannelID=" & PE_Clng(ChannelID)
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 1
    If rsChannel.BOF And rsChannel.EOF Then
        IsUpload = False
    Else
        If rsChannel("Disabled") = True Then
            IsUpload = False
        Else
            If rsChannel("EnableUploadFile") = False Then
                IsUpload = False
            End If
            FilesPath = InstallDir & rsChannel("ChannelDir") & "/" & rsChannel("UploadDir") & "/"
            ModuleType = rsChannel("ModuleType")
        End If
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    If IsUpload = True Then
        Select Case ModuleType
        Case 1, 2, 3, 4, 5, 6, 7, 8 
            IsUpload = True
        Case Else
            IsUpload = False
        End Select
    End If
End If

Call CloseConn

%>