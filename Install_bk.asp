<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.MD5.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Edition.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

InstallDir = Trim(Request.ServerVariables("SCRIPT_NAME"))
InstallDir = Left(InstallDir, InStrRev(LCase(InstallDir), "/"))
Site_Sn = Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & InstallDir), "/", ""), ".", "")

If action = "ChkLogin" Then
    Call ChkLogin
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>动易 SiteWeaver " & SystemEdition & " 6.6版安装向导</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='" & AdminDir & "/Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Dim AgreeLicense
AgreeLicense = Session("AgreeLicense")
If AgreeLicense = "" Then
    AgreeLicense = Trim(Request("AgreeLicense"))
    Session("AgreeLicense") = AgreeLicense
End If
If AgreeLicense <> "Yes" Then
    Call ShowLicense
    Response.End
End If

If CheckAdminLogin = False Then
    Call Check  '检查管理员权限
Else
    Dim sqlConfig, rsConfig
    sqlConfig = "select * from PE_Config"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn, 1, 3
    If rsConfig.BOF And rsConfig.EOF Then
        Response.Write "网站配置数据丢失！系统无法正常运行！"
    Else
        If action = "" Then
            action = "Step1"
        End If
        Select Case action
        Case "Step1"
            Call Step1  '网站信息配置1
        Case "Step2"
            Call Step2  '导入模板
        Case "Stepdel"
            Call Stepdel
        End Select
    End If
    rsConfig.Close
    Set rsConfig = Nothing
End If

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"

Call CloseConn

Sub ShowLicense()
    Response.Write "<form name='myform' id='myform' method='POST' action='Install.asp'>" & vbCrLf
    Response.Write "  <table width='60%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
    Response.Write "    <tr class='topbg'>" & vbCrLf
    Response.Write "      <td height='22' align='center'><strong>阅读许可协议</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='center'><textarea name='License' cols='120' rows='30' id='License' readonly>"
%>
<!--#include file="License.txt"-->
<%
    Response.Write "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='left'><input name='AgreeLicense' type='checkbox' id='AgreeLicense' value='Yes' onclick='document.myform.submit.disabled=!this.checked;'><label for='AgreeLicense'>我已经阅读并同意此协议</label></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value=' 下一步 ' disabled>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf

End Sub

Sub Check()
    Response.Write "<br><br>" & vbCrLf
    Response.Write "<form name='myform' id='myform' method='POST' action='Install.asp'>" & vbCrLf
    Response.Write "  <table width='50%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
    Response.Write "    <tr class='topbg'>" & vbCrLf
    Response.Write "      <td height='22' colspan='2' align='center'><strong>管理员登录</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' class='tdbg5'><strong>用户名称：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='UserName' type='text' id='UserName' value='' size='30' maxlength='20' style='width:150px;'> 默认用户名为：admin</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' class='tdbg5'><strong>用户密码：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='password' type='password' id='password' value='' size='30' maxlength='20' style='width:150px;'> 默认用户密码为：admin888</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' class='tdbg5'><strong>验 证 码：</strong></td>" & vbCrLf
    Response.Write "      <td valign='top'><input name='CheckCode' type='text' id='CheckCode' value='' size='6' maxlength='6'> <img id='checkcode' src='inc/checkcode.asp' style='border: 1px solid #ffffff'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='ChkLogin'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value=' 登 录 '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Function CheckAdminLogin()
    Dim AdminName, AdminPassword, RndPassword
    AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
    AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
    RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
    If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Then
        CheckAdminLogin = False
    Else
        CheckAdminLogin = True
    End If
End Function

Sub ChkLogin()
    Dim sql, rs
    Dim UserName, Password, CheckCode, RndPassword
    UserName = ReplaceBadChar(Trim(Request("UserName")))
    Password = ReplaceBadChar(Trim(Request("Password")))
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))

    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>用户名不能为空！</li>"
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>密码不能为空！</li>"
    End If
    If CheckCode = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>验证码不能为空！</li>"
    End If
    If Trim(Session("CheckCode")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>你登录时间过长，请重新返回登录页面进行登录。</li>"
    End If
    If CheckCode <> Session("CheckCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>您输入的确认码和系统产生的不一致，请重新输入。</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    Password = md5(Password, 16)
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_Admin where Password='" & Password & "' and AdminName='" & UserName & "'"
    rs.Open sql, Conn, 1, 3
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>用户名或密码错误！！！</li>"
    Else
        If Password <> rs("Password") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>用户名或密码错误！！！</li>"
        End If
    End If
    If FoundErr = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    RndPassword = GetRndPassword(16)
    Response.Cookies(Site_Sn)("AdminName") = rs("AdminName")
    Response.Cookies(Site_Sn)("AdminPassword") = rs("Password")
    Response.Cookies(Site_Sn)("RndPassword") = RndPassword
    rs("RndPassword") = RndPassword
    rs.Update
    rs.Close
    Set rs = Nothing
    Response.Redirect "install.asp"
End Sub

Sub Step1()
    Response.Write "<form name='myform' id='myform' method='POST' action='Install.asp'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
    Response.Write "    <tr class='topbg'>" & vbCrLf
    Response.Write "      <td height='22' colspan='2' align='center'><strong>动易 SiteWeaver " & SystemEdition & " 6.5版安装向导</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站名称：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteName' type='text' id='SiteName' value='" & rsConfig("SiteName") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站标题：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteTitle' type='text' id='SiteTitle' value='" & rsConfig("SiteTitle") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站地址：</strong><br>请填写完整URL地址</td>" & vbCrLf
    Response.Write "      <td><input name='SiteUrl' type='text' id='SiteUrl' value='" & rsConfig("SiteUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><font color=red><strong>安装目录：</strong><br>系统安装目录（相对于根目录的位置）<br>系统会自动获得正确的路径，但需要手工保存设置。</font></td>" & vbCrLf
    Response.Write "      <td><input name='InstallDir' type='text' id='InstallDir' value='" & InstallDir & "' size='40' maxlength='30' readonly></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>LOGO地址：</strong><br>请填写完整URL地址</td>" & vbCrLf
    Response.Write "      <td><input name='LogoUrl' type='text' id='LogoUrl' value='" & rsConfig("LogoUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>Banner地址：</strong><br>请填写完整URL地址</td>" & vbCrLf
    Response.Write "      <td><input name='BannerUrl' type='text' id='BannerUrl' value='" & rsConfig("BannerUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>FSO(FileSystemObject)组件的名称：</strong><br>某些网站为了安全，将FSO组件的名称进行更改以达到禁用FSO的目的。如果你的网站是这样做的，请在此输入更改过的名称。</td>" & vbCrLf
    Response.Write "      <td><input name='objName_FSO' type='text' id='objName_FSO' value='" & rsConfig("objName_FSO") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>后台管理目录：</strong><br>为了安全，您可以修改后台管理目录（默认为Admin），改过以后，需要再设置此处</td>" & vbCrLf
    Response.Write "      <td><input name='AdminDir' type='text' id='AdminDir' value='" & rsConfig("AdminDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站广告目录：</strong><br>为了不让广告拦截软件拦截网站的广告，您可以修改广告JS的存放目录（默认为AD），改过以后，需要再设置此处</td>" & vbCrLf
    Response.Write "      <td><input name='ADDir' type='text' id='ADDir' value='" & rsConfig("ADDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>站长姓名：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterName' type='text' id='WebmasterName' value='" & rsConfig("WebmasterName") & "' size='40' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>站长信箱：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterEmail' type='text' id='WebmasterEmail' value='" & rsConfig("WebmasterEmail") & "' size='40' maxlength='100'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>版权信息：</strong><br>支持HTML标记，不能使用双引号</td>" & vbCrLf
    Response.Write "      <td><textarea name='Copyright' cols='60' rows='4' id='Copyright'>" & rsConfig("Copyright") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Step2'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value=' 下一步 '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Step2()
    Call SaveConfig
    If SystemDatabaseType = "SQL" Then
        Call DoImport
    End If
    Call DelTagFile
    Call CreatSkinFile
    Call ClearSiteCache
    Call WriteSuccessMsg("系统安装完成！现在你可以<a href='Index.asp'>使用系统</a>了。<br>为了<font color='red'>系统安全</font>，请点击下面的按钮删除此安装文件（Install.asp）<br><br><div align='center'><input name='delfile' type='button' id='delfile' value=' 删除此安装文件 ' onclick=""location='install.asp?Action=Stepdel'""></div><br>", ComeUrl)
End Sub

Sub DelTagFile()
    On Error Resume Next
    If fso.FileExists(Server.MapPath("NotInsalled.txt")) Then
        fso.DeleteFile Server.MapPath("NotInsalled.txt")
    End If
End Sub

Sub Stepdel()
    On Error Resume Next
    If fso.FileExists(Server.MapPath("NotInsalled.txt")) Then
        fso.DeleteFile Server.MapPath("NotInsalled.txt")
    End If
    If fso.FileExists(Server.MapPath("install.asp")) Then
        fso.DeleteFile Server.MapPath("install.asp")
    End If
    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<br><li>删除此安装文件（Install.asp）失败，错误原因：" & Err.Description & "<br>请手动删除此文件。"
        Err.Clear
        Exit Sub
    Else
        Call WriteSuccessMsg("删除此安装文件（Install.asp）成功！<br><br><a href='Index.asp'>点此开始使用系统</a>", ComeUrl)
    End If
    Response.Cookies(Site_Sn)("AdminName") = ""
    Response.Cookies(Site_Sn)("AdminPassword") = ""
    Response.Cookies(Site_Sn)("RndPassword") = ""
End Sub

Sub SaveConfig()
    Dim sqlConfig, rsConfig
    If action = "Step2" Then
        sqlConfig = "select * from PE_Config"
        Set rsConfig = Server.CreateObject("ADODB.Recordset")
        rsConfig.Open sqlConfig, Conn, 1, 3
        If rsConfig.BOF And rsConfig.EOF Then
            rsConfig.addnew
        End If
        rsConfig("SiteName") = Trim(Request("SiteName"))
        rsConfig("SiteTitle") = Trim(Request("SiteTitle"))
        rsConfig("SiteUrl") = Trim(Request("SiteUrl"))
        rsConfig("InstallDir") = InstallDir
        rsConfig("LogoUrl") = Trim(Request("LogoUrl"))
        rsConfig("BannerUrl") = Trim(Request("BannerUrl"))
        rsConfig("WebmasterName") = Trim(Request("WebmasterName"))
        rsConfig("WebmasterEmail") = Trim(Request("WebmasterEmail"))
        rsConfig("Copyright") = Trim(Request("Copyright"))
        rsConfig("objName_FSO") = Trim(Request("objName_FSO"))
        rsConfig("AdminDir") = Trim(Request("AdminDir"))
        rsConfig("ADDir") = Trim(Request("ADDir"))

        rsConfig.Update
        rsConfig.Close
        Set rsConfig = Nothing
    End If
End Sub

Sub DoImport()
    'On Error Resume Next
    Dim mdbname, tconn, trs, rs, sql
    mdbname = "Database/SiteWeaver.mdb"
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If
    '导入方案
    Set rs = tconn.Execute("select * from PE_TemplateProject where TemplateProjectID order by TemplateProjectID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_TemplateProject", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("TemplateProjectID") = rs("TemplateProjectID")
            trs("TemplateProjectName") = rs("TemplateProjectName")
            trs("Intro") = rs("Intro")
            trs("IsDefault") = rs("IsDefault")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    
    '导入模板
    Set rs = tconn.Execute(" select * from PE_Template order by TemplateID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Template", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("ChannelID") = rs("ChannelID")
            trs("TemplateName") = rs("TemplateName")
            trs("TemplateType") = rs("TemplateType")
            trs("TemplateContent") = rs("TemplateContent")
            trs("IsDefault") = rs("IsDefault")
            trs("IsDefaultInProject") = rs("IsDefaultInProject")
            trs("ProjectName") = rs("ProjectName")
            trs("Deleted") = rs("Deleted")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    '导入自定义标签
    Set rs = tconn.Execute(" select * from PE_Label order by LabelID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Label", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("LabelName") = rs("LabelName")
            trs("LabelClass") = rs("LabelClass")
            trs("LabelType") = rs("LabelType")
            trs("PageNum") = rs("PageNum")
            trs("reFlashTime") = rs("reFlashTime")
            trs("fieldlist") = rs("fieldlist")
            trs("LabelIntro") = rs("LabelIntro")
            trs("Priority") = rs("Priority")
            trs("LabelContent") = rs("LabelContent")
            trs("AreaCollectionID") = rs("AreaCollectionID")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    '导入风格
    Set rs = tconn.Execute(" select * from PE_Skin order by SkinID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Skin", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("SkinName") = rs("SkinName")
            trs("Skin_CSS") = rs("Skin_CSS")
            trs("IsDefault") = rs("IsDefault")
            trs("ProjectName") = rs("ProjectName")
            trs("IsDefaultInProject") = rs("IsDefaultInProject")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    '导入国家
    Set rs = tconn.Execute(" select * from PE_Country")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Country", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("Country") = rs("Country")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    '导入省份
    Set rs = tconn.Execute(" select * from PE_Province")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Province", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("Province") = rs("Province")
            trs("Country") = rs("Country")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    '导入城市
    Set rs = tconn.Execute(" select * from PE_City")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_City", Conn, 1, 3
    If trs.BOF And trs.EOF Then
        Do While Not rs.EOF
            trs.addnew
            trs("Area") = rs("Area")
            trs("Country") = rs("Country")
            trs("Province") = rs("Province")
            trs("City") = rs("City")
            trs("Area") = rs("Area")
            trs("Postcode") = rs("Postcode")
            trs("AreaCode") = rs("AreaCode")
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    tconn.Close
    Set tconn = Nothing
End Sub

Sub CreatSkinFile()
    If Not fso.FolderExists(Server.MapPath(InstallDir & "Skin")) Then
        fso.CreateFolder (Server.MapPath(InstallDir & "Skin"))
    End If

    Dim rsSkin, sqlSkin, hf
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Server.CreateObject("adodb.recordset")
    rsSkin.Open sqlSkin, Conn, 1, 1
    Do While Not rsSkin.EOF
        Set hf = fso.CreateTextFile(Server.MapPath(InstallDir & "Skin/Skin" & rsSkin("SkinID") & ".css"), True)
        hf.Write Replace_CaseInsensitive(rsSkin("Skin_CSS"), "Skin/", InstallDir & "Skin/")
        hf.Close
        rsSkin.MoveNext
    Loop
    rsSkin.Close
    sqlSkin = "select * from PE_Skin where IsDefault=" & PE_True & ""
    rsSkin.Open sqlSkin, Conn, 1, 1
    If rsSkin.BOF And rsSkin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>没有默认风格！</li>"
    Else
        Set hf = fso.CreateTextFile(Server.MapPath(InstallDir & "Skin/DefaultSkin.css"), True)
        hf.Write Replace_CaseInsensitive(rsSkin("Skin_CSS"), "Skin/", InstallDir & "Skin/")
        hf.Close
    End If
    rsSkin.Close
    Set rsSkin = Nothing
End Sub

Function Replace_CaseInsensitive(expression, find, replacewith)
    regEx.Pattern = find
    Replace_CaseInsensitive = regEx.Replace(expression, replacewith)
End Function

Function IsRadioChecked(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function

Function IsOptionSelected(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function

Sub ClearSiteCache()
    PE_Cache.DelAllCache
End Sub

%>
