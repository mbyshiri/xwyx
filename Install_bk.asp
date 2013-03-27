<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.MD5.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Edition.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
Response.Write "<title>���� SiteWeaver " & SystemEdition & " 6.6�氲װ��</title>" & vbCrLf
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
    Call Check  '������ԱȨ��
Else
    Dim sqlConfig, rsConfig
    sqlConfig = "select * from PE_Config"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn, 1, 3
    If rsConfig.BOF And rsConfig.EOF Then
        Response.Write "��վ�������ݶ�ʧ��ϵͳ�޷��������У�"
    Else
        If action = "" Then
            action = "Step1"
        End If
        Select Case action
        Case "Step1"
            Call Step1  '��վ��Ϣ����1
        Case "Step2"
            Call Step2  '����ģ��
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
    Response.Write "      <td height='22' align='center'><strong>�Ķ����Э��</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='center'><textarea name='License' cols='120' rows='30' id='License' readonly>"
%>
<!--#include file="License.txt"-->
<%
    Response.Write "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='left'><input name='AgreeLicense' type='checkbox' id='AgreeLicense' value='Yes' onclick='document.myform.submit.disabled=!this.checked;'><label for='AgreeLicense'>���Ѿ��Ķ���ͬ���Э��</label></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value=' ��һ�� ' disabled>" & vbCrLf
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
    Response.Write "      <td height='22' colspan='2' align='center'><strong>����Ա��¼</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' class='tdbg5'><strong>�û����ƣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='UserName' type='text' id='UserName' value='' size='30' maxlength='20' style='width:150px;'> Ĭ���û���Ϊ��admin</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' class='tdbg5'><strong>�û����룺</strong></td>" & vbCrLf
    Response.Write "      <td><input name='password' type='password' id='password' value='' size='30' maxlength='20' style='width:150px;'> Ĭ���û�����Ϊ��admin888</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' class='tdbg5'><strong>�� ֤ �룺</strong></td>" & vbCrLf
    Response.Write "      <td valign='top'><input name='CheckCode' type='text' id='CheckCode' value='' size='6' maxlength='6'> <img id='checkcode' src='inc/checkcode.asp' style='border: 1px solid #ffffff'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='ChkLogin'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value=' �� ¼ '>" & vbCrLf
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
        ErrMsg = ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>���벻��Ϊ�գ�</li>"
    End If
    If CheckCode = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��֤�벻��Ϊ�գ�</li>"
    End If
    If Trim(Session("CheckCode")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>���¼ʱ������������·��ص�¼ҳ����е�¼��</li>"
    End If
    If CheckCode <> Session("CheckCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣</li>"
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
        ErrMsg = ErrMsg & "<br><li>�û�����������󣡣���</li>"
    Else
        If Password <> rs("Password") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>�û�����������󣡣���</li>"
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
    Response.Write "      <td height='22' colspan='2' align='center'><strong>���� SiteWeaver " & SystemEdition & " 6.5�氲װ��</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ���ƣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteName' type='text' id='SiteName' value='" & rsConfig("SiteName") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ���⣺</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteTitle' type='text' id='SiteTitle' value='" & rsConfig("SiteTitle") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ��ַ��</strong><br>����д����URL��ַ</td>" & vbCrLf
    Response.Write "      <td><input name='SiteUrl' type='text' id='SiteUrl' value='" & rsConfig("SiteUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><font color=red><strong>��װĿ¼��</strong><br>ϵͳ��װĿ¼������ڸ�Ŀ¼��λ�ã�<br>ϵͳ���Զ������ȷ��·��������Ҫ�ֹ��������á�</font></td>" & vbCrLf
    Response.Write "      <td><input name='InstallDir' type='text' id='InstallDir' value='" & InstallDir & "' size='40' maxlength='30' readonly></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>LOGO��ַ��</strong><br>����д����URL��ַ</td>" & vbCrLf
    Response.Write "      <td><input name='LogoUrl' type='text' id='LogoUrl' value='" & rsConfig("LogoUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>Banner��ַ��</strong><br>����д����URL��ַ</td>" & vbCrLf
    Response.Write "      <td><input name='BannerUrl' type='text' id='BannerUrl' value='" & rsConfig("BannerUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>FSO(FileSystemObject)��������ƣ�</strong><br>ĳЩ��վΪ�˰�ȫ����FSO��������ƽ��и����Դﵽ����FSO��Ŀ�ġ���������վ���������ģ����ڴ�������Ĺ������ơ�</td>" & vbCrLf
    Response.Write "      <td><input name='objName_FSO' type='text' id='objName_FSO' value='" & rsConfig("objName_FSO") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��̨����Ŀ¼��</strong><br>Ϊ�˰�ȫ���������޸ĺ�̨����Ŀ¼��Ĭ��ΪAdmin�����Ĺ��Ժ���Ҫ�����ô˴�</td>" & vbCrLf
    Response.Write "      <td><input name='AdminDir' type='text' id='AdminDir' value='" & rsConfig("AdminDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��վ���Ŀ¼��</strong><br>Ϊ�˲��ù���������������վ�Ĺ�棬�������޸Ĺ��JS�Ĵ��Ŀ¼��Ĭ��ΪAD�����Ĺ��Ժ���Ҫ�����ô˴�</td>" & vbCrLf
    Response.Write "      <td><input name='ADDir' type='text' id='ADDir' value='" & rsConfig("ADDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>վ��������</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterName' type='text' id='WebmasterName' value='" & rsConfig("WebmasterName") & "' size='40' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>վ�����䣺</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterEmail' type='text' id='WebmasterEmail' value='" & rsConfig("WebmasterEmail") & "' size='40' maxlength='100'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>��Ȩ��Ϣ��</strong><br>֧��HTML��ǣ�����ʹ��˫����</td>" & vbCrLf
    Response.Write "      <td><textarea name='Copyright' cols='60' rows='4' id='Copyright'>" & rsConfig("Copyright") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Step2'>" & vbCrLf
    Response.Write "        <input name='submit' type='submit' id='submit' value=' ��һ�� '>" & vbCrLf
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
    Call WriteSuccessMsg("ϵͳ��װ��ɣ����������<a href='Index.asp'>ʹ��ϵͳ</a>�ˡ�<br>Ϊ��<font color='red'>ϵͳ��ȫ</font>����������İ�ťɾ���˰�װ�ļ���Install.asp��<br><br><div align='center'><input name='delfile' type='button' id='delfile' value=' ɾ���˰�װ�ļ� ' onclick=""location='install.asp?Action=Stepdel'""></div><br>", ComeUrl)
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
        ErrMsg = ErrMsg & "<br><li>ɾ���˰�װ�ļ���Install.asp��ʧ�ܣ�����ԭ��" & Err.Description & "<br>���ֶ�ɾ�����ļ���"
        Err.Clear
        Exit Sub
    Else
        Call WriteSuccessMsg("ɾ���˰�װ�ļ���Install.asp���ɹ���<br><br><a href='Index.asp'>��˿�ʼʹ��ϵͳ</a>", ComeUrl)
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
        ErrMsg = ErrMsg & "<br><li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If
    '���뷽��
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
    
    
    '����ģ��
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

    '�����Զ����ǩ
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

    '������
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
    '�������
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
    '����ʡ��
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
    '�������
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
        ErrMsg = ErrMsg & "<br><li>û��Ĭ�Ϸ��</li>"
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
