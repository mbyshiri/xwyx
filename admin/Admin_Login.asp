<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
'��Ҫ��ʹ������ֵ�ͼƬ�������
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

If Action = "Login" Then
    Call ChkLogin
ElseIf Action = "Logout" Then
    Call Logout
Else
    Call main
End If
If FoundErr = True Then
    Call WriteErrMsg
End If
Call CloseConn

Sub main()
    Response.Write "<html>" & vbCrLf
    Response.Write "<head>" & vbCrLf
    Response.Write "<title>����Ա��¼</title>" & vbCrLf
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link rel='stylesheet' href='Admin_Style.css'>" & vbCrLf
    Response.Write "<script language='JavaScript' src='../images/softkeyboard.js'></script>" & vbCrLf
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "var closestr=0;" & vbCrLf
    Response.Write "function SetFocus() {" & vbCrLf
    Response.Write "if(document.Login.UserName.value == '')" & vbCrLf
    Response.Write "  document.Login.UserName.focus();" & vbCrLf
    Response.Write "else" & vbCrLf
    Response.Write "  document.Login.UserName.select();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm() {" & vbCrLf
    Response.Write "  if(document.Login.UserName.value == '') {" & vbCrLf
    Response.Write "    alert('�������û�����');" & vbCrLf
    Response.Write "    document.Login.UserName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.Login.password.value == '') {" & vbCrLf
    Response.Write "    alert('���������룡');" & vbCrLf
    Response.Write "    document.Login.password.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.Login.CheckCode.value == '') {" & vbCrLf
    Response.Write "    alert ('������������֤�룡');" & vbCrLf
    Response.Write "    document.Login.CheckCode.focus();" & vbCrLf
    Response.Write "    return(false);" & vbCrLf
    Response.Write "  }" & vbCrLf
    If EnableSiteManageCode = True Then
        Response.Write "  if (document.Login.AdminLoginCode.value == '') {" & vbCrLf
        Response.Write "    alert ('���������Ĺ�����֤�룡');" & vbCrLf
        Response.Write "    document.Login.AdminLoginCode.focus();" & vbCrLf
        Response.Write "    return(false);" & vbCrLf
        Response.Write "  }" & vbCrLf
    End If
    Response.Write "}" & vbCrLf
    Response.Write "function CheckBrowser() {" & vbCrLf
    Response.Write "  var app=navigator.appName;" & vbCrLf
    Response.Write "  var verStr=navigator.appVersion;" & vbCrLf
    Response.Write "  if(app.indexOf('Netscape') != -1) {" & vbCrLf
    Response.Write "    alert('����������ʾ��\n    ��ʹ�õ���Netscape��Firefox����������IE����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��');" & vbCrLf
    Response.Write "  } else if(app.indexOf('Microsoft') != -1) {" & vbCrLf
    Response.Write "    if (verStr.indexOf('MSIE 3.0')!=-1 || verStr.indexOf('MSIE 4.0') != -1 || verStr.indexOf('MSIE 5.0') != -1 || verStr.indexOf('MSIE 5.1') != -1)" & vbCrLf
    Response.Write "      alert('����������ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��');" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function refreshimg(){document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "</head>" & vbCrLf
    Response.Write "<body>" & vbCrLf
    
    Response.Write "<table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0'><tr><td>" & vbCrLf
    Response.Write "<form name='Login' action='Admin_Login.asp' method='post' target='_parent' onSubmit='return CheckForm();'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td width='219' height='164' background='images/Admin_Login1_0_02.gif'></td>" & vbCrLf
    Response.Write "      <td width='64' height='164' background='images/Admin_Login1_0_04.gif'></td>" & vbCrLf
    Response.Write "      <td valign='top' background='images/Admin_Login1_0_09.gif'><table border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td><table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td width='270' height='79' background='images/Admin_Login1_0_05.gif'></td>" & vbCrLf
    Response.Write "              <td width='150' height='79' background='images/Admin_Login1_0_06.gif'></td>" & vbCrLf
    Response.Write "              <td valign='top'><table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "                <tr>" & vbCrLf
    Response.Write "                  <td height='21'></td>" & vbCrLf
    Response.Write "                  <td></td>" & vbCrLf
    Response.Write "                </tr>" & vbCrLf
    Response.Write "                <tr>" & vbCrLf
    Response.Write "                  <td><input type='hidden' name='Action' value='Login' /><input type='image' name='Submit' src='Images/Admin_Login1_0_10.gif' style='width:50px; HEIGHT: 50px;' /></td>" & vbCrLf
    Response.Write "                  <td width='58' height='50' background='images/Admin_Login1_0_11.gif'></td>" & vbCrLf
    Response.Write "                </tr>" & vbCrLf
    Response.Write "              </table></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "           </table></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td height='85'><table border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td width='22' rowspan='2' valign='bottom'><img src='images/Admin_Login1_0_15.gif' alt='' /></td>" & vbCrLf
    Response.Write "              <td width='80'><font color='#ffffff'>�û����ƣ�</font></td>" & vbCrLf
    Response.Write "              <td width='22' rowspan='2' valign='bottom'><img src='images/Admin_Login1_0_19.gif' alt='' /></td>" & vbCrLf
    Response.Write "              <td width='80'><font color='#ffffff'>�û����룺</font></td>" & vbCrLf
    If EnableSiteManageCode = True Then
        Response.Write "              <td width='22' rowspan='2' valign='bottom'><img src='images/Admin_Login1_admin.gif' alt='' /></td>" & vbCrLf
        Response.Write "              <td width='80'><font color='#ffffff'>������֤�룺</font></td>" & vbCrLf
    End If
    Response.Write "              <td width='22' rowspan='2' valign='bottom'><img src='images/Admin_Login1_0_23.gif' alt='' /></td>" & vbCrLf
    Response.Write "              <td colspan='2'><font color='#ffffff'>��֤�룺</font></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td><input name='UserName' type='text' id='UserName' maxlength='20' style='width:70px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7' onmouseover=''this.style.background='#ffffff';'' onmouseout=''this.style.background='#F7F7F7''' onFocus='this.select();'></td>" & vbCrLf
    If EnableSoftKey = True Then
        Response.Write "              <td><input name='password'  type='password' maxlength='20' readOnly style='width:70px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7' onmouseover=''this.style.background='#ffffff';'' onmouseout=''this.style.background='#F7F7F7''' onFocus=""password1=this;showkeyboard();Calc.password.value=''""></td>" & vbCrLf
    Else
        Response.Write "              <td><input name='password'  type='password' maxLength='20' style='width:70px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7' onmouseover=''this.style.background='#ffffff';'' onmouseout=''this.style.background='#F7F7F7''' onFocus='this.select();'></td>" & vbCrLf
    End If
    If EnableSiteManageCode = True Then
        Response.Write "              <td><input name='AdminLoginCode'  type='password' maxLength='20' style='width:70px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7' onmouseover=''this.style.background='#ffffff';'' onmouseout=''this.style.background='#F7F7F7''' onFocus='this.select();'></td>" & vbCrLf
    End If
    Response.Write "              <td width='53'><input name='CheckCode' size='6' maxlength='6' style='width:50px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7; ime-mode:disabled;' onmouseover=''this.style.background='#ffffff';'' onmouseout=''this.style.background='#F7F7F7''' onFocus='this.select();'></td>" & vbCrLf
    Response.Write "              <td width='51'><a href='javascript:refreshimg()' title='�������������ͼƬ'><img id='checkcode' src='../Inc/CheckCode.asp' style='border: 1px solid #ffffff' /></a></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table></td>" & vbCrLf
    Response.Write "   </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf

    If EnableSiteManageCode = True And SiteManageCode = "PowerEasy2008" Then
        Response.Write "      <br><center><font color=""red"">��ʹ�õĺ�̨������֤��Ϊϵͳ��ʼֵ��PowerEasy2008�������޸�Config.asp�ļ�����Ӧ��SiteManageCodeֵ��</font></center>" & vbCrLf
    End If


    Response.Write "</form>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "CheckBrowser();" & vbCrLf
    Response.Write "SetFocus();" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "</body>" & vbCrLf
    Response.Write "</html>" & vbCrLf
End Sub

Sub ChkLogin()
    Dim sql, rs
    Dim UserName, Password, CheckCode, RndPassword, AdminLoginCode
    UserName = ReplaceBadChar(Trim(Request("UserName")))
    Password = ReplaceBadChar(Trim(Request("Password")))
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
    AdminLoginCode = Trim(Request("AdminLoginCode"))

    If CSng(ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion) < 5.6 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ű��������棨VBScript���汾���ͣ�����ϵ���Ŀռ��̻����������Ա���¡�</li>"
        ErrMsg = ErrMsg & "<li><a href='http://www.microsoft.com/downloads/release.asp?ReleaseID=33136' target='_blank'><font color='green'>�ű������������ص�ַ</font></a></li>"
    End If

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
        ErrMsg = ErrMsg & "<br><li>���ڹ����¼ͣ����ʱ�������������֤��ʧЧ�������·��ص�¼ҳ����е�¼��</li>"
    End If
    If CheckCode <> Session("CheckCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>���������֤���ϵͳ�����Ĳ�һ�£����������롣</li>"
    End If
    If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ĺ�̨������֤�벻�ԣ����������롣</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))

    Password = MD5(Password, 16)
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_Admin where Password='" & Password & "' and AdminName='" & UserName & "'"
    rs.Open sql, Conn, 1, 3
    If rs.bof And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�û�����������󣡣���</li>"
    Else
        If Password <> rs("Password") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>�û�����������󣡣���</li>"
        End If
    End If
    If FoundErr = True Then
        Call InsertLog(1, -1, UserName, UserTrueIP, "��¼ʧ��", ComeUrl, "")
        Session("AdminName") = ""
        Session("AdminPassword") = ""
        Session("RndPassword") = ""
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    UserName = rs("UserName")
    RndPassword = GetRndPassword(16)
    rs("LastLoginIP") = UserTrueIP
    rs("LastLoginTime") = Now()
    rs("LoginTimes") = rs("LoginTimes") + 1
    rs("RndPassword") = RndPassword
    rs.Update
    Call InsertLog(1, 0, UserName, UserTrueIP, "��¼�ɹ�", ComeUrl, "")

    InstallDir = GetInstallDir(Trim(Request.ServerVariables("SCRIPT_NAME")), 1)
    Site_Sn = Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & InstallDir), "/", ""), ".", "")
    Response.Cookies(Site_Sn)("AdminName") = rs("AdminName")
    Response.Cookies(Site_Sn)("AdminPassword") = rs("Password")
    Response.Cookies(Site_Sn)("RndPassword") = RndPassword
    Response.Cookies(Site_Sn)("AdminLoginCode") = AdminLoginCode
    rs.Close

    sql = "select UserID,UserPassword,LastPassword,LastLoginIP,LastLoginTime,LoginTimes from PE_User where UserName='" & UserName & "'"
    rs.Open sql, Conn, 1, 3
    If Not (rs.bof And rs.EOF) Then
        rs("LastPassword") = RndPassword
        rs("LastLoginIP") = UserTrueIP
        rs("LastLoginTime") = Now()
        rs("LoginTimes") = rs("LoginTimes") + 1
        rs.Update
        Response.Cookies(Site_Sn)("UserName") = UserName
        Response.Cookies(Site_Sn)("UserPassword") = rs("UserPassword")
        Response.Cookies(Site_Sn)("LastPassword") = RndPassword
        Session("UserID") = rs("UserID")
    End If
    rs.Close
    Set rs = Nothing

    Call CloseConn
    Response.Redirect "Admin_Index.asp"
End Sub

Sub Logout()
    Conn.Execute ("update PE_Admin set LastLogoutTime=" & PE_Now & "  where AdminName='" & ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName"))) & "'")
    Response.Cookies(Site_Sn)("AdminName") = ""
    Response.Cookies(Site_Sn)("AdminPassword") = ""
    Response.Cookies(Site_Sn)("RndPassword") = ""
    Response.Cookies(Site_Sn)("UserName") = ""
    Response.Cookies(Site_Sn)("UserPassword") = ""
    Response.Cookies(Site_Sn)("LastPassword") = ""
    Response.Cookies(Site_Sn)("UnreadMsg") = ""
    Call CloseConn
    Response.Redirect "../Index.asp"
End Sub

'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
Sub WriteErrMsg()
    Response.Write "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center'><td height='22' class='title'><strong>������Ϣ</strong></td></tr>" & vbCrLf
    Response.Write "  <tr><td height='100' class='tdbg' valign='top'><b>��������Ŀ���ԭ��</b><br>" & ErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center'><td class='tdbg'><a href='Admin_Login.asp'>&lt;&lt; ���ص�¼ҳ��</a></td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub

Sub InsertLog(LogType, ChannelID, UserName, UserIP, LogContent, ScriptName, PostString)
    Dim sqlLog, rsLog
    sqlLog = "select top 1 * from PE_Log"
    Set rsLog = Server.CreateObject("Adodb.RecordSet")
    rsLog.Open sqlLog, Conn, 1, 3
    rsLog.addnew
    rsLog("LogType") = LogType
    rsLog("ChannelID") = ChannelID
    rsLog("LogTime") = Now()
    rsLog("UserName") = UserName
    rsLog("UserIP") = UserIP
    rsLog("LogContent") = LogContent
    rsLog("ScriptName") = ScriptName
    rsLog("PostString") = PostString
    rsLog.Update
    rsLog.Close
    Set rsLog = Nothing
End Sub

'**************************************************
'��������GetInstallDir
'��  �ã�����ǵ�ǰҳ���ڹ����̨�����û���̨,��ȡ����һ����Ŀ¼Ϊϵͳ��װ·��,�����ǰҳ���ڸ�Ŀ¼��,��ȡ��ǰ·��
'��  ����ScriptName ----·������
'        ParentLevel ---- 1 ϵͳ��װ·��,0 ��ǰ·��
'����ֵ������·��
'**************************************************
Function GetInstallDir(ByVal ScriptName, ParentLevel)
    Dim i, strTemp
    GetInstallDir = "/"
    If ScriptName = "" Or IsNull(ScriptName) Then Exit Function
    If ParentLevel > 1 Then ParentLevel = 1
    If ParentLevel = 0 Then
        strTemp = Left(ScriptName, InStrRev(ScriptName, "/"))
    ElseIf ParentLevel = 1 Then
        i = InStrRev(ScriptName, "/") - 1
        If i < 1 Then i = 1
        strTemp = Left(ScriptName, InStrRev(ScriptName, "/", i))
    End If
    If Right(strTemp, 1) <> "/" Then strTemp = strTemp & "/"
    GetInstallDir = strTemp
End Function
%>
