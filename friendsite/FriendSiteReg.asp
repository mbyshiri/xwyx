<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If EnableLinkReg <> True Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br><li>����Աû�п��������������룡</li>"
Else
    If Action = "Reg" Then
        Call SaveLinkSite
    Else
        Call main
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn

Sub main()
%>
<html>
<head>
<title>������������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
<script language = "JavaScript">
<!--
function CheckForm(){
  if(document.myform.SiteName.value==""){
    alert("��������վ���ƣ�");
    document.myform.SiteName.focus();
    return false;
  }
  if(document.myform.SiteUrl.value=="" || document.myform.SiteUrl.value=="http://"){
    alert("��������վ��ַ��");
    document.myform.SiteUrl.focus();
    return false;
  }
  if(document.myform.SiteAdmin.value==""){
    alert("������վ��������");
    document.myform.SiteAdmin.focus();
    return false;
  }
  if(document.myform.SitePassword.value==""){
    alert("��������վ���룡");
    document.myform.SitePassword.focus();
    return false;
  }
  if(document.myform.SitePwdConfirm.value==""){
    alert("������ȷ�����룡");
    document.myform.SitePwdConfirm.focus();
    return false;
  }
  if(document.myform.SitePwdConfirm.value!=document.myform.SitePassword.value){
    alert("��վ������ȷ�����벻һ�£�");
    document.myform.SitePwdConfirm.focus();
    document.myform.SitePwdConfirm.select();
    return false;
  }
  if(document.myform.SiteIntro.value==""){
    alert("��������վ��飡");
    document.myform.SiteIntro.focus();
    return false;
  }
}
function refreshimg(){document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();}
//-->
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="myform" onSubmit="return CheckForm()"  action="FriendSiteReg.asp">
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="center_tdbgall">
    <tr>
      <td align="center">
        <table width="400" border="0" cellspacing="0" cellpadding="0" class="main_title_575">
          <tr>
            <td><b>��վ������Ϣ</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" width="400" class="main_tdbg_575">
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">��վ���ƣ�</td>
            <td width="307" height="25"><%=SiteName%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վ��ַ��</td>
            <td height="25"><%=SiteUrl%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վLogo��</td>
            <td height="25"><%= GetLogo(88, 31) %></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">վ��������</td>
            <td height="25"><%=WebmasterName%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">�����ʼ���</td>
            <td height="25"><%=WebmasterEmail%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" align="right">��վ��飺</td>
            <td valign="top">���������ӵ�ͬʱ���ñ�վ�����ӡ�</td>
          </tr>
        </table>
        <br>
        <table width="400" border="0" cellspacing="0" cellpadding="0" class="main_title_575">
          <tr>
            <td><b>������������</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" width="400" class="main_tdbg_575">
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">�������</td>
            <td height="25"><select name="KindID" id="KindID">
                <%=GetFsKind_Option(1, 0)%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">����ר�⣺</td>
            <td height="25"><select name="SpecialID" id="SpecialID">
                <%=GetFsKind_Option(2, 0)%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">��վ���ƣ�</td>
            <td height="25"><input name="SiteName" size="30" maxlength="20" title="����������������վ���ƣ����Ϊ20������">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վ��ַ��</td>
            <td height="25"><input name="SiteUrl" size="30" maxlength="100" type="text" value="http://" title="����������������վ��ַ�����Ϊ50���ַ���ǰ������http://">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վLogo��</td>
            <td height="25"><input name="LogoUrl" size="30" maxlength="100" type="text" value="http://" title="����������������վLogoUrl��ַ�����Ϊ50���ַ���������ڵ�һѡ��ѡ������������ӣ�����Ͳ�����"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">վ��������</td>
            <td height="25"><input name="SiteAdmin" size="30" maxlength="20" type="text" title="�������������Ĵ����ˣ���Ȼ��֪������˭�������Ϊ20���ַ�">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">�����ʼ���</td>
            <td height="25"><input name="SiteEmail" size="30" maxlength="30" type="text" value title="����������������ϵ�����ʼ������Ϊ30���ַ�"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">��վ���룺</td>
            <td height="25"><input name="SitePassword" type="password" id="SitePassword" size="20" maxlength="20">
              <font color="#FF0000">*</font> �����޸���Ϣʱ�á�</td>
          </tr>
          <tr class="tdbg">
            <td height="25" align="right">ȷ�����룺</td>
            <td height="25"><input name="SitePwdConfirm" type="password" id="SitePwdConfirm" size="20" maxlength="20">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" align="right">��վ��飺</td>
            <td valign="middle"><textarea name="SiteIntro" cols="40" rows="5" id="SiteIntro" title="����������������վ�ļ򵥽���"></textarea></td>
          </tr>       
<%IF FriendSiteCheckCode = True then%>
          <tr class="tdbg">
          <td vAlign=center align=middle>  ��֤�룺</td>
          <td vAlign=top colSpan=2><a href='javascript:refreshimg()' title='�������������ͼƬ'><img id='checkcode' src='../Inc/CheckCode.asp' style='border: 1px solid #ffffff' /></a>
          <Input maxLength=6 size=10 name=CheckCode><FONT color=red> *</FONT>
          </td></tr> 
<%End If%>
          <tr class="tdbg">
            <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Reg"><input type="submit" value=" ȷ �� " name="cmdOk">
              <input type="reset" value=" �� �� " name="cmdReset"></td>
          </tr>
        </table>
        <br>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
<%
End Sub

Sub SaveLinkSite()
    Dim KindID, SpecialID, LinkType, LinkSiteName, LinkSiteUrl, LinkLogoUrl, LinkSiteAdmin, LinkSiteEmail, LinkSitePassword, LinkSitePwdConfirm, LinkSiteIntro, LinkCheckCode
    KindID = PE_CLng(Trim(request.Form("KindID")))
    SpecialID = PE_CLng(Trim(request.Form("SpecialID")))
    LinkSiteName = Trim(request("SiteName"))
    LinkSiteUrl = Trim(request("SiteUrl"))
    LinkLogoUrl = Trim(request("LogoUrl"))
    LinkSiteAdmin = Trim(request("SiteAdmin"))
    LinkSiteEmail = Trim(request("SiteEmail"))
    LinkSitePassword = Trim(request("SitePassword"))
    LinkSitePwdConfirm = Trim(request("SitePwdConfirm"))
    LinkSiteIntro = Trim(request("SiteIntro"))
    LinkCheckCode = Trim(request("CheckCode"))
    If LinkSiteName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ���Ʋ���Ϊ�գ�</li>"
    End If
    If LinkSiteUrl = "" Or LinkSiteUrl = "http://" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ��ַ����Ϊ�գ�</li>"
    End If
    If LinkSiteAdmin = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>վ����������Ϊ�գ�</li>"
    End If
    If LinkSiteEmail <> "" And IsValidEmail(LinkSiteEmail) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�����ʼ���ַ����</li>"
    End If
    If LinkSitePassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ���벻��Ϊ�գ�</li>"
    End If
    If LinkSitePwdConfirm = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>ȷ�����벻��Ϊ�գ�</li>"
    End If
    If LinkSitePwdConfirm <> LinkSitePassword Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ������ȷ�����벻һ�£�</li>"
    End If
    If LinkSiteIntro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��վ��鲻��Ϊ�գ�</li>"
    End If
    If FriendSiteCheckCode = True then
        If LinkCheckCode = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>��������֤�룡</li>"
        End If
        If FriendSiteCheckCode = True then
            If LinkCheckCode <> Session("CheckCode") Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣</li>"
            End If
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If LinkLogoUrl = "" Or LinkLogoUrl = "http://" Then
        LinkType = 2
    Else
        LinkType = 1
    End If

    Dim sqlLink, rsLink
    LinkSiteName = ReplaceBadChar(LinkSiteName)
    LinkSiteUrl = ReplaceUrlBadChar(LinkSiteUrl)
    sqlLink = "select top 1 * from PE_FriendSite where SiteName='" & LinkSiteName & "' and SiteUrl='" & LinkSiteUrl & "'"
    Set rsLink = Server.CreateObject("Adodb.RecordSet")
    rsLink.open sqlLink, Conn, 1, 3
    If Not (rsLink.bof And rsLink.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>���������վ�Ѿ����ڣ��벻Ҫ�ظ����룡</li>"
    Else
        rsLink.Addnew
        rsLink("KindID") = KindID
        rsLink("SpecialID") = SpecialID
        rsLink("LinkType") = LinkType
        rsLink("SiteName") = LinkSiteName
        rsLink("SiteUrl") = LinkSiteUrl
        rsLink("LogoUrl") = ReplaceUrlBadChar(LinkLogoUrl)
        rsLink("SiteAdmin") = PE_HTMLEncode(LinkSiteAdmin)
        rsLink("SiteEmail") = PE_HTMLEncode(LinkSiteEmail)
        rsLink("SitePassword") = MD5(LinkSitePassword, 16)
        rsLink("SiteIntro") = PE_HTMLEncode(LinkSiteIntro)
        rsLink("Hits") = 0
        rsLink("UpdateTime") = Now
        rsLink("Passed") = False
        rsLink.Update
        Call WriteSuccessMsg("�����������ӳɹ�����ȴ�����Ա���ͨ����", ComeUrl)
    End If
    rsLink.Close
    Set rsLink = Nothing
End Sub

Function GetLogo(LogoWidth, LogoHeight)
    Dim strLogo, strLogoUrl
    If LogoUrl <> "" Then
        If LCase(Left(LogoUrl, 7)) = "http://" Or Left(LogoUrl, 1) = "/" Then
            strLogoUrl = LogoUrl
        Else
            strLogoUrl = InstallDir & LogoUrl
        End If
        If LCase(Right(strLogoUrl, 3)) = "swf" Then
            strLogo = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0'"
            If LogoWidth > 0 Then strLogo = strLogo & " width='" & LogoWidth & "'"
            If LogoHeight > 0 Then strLogo = strLogo & " height='" & LogoHeight & "'"
            strLogo = strLogo & "><param name='movie' value='" & strLogoUrl & "'>"
            strLogo = strLogo & "<param name='wmode' value='transparent'>"
            strLogo = strLogo & "<param name='quality' value='autohigh'>"
            strLogo = strLogo & "<embed"
            If LogoWidth > 0 Then strLogo = strLogo & " width='" & LogoWidth & "'"
            If LogoHeight > 0 Then strLogo = strLogo & " height='" & LogoHeight & "'"
            strLogo = strLogo & " src='" & strLogoUrl & "'"
            strLogo = strLogo & " wmode='transparent'"
            strLogo = strLogo & " quality='autohigh'"
            strLogo = strLogo & "pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>"
            strLogo = strLogo & "</object>"
        Else
            strLogo = "<a href='" & SiteUrl & "' title='" & SiteName & "' target='_blank'>"
            strLogo = strLogo & "<img src='" & strLogoUrl & "'"
            If LogoWidth > 0 Then strLogo = strLogo & " width='" & LogoWidth & "'"
            If LogoHeight > 0 Then strLogo = strLogo & " height='" & LogoHeight & "'"
            strLogo = strLogo & " border='0'>"
            strLogo = strLogo & "</a>"
        End If
    End If
    GetLogo = strLogo
End Function

Function GetFsKind_Option(iKindType, KindID)
    Dim sqlFsKind, rsFsKind, strOption
    strOption = "<option value='0'"
    If KindID = "" Then
        strOption = strOption & " selected"
    End If
    If iKindType = 1 Then
        strOption = strOption & ">�������κ����</option>"
    ElseIf iKindType = 2 Then
        strOption = strOption & ">�������κ�ר��</option>"
    End If
    sqlFsKind = "select * from PE_FsKind"
    If iKindType > 0 Then
        sqlFsKind = sqlFsKind & " where KindType=" & iKindType
    End If
    sqlFsKind = sqlFsKind & " order by KindID"
    Set rsFsKind = Conn.Execute(sqlFsKind)
    Do While Not rsFsKind.EOF
        If rsFsKind("KindID") = KindID Then
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "' selected>" & rsFsKind("KindName") & "</option>"
        Else
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "'>" & rsFsKind("KindName") & "</option>"
        End If
        rsFsKind.movenext
    Loop
    rsFsKind.Close
    Set rsFsKind = Nothing
    GetFsKind_Option = strOption
End Function
%>

