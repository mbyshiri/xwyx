<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

If EnableLinkReg <> True Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br><li>管理员没有开放友情链接申请！</li>"
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
<title>申请友情链接</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
<script language = "JavaScript">
<!--
function CheckForm(){
  if(document.myform.SiteName.value==""){
    alert("请输入网站名称！");
    document.myform.SiteName.focus();
    return false;
  }
  if(document.myform.SiteUrl.value=="" || document.myform.SiteUrl.value=="http://"){
    alert("请输入网站地址！");
    document.myform.SiteUrl.focus();
    return false;
  }
  if(document.myform.SiteAdmin.value==""){
    alert("请输入站长姓名！");
    document.myform.SiteAdmin.focus();
    return false;
  }
  if(document.myform.SitePassword.value==""){
    alert("请输入网站密码！");
    document.myform.SitePassword.focus();
    return false;
  }
  if(document.myform.SitePwdConfirm.value==""){
    alert("请输入确认密码！");
    document.myform.SitePwdConfirm.focus();
    return false;
  }
  if(document.myform.SitePwdConfirm.value!=document.myform.SitePassword.value){
    alert("网站密码与确认密码不一致！");
    document.myform.SitePwdConfirm.focus();
    document.myform.SitePwdConfirm.select();
    return false;
  }
  if(document.myform.SiteIntro.value==""){
    alert("请输入网站简介！");
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
            <td><b>本站链接信息</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" width="400" class="main_tdbg_575">
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">本站名称：</td>
            <td width="307" height="25"><%=SiteName%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">本站地址：</td>
            <td height="25"><%=SiteUrl%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">本站Logo：</td>
            <td height="25"><%= GetLogo(88, 31) %></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">站长姓名：</td>
            <td height="25"><%=WebmasterName%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">电子邮件：</td>
            <td height="25"><%=WebmasterEmail%></td>
          </tr>
          <tr class="tdbg">
            <td width="82" align="right">本站简介：</td>
            <td valign="top">请申请链接的同时做好本站的链接。</td>
          </tr>
        </table>
        <br>
        <table width="400" border="0" cellspacing="0" cellpadding="0" class="main_title_575">
          <tr>
            <td><b>申请友情链接</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" width="400" class="main_tdbg_575">
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">所属类别：</td>
            <td height="25"><select name="KindID" id="KindID">
                <%=GetFsKind_Option(1, 0)%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">所属专题：</td>
            <td height="25"><select name="SpecialID" id="SpecialID">
                <%=GetFsKind_Option(2, 0)%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">网站名称：</td>
            <td height="25"><input name="SiteName" size="30" maxlength="20" title="这里请输入您的网站名称，最多为20个汉字">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">网站地址：</td>
            <td height="25"><input name="SiteUrl" size="30" maxlength="100" type="text" value="http://" title="这里请输入您的网站地址，最多为50个字符，前面必须带http://">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">网站Logo：</td>
            <td height="25"><input name="LogoUrl" size="30" maxlength="100" type="text" value="http://" title="这里请输入您的网站LogoUrl地址，最多为50个字符，如果您在第一选项选择的是文字链接，这项就不必填"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">站长姓名：</td>
            <td height="25"><input name="SiteAdmin" size="30" maxlength="20" type="text" title="这里请输入您的大名了，不然我知道您是谁啊。最多为20个字符">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">电子邮件：</td>
            <td height="25"><input name="SiteEmail" size="30" maxlength="30" type="text" value title="这里请输入您的联系电子邮件，最多为30个字符"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">网站密码：</td>
            <td height="25"><input name="SitePassword" type="password" id="SitePassword" size="20" maxlength="20">
              <font color="#FF0000">*</font> 用于修改信息时用。</td>
          </tr>
          <tr class="tdbg">
            <td height="25" align="right">确认密码：</td>
            <td height="25"><input name="SitePwdConfirm" type="password" id="SitePwdConfirm" size="20" maxlength="20">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" align="right">网站简介：</td>
            <td valign="middle"><textarea name="SiteIntro" cols="40" rows="5" id="SiteIntro" title="这里请输入您的网站的简单介绍"></textarea></td>
          </tr>       
<%IF FriendSiteCheckCode = True then%>
          <tr class="tdbg">
          <td vAlign=center align=middle>  验证码：</td>
          <td vAlign=top colSpan=2><a href='javascript:refreshimg()' title='看不清楚，换个图片'><img id='checkcode' src='../Inc/CheckCode.asp' style='border: 1px solid #ffffff' /></a>
          <Input maxLength=6 size=10 name=CheckCode><FONT color=red> *</FONT>
          </td></tr> 
<%End If%>
          <tr class="tdbg">
            <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Reg"><input type="submit" value=" 确 定 " name="cmdOk">
              <input type="reset" value=" 重 填 " name="cmdReset"></td>
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
        ErrMsg = ErrMsg & "<br><li>网站名称不能为空！</li>"
    End If
    If LinkSiteUrl = "" Or LinkSiteUrl = "http://" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>网站地址不能为空！</li>"
    End If
    If LinkSiteAdmin = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>站长姓名不能为空！</li>"
    End If
    If LinkSiteEmail <> "" And IsValidEmail(LinkSiteEmail) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>电子邮件地址错误！</li>"
    End If
    If LinkSitePassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>网站密码不能为空！</li>"
    End If
    If LinkSitePwdConfirm = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>确认密码不能为空！</li>"
    End If
    If LinkSitePwdConfirm <> LinkSitePassword Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>网站密码与确认密码不一致！</li>"
    End If
    If LinkSiteIntro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>网站简介不能为空！</li>"
    End If
    If FriendSiteCheckCode = True then
        If LinkCheckCode = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>请输入验证码！</li>"
        End If
        If FriendSiteCheckCode = True then
            If LinkCheckCode <> Session("CheckCode") Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>您输入的确认码和系统产生的不一致，请重新输入。</li>"
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
        ErrMsg = ErrMsg & "<br><li>你申请的网站已经存在！请不要重复申请！</li>"
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
        Call WriteSuccessMsg("申请友情链接成功！请等待管理员审核通过。", ComeUrl)
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
        strOption = strOption & ">不属于任何类别</option>"
    ElseIf iKindType = 2 Then
        strOption = strOption & ">不属于任何专题</option>"
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

