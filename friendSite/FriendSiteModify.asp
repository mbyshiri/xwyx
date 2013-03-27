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
    If Action = "Modify" Then
        Call SaveModify
    Else
        Call main
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn

Sub main()
    Dim ID, rsLink, sqlLink
    ID = PE_CLng(Trim(request("ID")))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请指定友情链接ID</li>"
        Exit Sub
    End If
    sqlLink = "select * from PE_FriendSite where Passed=" & PE_True & " and ID=" & ID
    Set rsLink = Conn.Execute(sqlLink)
    If rsLink.bof And rsLink.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>找不到友情链接或者友情链接未审核通过！</li>"
        rsLink.Close
        Set rsLink = Nothing
        Exit Sub
    End If
%>
<html>
<head>
<title>修改友情链接</title>
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
  if(document.myform.OldSitePassword.value==""){
    alert("请输入原设密码！");
    document.myform.OldSitePassword.focus();
    return false;
  }
  if(document.myform.SitePwdConfirm.value!=""||document.myform.SitePassword.value!=""){
    if(document.myform.SitePwdConfirm.value!=document.myform.SitePassword.value){
      alert("网站密码与确认密码不一致！");
      document.myform.SitePwdConfirm.focus();
      document.myform.SitePwdConfirm.select();
      return false;
    }
  }
  if(document.myform.SiteIntro.value==""){
    alert("请输入网站简介！");
    document.myform.SiteIntro.focus();
    return false;
  }
}
//-->
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="myform" onsubmit="return CheckForm()" action="FriendSiteModify.asp">
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="center_tdbgall">
    <tr>
      <td align="center"><br>
        <table width="400" border="0" cellspacing="0" cellpadding="0" class="main_title_575">
          <tr>
            <td><b>修改友情链接信息</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" width="400" class="main_tdbg_575">
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">所属类别：</td>
            <td height="25"><select name="KindID" id="KindID">
                <%=GetFsKind_Option(1, rsLink("KindID"))%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">所属专题：</td>
            <td height="25"><select name="SpecialID" id="SpecialID">
                <%=GetFsKind_Option(2, rsLink("SpecialID"))%>
              </select></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right" valign="middle">网站名称：</td>
            <td height="25"><input name="SiteName" size="30" maxlength="20" title="这里请输入您的网站名称，最多为20个汉字" value="<%=rsLink("SiteName")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">网站地址：</td>
            <td height="25"><input name="SiteUrl" size="30" maxlength="100" type="text" value="<%=rsLink("SiteUrl")%>" title="这里请输入您的网站地址，最多为50个字符，前面必须带http://">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">网站Logo：</td>
            <td height="25"><input name="LogoUrl" size="30" maxlength="100" type="text" value="<%=rsLink("LogoUrl")%>" title="这里请输入您的网站LogoUrl地址，最多为50个字符，如果您在第一选项选择的是文字链接，这项就不必填"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">站长姓名：</td>
            <td height="25"><input name="SiteAdmin" size="30" maxlength="20" type="text" title="这里请输入您的大名了，不然我知道您是谁啊。最多为20个字符" value="<%=rsLink("SiteAdmin")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">电子邮件：</td>
            <td height="25"><input name="SiteEmail" size="30" maxlength="30" type="text" value="<%=rsLink("SiteEmail")%>" title="这里请输入您的联系电子邮件，最多为30个字符"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">原设密码：</td>
            <td height="25"><input name="OldSitePassword" type="password" id="OldSitePassword" size="20" maxlength="20">
              <font color="#FF0000">* 必须输入</font></td>
          </tr>
          <tr class="tdbg">
            <td width="82" height="25" align="right">新设密码：</td>
            <td height="25"><input name="SitePassword" type="password" id="SitePassword" size="20" maxlength="20">
              <font color="#0000FF">若不修改，请保持为空</font></td>
          </tr>
          <tr class="tdbg">
            <td height="25" align="right">确认密码：</td>
            <td height="25"><input name="SitePwdConfirm" type="password" id="SitePwdConfirm" size="20" maxlength="20"></td>
          </tr>
          <tr class="tdbg">
            <td width="82" align="right">网站简介：</td>
            <td valign="middle"><textarea name="SiteIntro" cols="40" rows="5" id="SiteIntro" title="这里请输入您的网站的简单介绍"><%=PE_ConvertBR(rsLink("SiteIntro"))%></textarea></td>
          </tr>
          <tr class="tdbg">
            <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Modify"><input type="submit" value=" 确 定 " name="cmdOk">
              <input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"></td>
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
    rsLink.Close
    Set rsLink = Nothing
End Sub

Sub SaveModify()
    Dim ID, KindID, SpecialID, LinkType, LinkSiteName, LinkSiteUrl, LinkLogoUrl, LinkSiteAdmin, LinkSiteEmail, OldSitePassword, LinkSitePassword, LinkSitePwdConfirm, LinkSiteIntro
    ID = PE_CLng(Trim(request.Form("ID")))
    KindID = PE_CLng(Trim(request.Form("KindID")))
    SpecialID = PE_CLng(Trim(request.Form("SpecialID")))
    LinkSiteName = Trim(request("SiteName"))
    LinkSiteUrl = Trim(request("SiteUrl"))
    LinkLogoUrl = Trim(request("LogoUrl"))
    LinkSiteAdmin = Trim(request("SiteAdmin"))
    LinkSiteEmail = Trim(request("SiteEmail"))
    OldSitePassword = Trim(request("OldSitePassword"))
    LinkSitePassword = Trim(request("SitePassword"))
    LinkSitePwdConfirm = Trim(request("SitePwdConfirm"))
    LinkSiteIntro = Trim(request("SiteIntro"))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>不能确定友情链接ID</li>"
    End If
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
    If OldSitePassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>网站原设密码不能为空！</li>"
    End If
    If LinkSitePwdConfirm <> LinkSitePassword Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>新网站密码与确认密码不一致！</li>"
    End If
    If LinkSiteIntro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>网站简介不能为空！</li>"
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
    sqlLink = "select * from PE_FriendSite where ID=" & ID
    Set rsLink = Server.CreateObject("Adodb.RecordSet")
    rsLink.open sqlLink, Conn, 1, 3
    If rsLink.bof And rsLink.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>找不到指定的友情链接！</li>"
    Else
        If MD5(OldSitePassword, 16) <> rsLink("SitePassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>你输入的旧网站密码不对，没有权限修改！</li>"
            rsLink.Close
            Set rsLink = Nothing
            Exit Sub
        End If
        rsLink("KindID") = KindID
        rsLink("SpecialID") = SpecialID
        rsLink("LinkType") = LinkType
        rsLink("SiteName") = ReplaceBadChar(LinkSiteName)
        rsLink("SiteUrl") = ReplaceUrlBadChar(LinkSiteUrl)
        rsLink("LogoUrl") = ReplaceUrlBadChar(LinkLogoUrl)
        rsLink("SiteAdmin") = PE_HTMLEncode(LinkSiteAdmin)
        rsLink("SiteEmail") = PE_HTMLEncode(LinkSiteEmail)
        If LinkSitePassword <> "" Then
            rsLink("SitePassword") = MD5(LinkSitePassword, 16)
        End If
        rsLink("SiteIntro") = PE_HTMLEncode(LinkSiteIntro)
        rsLink("UpdateTime") = Now
        rsLink("Passed") = False
        rsLink.Update
        Call WriteSuccessMsg("修改友情链接成功！请等待管理员审核通过。", ComeUrl)
    End If
    rsLink.Close
    Set rsLink = Nothing
End Sub

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
