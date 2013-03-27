<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ID
ID = PE_CLng(Trim(request("ID")))
If ID = 0 Then
    Call CloseConn
    response.Redirect "index.asp"
End If
Dim sqlLink, rsLink
sqlLink = "select * from PE_FriendSite where ID=" & ID
Set rsLink = Server.CreateObject("Adodb.RecordSet")
rsLink.open sqlLink, Conn, 1, 3
If rsLink.bof And rsLink.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br><li>找不到站点！</li>"
Else
    If Action = "Del" Then
        Dim OldSitePassword
        OldSitePassword = Trim(request("OldSitePassword"))
        If OldSitePassword = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>原设密码不能为空！</li>"
        End If
        If MD5(OldSitePassword, 16) <> rsLink("SitePassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>你输入的原设密码不对，没有权限删除！</li>"
        End If
        If FoundErr <> True Then
            rsLink.Delete
            rsLink.Update
            rsLink.Close
            Set rsLink = Nothing
            Call ClearSiteCache(0)			
            Call CloseConn
            response.Redirect "index.asp"
        End If
    End If
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
Else
%>
<html>
<head>
<title>删除友情链接</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
<script language = "JavaScript">
<!--
function CheckForm() {
  if (document.myform.OldSitePassword.value=="") {
    alert ("请输入原设密码！")
    document.myform.OldSitePassword.focus()
    return false
  }
}
//-->
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" name="myform" onsubmit="return CheckForm()" action="FriendSiteDel.asp">
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="center_tdbgall">
    <tr>
      <td align="center"><br>
        <table width="400" border="0" cellspacing="0" cellpadding="0" class="main_title_575">
          <tr>
            <td><b>删除友情链接信息</b></td>
          </tr>
        </table>
        <table border="0" cellpadding="2" cellspacing="1" align="center" width="400" class="main_tdbg_575">
          <tr>
            <td width="100" height="25" align="right">链接类型：</td>
            <td height="25">
              <%
              If rsLink("LinkType") = 1 Then
                response.write "Logo链接"
              Else
                response.write "文字链接"
              End If
              %>
            </td>
          </tr>
          <tr class="tdbg">
            <td width="100" height="25" align="right" valign="middle">网站名称：</td>
            <td height="25"><%=rsLink("SiteName")%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" height="25" align="right">网站地址：</td>
            <td height="25"><%=rsLink("SiteUrl")%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" height="25" align="right">网站Logo：</td>
            <td height="25"><%=rsLink("LogoUrl")%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" height="25" align="right">站长姓名：</td>
            <td height="25"><%=rsLink("SiteAdmin")%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" height="25" align="right">电子邮件：</td>
            <td height="25"><%=rsLink("SiteEmail")%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" align="right">网站简介：</td>
            <td valign="middle"><%=rsLink("SiteIntro")%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" height="25" align="right">原设密码：</td>
            <td height="25"><input name="OldSitePassword" type="password" id="OldSitePassword" size="20" maxlength="20"> <font color="#FF0000">* 必须输入</font></td>
          </tr>
          <tr class="tdbg">
            <td height="40" colspan="2" align="center"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"><input name="Action" type="hidden" id="Action" value="Del"><input type="submit" value=" 确 定 " name="cmdOk"></td>
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
End If
rsLink.Close
Set rsLink = Nothing
Call CloseConn
%>
