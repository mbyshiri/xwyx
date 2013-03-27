<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Call CloseConn
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>会员登录</title>
<link href="../images/style.CSS" rel="stylesheet" type="text/css">
<script language=javascript>
function refreshimg(){
  document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();
}
function SetFocus()
{
if (document.Login.UserName.value=="")
    document.Login.UserName.focus();
else
    document.Login.UserName.select();
}
function CheckForm()
{
    if(document.Login.UserName.value=="")
    {
        alert("请输入用户名！");
        document.Login.UserName.focus();
        return false;
    }
    if(document.Login.UserPassword.value == "")
    {
        alert("请输入密码！");
        document.Login.UserPassword.focus();
        return false;
    }
}
</script>
<style type="text/css">
<!--
body {margin-left: 0px;margin-top: 0px;margin-right: 0px;margin-bottom: 0px;}
-->
</style>
</head>
<body>
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
        <form name="Login" action="User_ChkLogin.asp" method="post" onSubmit="return CheckForm();">
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="120" height="164" background="images/User_Login_0_02.gif"></td>
              <td width="60" height="164" background="images/User_Login_0_04.gif"></td>
              <td valign="top" background="images/User_Login_0_08.gif"><table border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="220" height="79" background="images/User_Login_0_05.gif"></td>
                  <td width="279"><table width="100%" height="79"  border="0" cellpadding="0" cellspacing="0">
                    <tr></tr>
                    <tr>
                      <td ><font color="#ffffff">欢迎您登录<%=SiteName%>！<br>
                        如果您尚未注册，请先<a href="../Reg/User_Reg.asp"><font color="#FFFF00">注册</font></a>。</font></td>
                      <td width="85" valign="bottom" ><input name="ComeUrl" type="hidden" id="ComeUrl" value="<%=ComeUrl%>">
                          <input type="image" name="Submit" src="Images/User_Login_0_13.gif" style="width:85px; HEIGHT: 57px;"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="85" colspan="2"><table border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="35" rowspan="2"><img src="images/User_Login_0_15.gif" width="20" height="30" alt=""></td>
                      <td height="20"><font color="#ffffff">用户名称：</font></td>
                      <td width="45" rowspan="2" align="center" valign="middle"><img src="images/User_Login_0_19.gif" width="20" height="30" alt=""></td>
                      <td><font color="#ffffff">用户密码：</font></td>
<%If EnableCheckCodeOfLogin = True then%>
                      <td width="50" rowspan="2" align="center"><img src="images/User_Login_0_23.gif" width="29" height="30" alt=""></td>
                      <td><font color="#ffffff">验证码：</font></td>
                      <td>&nbsp;</td>
<%End If%>
                      <td width="35" rowspan="2" align="center"><img src="Images/imagesUser_Login_Cookie.gif" alt=""></td>
                      <td><font color="#ffffff">Cookie：</font></td>
                    </tr>
                    <tr>
                      <td><input name="UserName"  type="text"  id="UserName" maxlength="20" style="width:70px; BORDER-RIGHT: #ffffff 0px solid; BORDER-TOP: #ffffff 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #ffffff 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #ffffff"></td>
                      <td><input name="UserPassword"  type="password" id="UserPassword" maxlength="20" style="width:70px; BORDER-RIGHT: #ffffff 0px solid; BORDER-TOP: #ffffff 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #ffffff 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #ffffff"></td>
<%If EnableCheckCodeOfLogin = True then%>
                      <td><input name='CheckCode' size='6' maxlength='6' style='width:50px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7; ime-mode:disabled;' onmouseover=''this.style.background='#ffffff';'' onmouseout=''this.style.background='#F7F7F7''' onFocus='this.select();'></td>
                      <td>&nbsp;<a href='javascript:refreshimg()' title='看不清楚，换个图片'><img id='checkcode' src='../Inc/CheckCode.asp' style='border: 1px solid #ffffff' /></a></td>
<%End If%>
                      <td width="40"><select name='CookieDate'  style='border: 1px solid #ffffff'>
                          <option selected value='0'>不保存</option>
                          <option value='1'>保存一天</option>
                          <option value=2>保存一月</option>
                          <option value=3>保存一年</option>
                      </select></td>
                    </tr>
                  </table></td>
                </tr>
                
              </table></td>
            </tr>
          </table>
          </form>
        <script language="JavaScript" type="text/JavaScript">
        SetFocus();
        </script> </td>
  </tr>
</table>
</body>
</html>