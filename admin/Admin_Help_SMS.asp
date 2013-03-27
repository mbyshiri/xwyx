<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 0      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限    
Call CloseConn
%>
<html>
<head>
<title><%=SiteName & "--后台管理首页"%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="392" rowspan="2"><img src="Images/adminmain01.gif" width="392" height="126"></td>
    <td height="114" valign="top" background="Images/adminmain0line2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20"></td>
      </tr>
      <tr>
        <td><%=AdminName%>您好，今天是
          <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">手机短信管理</font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="3">
  <tr>
    <td width="20">&nbsp;</td>
    <td>　　欢迎您使用手机短信功能，本功能与“动易短信通”紧密集成，为您提供了一个高效益、低成本的移动短信商务平台！“动易短信通”是动易公司与中国电信合作的一项新业务，短信通用户可以WEB方式通过“动易短信通”服务平台向中国移动、中国联通、中国电信和中国网通用户实时或定时发送短消息。本业务可广泛应用于商品（订单）处理通知、促销信息发布、咨询服务、会议（紧急）通知、节日祝福、新产品发布、客户沟通等方面，实现移动办公、移动服务！
<br />　　初次使用请先到“动易短信通”服务平台注册会员并充值，然后在本系统中进行“手机短信设置”。</td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="http://sms.powereasy.net/" target="_blank">动易短信通</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_SiteConfig.asp" target=main>手机短信设置</A></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">　　在使用手机短信功能之前，您必须到“动易短信通”服务平台注册，以得到用户名、密码、身份识别号和MD5私钥等信息，并进行充值。<br>
      　　快捷菜单：<A href="http://sms.powereasy.net/Service.aspx" target="_blank"><font color="#FF0000"><u>服务导航</u></font></A> | <A href="http://sms.powereasy.net/Register.aspx" target="_blank"><font color="#FF0000"><u>注册新会员</u></font></A> | <A href="http://sms.powereasy.net/Member/Recharge.aspx" target="_blank"><font color="#FF0000"><u>短信通充值</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400">　　在“网站信息配置”的“网站选项”中可以设置是否启用“手机短信”功能，在“手机短信设置”中，填写相应的您在“动易短信通”平台中的注册用户名和MD5私钥，并预设相关操作的手机短信内容。<br>
      　　快捷菜单：<A href="Admin_SiteConfig.asp" target="main"><font color="#FF0000"><u>网站信息配置</u></font></A></td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_SMS.asp?SendTo=Member" target="main">发送手机短信</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_SMS.asp?SendTo=Other" target="main"></A>查看发送结果</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">　　您可以给网站中的注册会员、注册的联系人、订单中的收货人和其他人的手机或小灵通发送短信。也可以设置好当客户提交订单时系统自动发送手机短信通知管理员。<br>
      　　快捷菜单：发送给<A href="Admin_SMS.asp?SendTo=Member" target="main"><font color="#FF0000"><u>会员</u></font></A> | <A href="Admin_SMS.asp?SendTo=Contacter" target="main"><font color="#FF0000"><u>联系人</u></font></A> | <A href="Admin_SMS.asp?SendTo=Consignee" target="main"><font color="#FF0000"><u>订单中的收货人</u></font></A> | <A href="Admin_SMS.asp?SendTo=Other" target="main"><font color="#FF0000"><u>其他人</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　利用“动易短信通”发送的每一条短信无论成功与否都有详细的发送记录，若短信发送不成功不会计费。您可以查看每次发送手机短信的发送结果。<br>
      　　快捷菜单：<A href="Admin_SMSLog.asp" target="main"><font color="#FF0000"><u>查看短信发送结果</u></font></A></td>
    <td width="21">&nbsp;</td>
  </tr>
</table>

<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 class="topbg"><span class="Glow">Copyright 2003-2006 &copy; <%=SiteName%> All Rights Reserved.</span>
  </tr>
</table>
</body>
</html>
