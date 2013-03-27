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
<style type="text/css">
<!--
.STYLE4 {color: #000000}
-->
</style>
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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">留言板管理</font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">留言板管理</td>
    <td width="200">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_GuestBook.asp?Passed=False" target=main>留言审核</A></td>
    <td width="400">&nbsp;</td>
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
    <td width="300">　　留言板是网站交互的一种形式，支持<a href="#" title="UBB 代码是 Infopop 公司为其 Ultimate Bulletin Board 论坛制作的专用 HTML 代码。"><u>UBB代码</u></a>，具备头像与表情功能。用户可以在网站中留言，也可查看、发布和回复留言信息；管理员可审核、回复、修改、删除留言信息；管理员可以在后台设置是否开启留言审核功能。<br>
      　　留言模式有二种：<a href="#" title="指不需要注册成为本网站的注册用户就可以查看、发布和回复信息"><u>游客模式</u></a>和<a href="#" title="指注册成为本网站的注册用户后进行查看、发布和回复信息，有查看自己的留言功能。"><u>用户模式</u></a>。</td>
    <td width="40">&nbsp;</td>
    <td width="500" valign="top">　　为防止用户在留言时发表不良的言论，可开启系统的留言审核功能。开启了审核功能后，用户的留言需通过管理员审核后才能在前台显示。启用留言审核功能请在<a href="Admin_Channel.asp?Action=Modify&iChannelID=4"><u>留言频道管理</u></a>中设置。<br>
      　　管理员可以对用户留言进行修改、删除、回复和通过审核的操作。<br>
      　　快捷菜单：<A href="Admin_GuestBook.asp?Passed=False" target=main><font color="#FF0000"><U>网站留言审核</U></font></A>。</td>
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
    <td width="100" align="center">&nbsp;</td>
    <td width="200">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_GuestBook.asp?Passed=All" target=main>留言管理</A></td>
    <td width="400">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="1" colspan="2"></td>
    <td colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="300" valign="top">&nbsp;</td>
    <td width="40">&nbsp;</td>
    <td width="500" height="100" valign="top">　　管理员可以对用户留言进行修改、删除、回复和取消审核的操作。取消审核后的留言不会在前台显示。<br>
    　　快捷菜单：<A href="Admin_GuestBook.asp?Passed=All" target=main><font color="#FF0000"><U>网站留言管理</U></font></A>。</td>
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
