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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">用户管理</font></td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_User.asp" target=main>注册用户管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Admin.asp" target=main>管理员管理</A></td>
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
    <td width="400">　　本功能可以详细管理与设置网站注册用户的信息与权限。可以对用户进行修改、锁定、删除、续费的操作，也可以对用户进行删除、锁定和解锁的操作，并可移动用户到相应的<a href="Admin_UserGroup.asp" target="main" title="用户组是用户账户的集合，通过创建用户组，赋予相关用户享有授予组的权力和权限。具体的权限设置在“频道管理”及各频道的“栏目管理”中。"><U>用户组</U></a>。<br>
      　　快捷菜单：<A href="Admin_User.asp" target=main><font color="#FF0000"><u>注册用户管理</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　系统具有强大的网站权限管理，可设置管理员详细权限，如增删管理员和指定详细的管理权限，使网站的管理分级分类多人共同管理。设置网站<a href="Admin_Admin.asp?Action=Add" target="main" title="超级管理员：拥有所有权限。某些权限（如管理员管理、网站信息配置、网站选项配置等管理权限）只有超级管理员才有。"><U>超级管理员</U></a>和<a href="Admin_Admin.asp?Action=Add" target="main" title="普通管理员：捅有指定部分网站管理功能，需要详细指定每一项管理权限。"><U>普通管理员</U></a>，同一账号可设置是否允许多人同时使用此帐号登录。<br>
      　　快捷菜单：<A href="Admin_Admin.asp?Action=Add" target=main><font color="#FF0000"><u>管理员添加</u></font></A> | <A href="Admin_Admin.asp" target=main><font color="#FF0000"><u>管理</u></font></A>。</td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_UserGroup.asp" target=main>用户组管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Maillist.asp" target=main>邮件列表管理</A></td>
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
    <td width="400" valign="top">　　用户组是用户账户的集合，通过创建用户组，赋予相关用户享有授予组的权力和权限。用户组权限的数字越小，说明具有的权限越大（等级越高）。权限设置采用等级制，即高等级的用户会具有低等级用户的所有权限。具体的使用权限设置在“频道管理”及各频道的“栏目管理”中。<br>
      　　快捷菜单：<A href="Admin_UserGroup.asp" target=main><font color="#FF0000"><u>用户组管理</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　按用户类型、按用户姓名和按用户Email发送邮件。信息将发送到所有注册时完整填写了信箱的用户，邮件列表的使用将消耗大量的服务器资源，请慎重使用。导出功能可将邮件列表批量到数据库或文本。<br>
      　　快捷菜单：<A href="Admin_Maillist.asp" target=main><font color="#FF0000"><u>邮件列表</u></font></A> | <A href="Admin_Maillist.asp?Action=Export" target=main><font color="#FF0000"><u>列表导出</u></font></A><font color="#FF0000">&nbsp;</font>。</td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_User.asp?Action=Update" target=main>更新用户数据</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">管理短消息</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="1" colspan="2" class="topbg2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">　　本操作将重新计算用户的发表文章数。本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。修复起始ID号到结束ID之间的用户数据，之间的数值最好不要选择过大。<br>
      　　快捷菜单：<A href="Admin_User.asp?Action=Update" target=main><font color="#FF0000"><u>更新用户数据</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　系统提供了短消息功能，您也可以撰写短消息，与本站内的注册用户进行交流。请输入<a href="#" title="收件人只能输入本站注册用户的注册名。收件人可以用英文状态下的逗号将用户名隔开实现群发，最多5个用户。"><u>收件人</u></a>、<a href="#" title="最多50个字符"><u>标题</u></a>、<a href="#" title="最多1000个字符"><u>内容</u></a>。您可以管理短消息，随时查看自己的发件箱，删除过期的短消息以节省服务器的空间。<br>
    <%if AdminPurview=1 or CheckPurview_Other(AdminPurview_Others,"Message")=True then%>
　　快捷菜单：<A href="Admin_Message.asp" target=main><font color="#FF0000"><u>管理短消息</u></font></A>。</td>
    <%end if%>
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
