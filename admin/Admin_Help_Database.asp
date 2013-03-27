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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">数据库管理</font></td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Backup" target=main>备份数据库</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=SpaceSize" target=main>系统空间占用</A></td>
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
    <td width="400">　　系统将备份数据库，以备数据库出现问题时能进行恢复操作。请输入要备份数据库相对路径目录，如目录不存在，将自动创建。不用输入备份文件名的扩展名（默认为“.asa”）。如有同名文件，系统将自动覆盖。<br>
      　　快捷菜单：<A href="Admin_Database.asp?Action=Backup" target=main><font color="#FF0000"><U>备份数据库</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　查看网站系统占用空间的情况。您可以查看基本系统、后台管理、系统图片、各频道及其它文件占用空间的情况。<br>
      　　快捷菜单：<A href="Admin_Database.asp?Action=SpaceSize" target=main><font color="#FF0000"><u>系统空间占用</u></font></A>。</td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Restore" target=main>恢复数据库</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Init" target=main>系统初始化</A></td>
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
    <td width="400" valign="top">　　本功能<a href="Admin_CreateSiteIndex.asp" target="main">将</a><a href="Admin_CreateSiteIndex.asp" target="main">从系统备份的数据库</a>中恢复数据库。请注意备份数据库路径相对路径，并输入正确的数据库名。<br>
      　　快捷菜单：<A href="Admin_Database.asp?Action=Restore" target=main><font color="#FF0000"><U>恢复数据库</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　将网站系统进行初始化，将指定数据库内容的数据将会被清空。<FONT color=#0000FF>请慎用此功能，因为一旦清除将无法恢复！</FONT>。<br>
      　　快捷菜单：<A href="Admin_Database.asp?Action=Init" target=main><font color="#FF0000"><U>系统初始化</u></font></A>。</td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Compact" target=main>压缩数据库</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center">&nbsp;</td>
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
    <td width="400">　　在经常作删除数据等操作后，可以使用本功能保证数据库性能最优。压缩前，建议先备份数据库，以免发生意外错误。<br>
    　　快捷菜单：<A href="Admin_Database.asp?Action=Compact" target=main><font color="#FF0000"><U>压缩数据库</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400">&nbsp;</td>
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
