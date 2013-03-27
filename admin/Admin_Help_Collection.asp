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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">采集管理</font></td>
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
    <td width="20">&nbsp;</td>
    <td>　　欢迎您进入<%=SiteName%>采集管理模块！本系统是基于先进的Internet采集技术，具有自主采集网站信息、定制信息采集类别等个性化设置。您可以通过本模块及时采集互联网的信息内容，并将信息存储到本地网站数据库中。您还可以将采集的信息以审核的方式筛选发布。</td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_Collection.asp?Action=Main" target=main>文章采集</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_CollectionHistory.asp?Action=main" target=main>采集历史记录</A></td>
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
    <td width="400" valign="top">　　您可以在这里进行采集管理 ，查看采集的项目名称、采集地址、所属频道、所属栏目、所属专题、状态以及上次采集情况等。<br>
    　　快捷菜单：<A href="Admin_Collection.asp?Action=Main" target=main><font color="#FF0000"><u>文章采集</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　您可以随时查询您进行采集的历史。<br>
　　快捷菜单：<A href="Admin_CollectionHistory.asp?Action=main" target=main><font color="#FF0000"><u>采集历史记录</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_CollectionManage.asp?Action=ItemManage" target=main>项目管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">项目管理</td>
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
    <td width="400" valign="top">　　您可以添加采集项目，按以下步骤完成项目设置：添加项目 &gt;&gt; 基本设置 &gt;&gt; 列表设置 &gt;&gt; 链接设置 &gt;&gt; 正文设置 &gt;&gt; 采样测试 &gt;&gt; 属性设置 &gt;&gt; 完成。<br>
    　　快捷菜单：<A href="Admin_CollectionManage.asp?Action=ItemManage" target=main><font color="#FF0000"><u>项目管理</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　利用本功能，您可以将您的采集数据库利用系统提供的功能进行导入、导出。<br>
    　　快捷菜单：<A href="Admin_CollectionManage.asp?Action=Import" target=main><font color="#FF0000"><u>项目导入</u></font></A> | <A
href="Admin_CollectionManage.asp?Action=Export" target=main><font color="#FF0000"><u>项目导出</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_Filter.asp?Action=main" target=main>过滤管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">&nbsp;</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">　　您可对采集的内容进行过滤管理。您可以添加自定义过滤项目，并设定过滤名称、过滤对象、过滤类型、过滤内容与替换内容，并可以 随时启用或关闭过滤项目。<br>
    　　快捷菜单：<A href="Admin_Filter.asp?Action=main" target=main><font color="#FF0000"><u>过滤管理</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;</td>
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
