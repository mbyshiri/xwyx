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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">网站生成管理</font></td>
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
    <td width="100" align="center" class="topbg">网站生成管理</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">启用生成功能</td>
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
    <td width="400">　　系统具有强大的生成HTML功能。可将首页、栏目页、内容页、专题页……所有页面都可以生成完全的HTML页面（评论和点击数统计除外），以加快网页的访问速度，减轻服务器负担。</td>
    <td width="40">&nbsp;</td>
    <td width="400"> 　　系统具有独创的每个频道都可选择使用“生成HTML”功能，或选择普通ASP程序显示方式。要使频道具有“生成HTML”功能，请依次点击左栏管理导航中的[<A href="Admin_Channel.asp" target=main><font color="#FF0000">网站频道管理</font></A>]-[修改]-[频道类型]-[<FONT color=red>是否使用生成HTML功能</FONT>]-[是]。可以在此处随时更换网站的HTML显示方式或ASP显示方式。<br>
    </td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_CreateSiteIndex.asp" target="main">首页生成</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center"></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" class="topbg2"></td>
    <td height="1" colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">　　<a href="Admin_CreateSiteIndex.asp" target="main">第一次安装系统请<font color="#FF0000"><u>生成网站首页</u></font>。</a><br>
      　　您也可以随时<A href="Admin_Template.asp?ChannelID=0"><font color="#FF0000"><u>管理首页模板</u></font></A> | <A
href="Admin_Template.asp?ChannelID=0&amp;Action=Add&amp;TemplateType=1"><font color="#FF0000"><u>添加模板</u></font></A> | <A href="Admin_Template.asp?ChannelID=0&amp;Action=Import"><font color="#FF0000"><u>导入模板</u></font></A> | <A
href="Admin_Template.asp?ChannelID=0&amp;Action=Export"><font color="#FF0000"><u>导出模板</u></font></A> <A href="Admin_Template.asp?ChannelID=0" target=main><font color="#FF0000"></font></A> 修改完成后请生成网站首页。</td>
    <td width="40">&nbsp;</td>
    <td width="400">&nbsp;</td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
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
