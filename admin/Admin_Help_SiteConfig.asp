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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">系统设置管理</font></td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_SiteConfig.asp">网站信息配置</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Template.asp?ChannelID=0" target="main">首页模板管理</a></td>
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
    <td width="400">　　网站基本信息配置：如网站的名称、地址、LOGO、版权信息等；网站功能选项配置：如显示频道、保存远程图片等；用户选项配置：如是否允许新用户注册、是否需要认证等；另有邮件服务器选项等。<br>
    　　快捷菜单：<a href="Admin_SiteConfig.asp" target="main"><u><font color="#FF0000">网站信息配置</font></u></a> | <a href="Admin_Article.asp?ChannelID=1&amp;Action=Manage&amp;Passed=True&amp;ManageType=HTML" target=main><font color="#FF0000"><u></u></font></a><a href="Admin_SiteConfig.asp#SiteOption" target="main"><u><font color="#FF0000">网站选项配置</font></u></a> | <a href="Admin_Article.asp?ChannelID=1&amp;Action=Manage&amp;Passed=True&amp;ManageType=HTML" target=main><font color="#FF0000"><u></u></font></a><a href="Admin_SiteConfig.asp#User" target="main"><u><font color="#FF0000">用户选项</font></u></a>。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　第一次安装系统请<font color="#FF0000">生成网站首页</font>。首页、栏目页、内容页、专题页……都可以生成完全的HTML页面（评论和点击数统计除外）。各频道启用生成功能请在<a href=Admin_Channel.asp target=main title="动易网站管理系统中，频道是指某一功能模板的集合。某一频道可以是具备文章系统功能，或具备下载系统、图片系统的功能。"><U>网站频道管理</U></a>中设置。<br>
      　　快捷菜单：<a href="Admin_Template.asp?ChannelID=0"><font color="#FF0000"><u>管理网站首页模板</u></font></a> | <a href="Admin_CreateSiteIndex.asp"><font color="#FF0000"><u>生成网站首页</u></font></a>。</td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Channel.asp" target="main">网站频道管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Skin.asp" target="main">网站风格管理</a></td>
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
    <td width="400">　　管理网站的各个频道的功能模块，如文章、下载、图片和留言等频道。频道可分为<a href="#" title="系统内部频道指的是在MY动力现有功能模块（新闻、文章、图片等）基础上添加新的频道，新频道具备和所使用功能模块完全相同的功能。"><U>系统内部频道</U></a>与<a href="#" title="外部频道指链接到MY动力系统以外的地址中。当此频道准备链接到网站中的其他系统时"><U>外部频道</U></a>二类。系统的一些重要功能，如生成HTML功能、频道的审核功能、上传文件类型、顶部导航栏每行显示的栏目数、底部栏目导航的显示方式、都在此进行设置。<br>
    　　快捷菜单：<A href="Admin_Channel.asp?Action=Add" target=main><font color="#FF0000"><u>添加网站频道</u></font></A> | <A href="Admin_Channel.asp" target=main><font color="#FF0000"><u>管理网站频道</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　风格模板是控制整个网站在前台显示时看到的的字体、风格、图片等，通常是用css网页样式语句来进行设计和控制的。利用网页技术中的层叠样式表(CSS)样式来定义特定的HTML标签以按照特定方式设置文本格式。系统具有自定义CSS样式的功能，并随时可以修改样式。<br>
      　　快捷菜单：<A href="Admin_Skin.asp?Action=Add" target=main><font color="#FF0000"><u>添加网站风格</u></font></A> | <A href="Admin_Skin.asp" target=main><font color="#FF0000"><u>管理网站风格</u></font></A>。</td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Announce.asp" target="main">网站公告管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Vote.asp" target="main">网站调查管理</a></td>
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
    <td width="400">　　可以发布、修改和删除网站公告。可以设置为频道共用公告，也可以发布各频道不同的公告。公告支持<a href="#" title="UBB 代码是 Infopop 公司为其 Ultimate Bulletin Board 论坛制作的专用 HTML 代码。"><u>UBB代码</u></a>。有全部（滚动及弹出）、滚动和弹出三种显示类型，只有将公告设为最新时才会在前台显示。<br>
      　　快捷菜单：<A href="Admin_Announce.asp?Action=Add" target=main><u><font color="#FF0000">发布网站公告</font></u></A> | <A href="Admin_Announce.asp" target=main><font color="#FF0000"><u>管理网站公告</u></font></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　可以发布、修改和删除网站调查。可以设置为频道共用调查，也可以发布各频道不同的调查。可以发布单主题调查，也可以发布多主题调查。有单选和多选二种调查类型，只有将调查设为最新调查后才会在前台显示。<br>
      　　快捷菜单：<A href="Admin_Vote.asp?Action=Add" target=main><u><font color="#FF0000">发布网站调查</font></u></A> | <A href="Admin_Vote.asp" target=main><font color="#FF0000"><u>管理网站调查</u></font></A>。</td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Advertisement.asp" target="main">网站广告管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_FriendSite.asp" target="main">友情链接管理</a></td>
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
    <td width="400">　　可以发布、修改和删除网站广告。广告生成为静态JS文件，以方便更新和加快显示速度。可以设置广告版位，广告版位支持网络广告通用标准。同一版位可以设置多个广告，同一广告可以属于不同版位，并有多种显示方式。可以上传广告图片并设置大小，广告支持flash格式。只有将广告版位设为活动时才会在前台显示。<br>
      　　快捷菜单：<A href="Admin_Advertisement.asp" target=main><u><font color="#FF0000">管理网站广告</font></u></A> | <A href="Admin_Advertisement.asp?Action=AddZone" target=main><u><font color="#FF0000">添加广告版位</font></u></A> | <A href="Admin_Advertisement.asp?Action=AddAD" target=main><u><font color="#FF0000">添加新广告</font></u></A>。</td>
    <td width="40">&nbsp;</td>
    <td width="400"> 　　系统具备管理、审核其它网站申请的友情链接功能，可执行添加、修改、删除等操作，可设置推荐链。友情链接分成<A href="Admin_FriendSite.asp?LinkType=2" title="显示以网站标题文字为主的链接形式。"><u>文字链接</u></A>和<A href="Admin_FriendSite.asp?LinkType=1" title="显示以网站logo图片为主的链接形式。"><u>LOGO链接</u></A>二种显示形式。<br>
      　　快捷菜单：<A href="Admin_FriendSite.asp?Action=Add" target=main><font color="#FF0000"><u>添加友情链接</u></font></A> | <A href="Admin_FriendSite.asp" target=main><font color="#FF0000"><u>管理友情链接</u></font></A>。<br>
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
    <td width="100" align="center">&nbsp;</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Counter.asp" target=main>网站统计分析</A></td>
    <td width="300">&nbsp;</td>
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
    <td width="400">&nbsp;</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　显示详细的网站统计信息，可查看网站综合统计信息、最近访问记录、访问次数、链接页面、操作系统等分类信息。<br>
      　　快捷菜单：<A href="Admin_Counter.asp" target=main><font color="#FF0000"><u>网站统计分析</u></font></A> | <A href="Admin_Counter.asp?Action=FVisitor" target=main><font color="#FF0000"><u>最近访问记录</u></font></A>。<br>
    </td>
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
