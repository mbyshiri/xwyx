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
%>
<html>
<head>
<title><%=SiteName & "--后台管理首页"%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
<style type="text/css">
<!--
body {
    background-color: #FFFFFF;
    margin-left: 0px;
}
.STYLE4 {color: #000000}
-->
</style>
</head>
<body topmargin="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="392" rowspan="2"><img src="Images/adminmain01.gif" width="392" height="126"></td>
    <td height="114" valign="top" background="Images/adminmain0line2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20"></td>
      </tr>
      <tr>
        <td><span class="STYLE4">帮助公告</span></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><div id="peinfo1">正在读取数据中...</div><div id="peinfo2" style="z-index: 1; visibility: hidden; position: absolute"></div>
          <div id="peinfo5" style=" visibility: hidden;"></div>
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellpadding="3" cellspacing="0">
      
      <tr>
        <td width="31%" height="87" align="right"><table width="94%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><%=AdminName%>您好， </td>
            </tr>
            <tr>
              <td valign="top">今天是
                <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script>
                您尚有：</td>
            </tr>

            <tr>
              <td valign="top"><%=ShowUnPassedInfo%></td>
            </tr>
    </table></td>
        <td width="1%"><table width="3" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="3" height="65" bgcolor="#1890CC"></td>
            </tr>
        </table></td>
        <td width="68%">欢迎您进入<%=SiteName%>网站后台管理系统！在这里您可以利用系统提供的强大的HTML生成功能，便捷的后台管理功能，栏目无限级分类，任意添加网站频道功能，栏目批量设置、批量移动等功能有效地管理网站。您可以随时使用顶部的<font color="#FF0000">关闭左栏</font>功能关闭或打开左边的管理导航，以扩展操作界面。初次架设网站请配置以下信息：</td>
      </tr>
      <tr>
        <td height="5" colspan="3"></td>
        </tr>
    </table></td>
  </tr>
</table>
<%
Dim rsArticleInfo, sqlArticleInfo, rsSoftInfo, sqlSoftInfo, rsGuestBookInfo, sqlGuestBookInfo, rsPhotoInfo, sqlPhotoInfo, rsChannelInfo, sqlChannelInfo, rsCommentInfo, sqlCommentInfo, rsUserInfo, sqlUserInfo
Dim channelIDinfo, channelIDinfonew, a1, a2, vNoApproveUser, vTestUser, c1, c2, cdir
Dim vArticleID, vChannelID, vChannelIDAry(50), vChannelName(50), vCommentCount(50), x, y
%>
<%
If ShowUnpass = True Then
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">待审文章</td>
    <td width="100"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">待审软件</td>
    <td width="100"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">待审图片</td>
    <td width="100"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">待审留言</td>
    <td width="40"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">待审评论</td>
    <td width="40"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">待审会员</td>
    <td width="40"></td>
    <td width="20" rowspan="2"></td>
  </tr>
  <tr>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20"></td>
    <td width="180" valign="top"><% ShowUnPassedChannel(1) %></td>
    <td width="20"></td>
    <td width="180" valign="top"><% ShowUnPassedChannel(2) %></td>
    <td width="20"></td>
    <td width="180" valign="top"><% ShowUnPassedChannel(3) %></td>
    <td width="20"></td>
    <td width="120" valign="top"><% ShowUnPassedGuestBook() %></td>
    <td width="20"></td>
    <td width="120" valign="top"><% ShowUnPassedComment() %></td>
    <td width="20"></td>
    <td width="120" valign="top"><% ShowUnPassedUser() %></td>
  </tr>
</table>

<%
End If
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="400" align="center" class="topbg"><span class="Glow">建 站 管 理 快 捷 入 口</span></td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="400" align="center" class="topbg"><span class="Glow">日 常 管 理 快 捷 入 口</span></td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1"></td>
    <td></td>
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
        <td width="100" align="center" class="topbg"><a class='Class' href="Admin_SiteConfig.asp">网站信息配置</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
        <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Class.asp?ChannelID=1" target="main">网站栏目管理</a></td>
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
      　　快捷菜单：<a href="Admin_SiteConfig.asp" target="main"><u><font color="#FF0000">网站信息配置</font></u></a>
      | <a href="Admin_SiteConfig.asp#SiteOption" target="main"><u><font color="#FF0000">网站选项配置</font></u></a>
    | <a href="Admin_SiteConfig.asp#User" target="main"><u><font color="#FF0000">用户选项</font></u></a></td>
    <td width="40">&nbsp;</td>
    <td width="400">　　管理网站各频道中所设置的各级栏目，栏目具有无级分类功能。可对栏目进行添加、删除、排序、复位、合并和批量设置等管理。<br>
      　　<font color="#0000FF">初次安装请在各频道中先添加栏目</font>。<br>
      　　快捷菜单：<a href="Admin_Class.asp?ChannelID=1" target="main"><u><font color="#FF0000">文章栏目管理</font></u></a>
      | <a href="Admin_Class.asp?ChannelID=2" target="main"><u><font color="#FF0000">下载栏目管理</font></u></a>
      | <a href="Admin_Class.asp?ChannelID=3" target="main"><u><font color="#FF0000">图片栏目管理</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_CreateSiteIndex.asp" target="main">首页生成管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Article.asp?ChannelID=1&Action=Add&AddType=2" target="main">网站内容添加</a></td>
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
    <td width="400">　　首次安装系统请<a href="Admin_CreateSiteIndex.asp" target="main"><font color="#FF0000">生成网站首页</font></a>。首页、栏目页、内容页、专题页……都可以生成完全的HTML页面（评论和点击数统计除外）。各频道启用生成功能请在<a href=Admin_Channel.asp target=main title="动易网站管理系统中，频道是指某一功能模板的集合。某一频道可以是具备文章系统功能，或具备下载系统、图片系统的功能。"><U><font color="#FF0000">网站频道管理</font></U></a>中设置。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　系统提供强大的<a href="#" title="在线编辑器能够在网页上实现许多桌面编辑软件（如：Word）所具有的强大可视编辑内容的功能。"><u>在线编辑器</u></a>（文章中心），增加、删除、修改网站各个频道下各栏目的相关内容（文字、软件、图片等），方便设置内容的相关属性等。<br>
      　　快捷菜单：<a href="Admin_Article.asp?ChannelID=1&Action=Add&AddType=2" target="main"><u><font color="#FF0000">添加文章</font></u>
      </a><a href="Admin_Article.asp?ChannelID=1&Action=Manage&Passed=All"><u><font color="#FF0000">管理</font></u></a>
      | <a href="Admin_Soft.asp?ChannelID=2&Action=Add&AddType=2" target="main"><u><font color="#FF0000">添加软件</font></u></a>
      <a href="Admin_Article.asp?ChannelID=2&Action=Manage&Passed=All"><u><font color="#FF0000">管理</font></u></a> | <a href="Admin_Photo.asp?ChannelID=3&Action=Add&AddType=2" target="main"><u><font color="#FF0000">添加图片</font></u></a>
      <a href="Admin_Article.asp?ChannelID=3&Action=Manage&Passed=All"><u><font color="#FF0000">管理</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Admin.asp" target="main">管理员管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_GuestBook.asp?Passed=All" target="main">网站留言管理</a></td>
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
    <td width="400">　　强大的<a href="Admin_Admin.asp" target="main"><U><font color="#FF0000">网站权限管理</font></U></a>，可增删管理员和指定详细的管理权限，使网站的管理分级分类多人共同管理。设置网站添加<a href="Admin_Admin.asp?Action=Add" target="main" title="超级管理员：拥有所有权限。某些权限（如管理员管理、网站信息配置、网站选项配置等管理权限）只有超级管理员才有。"><U>超级管理员</U></a>和<a href="Admin_Admin.asp?Action=Add" target="main" title="普通管理员：捅有指定部分网站管理功能，需要详细指定每一项管理权限。"><U>普通管理员</U></a>，您也可以自由设定<a href="Admin_UserGroup.asp" target="main" title="用户组是用户账户的集合，通过创建用户组，赋予相关用户享有授予组的权力和权限。具体的权限设置在“频道管理”及各频道的“栏目管理”中。"><font color="#FF0000"><U>用户组</U></font></a>以管理注册用户级别。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　对用户的留言进行审核、修改、删除、回复等操作。留言的审核功能请在<a href=Admin_Channel.asp target=main title="网站管理系统中，频道是指某一功能模板的集合。某一频道可以是具备文章系统功能，或具备下载系统、图片系统的功能。"><U>网站频道管理</U></a>中设置。<br>
      　　快捷菜单：<a href="Admin_GuestBook.asp?Passed=False" target="main"><u><font color="#FF0000">审核留言</font></u></a>
      | <a href="Admin_GuestBook.asp?Passed=All" target="main"><u><font color="#FF0000">管理留言</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Channel.asp" target="main" title="网站管理系统中，频道是指某一功能模板的集合。某一频道可以是具备文章系统功能，或具备下载系统、图片系统的功能。">网站频道管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Advertisement.asp" target="main">网站广告管理</a></td>
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
    <td width="400">　　<a href="Admin_Channel.asp" target="main">管理网站的各个频道的功能模块，如文章、下载、图片和留言等频道。</a>频道可分为<a href="#" title="系统内部频道指的是在MY动力现有功能模块（新闻、文章、图片等）基础上添加新的频道，新频道具备和所使用功能模块完全相同的功能。"><U>系统内部频道</U></a>与<a href="#" title="外部频道指链接到MY动力系统以外的地址中。当此频道准备链接到网站中的其他系统时"><U>外部频道</U></a>。频道类型请慎重选择，频道一旦添加后就不能再更改频道类型。</td>
    <td width="40">&nbsp;</td>
    <td width="400">　　系统提供强大的<a href="Admin_Advertisement.asp" target=main title="在线编辑器能够在网页上实现许多桌面编辑软件（如：Word）所具有的强大可视编辑内容的功能。"><u>网站广告管理</u></a>功能，对网站的广告版位及广告进行增加、修改、删除等操作。<br>
    快捷菜单：<a href="Admin_Advertisement.asp?Action=ZoneList" target="main"><u><font color="#FF0000">广告版位管理</font></u></a> | <a href="Admin_Advertisement.asp?Action=AddZone" target="main"><u><font color="#FF0000">添加新版位</font></u></a> | <a href="Admin_Advertisement.asp?Action=ADList" target="main"><u><font color="#FF0000">网站广告管理</font></u></a> | <a href="Admin_GuestBook.asp?Passed=All" target="main"></a><a href="Admin_Advertisement.asp?Action=AddAD" target="main"><u><font color="#FF0000">添加新广告</font></u></a></td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 colspan=2 class="topbg"><span class="Glow">服 务 器 信 息</span>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">服务器类型：      <%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td width="50%">脚本解释引擎：
    <%
    response.write ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion
    If CSng(ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion) < 5.6 Then
        response.write "&nbsp;&nbsp;<a href='http://www.microsoft.com/downloads/release.asp?ReleaseID=33136' target='_blank'><font color='green'>版本过低，请点此更新</font></a>"
    End If
    %>
    </td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">站点物理路径：      <%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td width="50%">数据库使用：<%ShowObjectInstalled("adodb.connection")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">FSO文本读写：<%ShowObjectInstalled(objName_FSO)%></td>
    <td width="50%">数据流读写：<%ShowObjectInstalled("Adodb.Stream")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">XMLHTTP组件支持：<%ShowObjectInstalled("Microsoft.XMLHTTP")%></td>
    <td width="50%">XMLDOM组件支持：<%ShowObjectInstalled("Microsoft.XMLDOM")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">XML组件支持：<%ShowObjectInstalled("MSXML2.XMLHTTP")%></td>
    <td width="50%">AspJpeg组件支持：<%ShowObjectInstalled("Persits.Jpeg")%></td>
  </tr>
  
  <tr class="tdbg" height=23>
    <td width="50%">Jmail组件支持：<%ShowObjectInstalled("JMail.SMTPMail")%></td>
    <td width="50%">CDONTS组件支持：<%ShowObjectInstalled("CDONTS.NewMail")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">ASPEMAIL组件支持：<%ShowObjectInstalled("Persits.MailSender")%></td>
    <td width="50%">WebEasyMail组件支持：<%ShowObjectInstalled("easymail.MailSend")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%"> </td>
    <td width="50%" align="right"><a href="Admin_ServerInfo.asp">点此查看更详细的服务器信息&gt;&gt;&gt;</a></td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 class="topbg"><span class="Glow">Copyright 2003-2006 &copy; <%=SiteName%> All Rights Reserved.</span>
  </tr>
</table>
<div id="peinfo3" style="height:1;overflow=auto;visibility:hidden;">
<script src="http://www.powereasy.net/PowerEasy2006_Info.asp?Action=ShowAnnounce"></script>
</div>
<div id="peinfo4" style="height:1;overflow=auto;visibility:hidden;">
<script src="http://www.powereasy.net/PowerEasy2006_Info.asp?Action=ShowPatchAnnounce"></script>
</div>
<script language="JavaScript">
marqueesHeight=36;
scrillHeight=20;
scrillspeed=60;
stoptimes=50;
stopscroll=false;preTop=0;currentTop=0;stoptime=0;
peinfo1.scrollTop=0;
with (peinfo1)
{
  style.width=0;
  style.height=marqueesHeight;
  style.overflowX='visible';
  style.overflowY='hidden';
  noWrap=true;
  onmouseover=new Function("stopscroll=true");
  onmouseout=new Function("stopscroll=false");
}
function init_srolltext()
{
  peinfo2.innerHTML='';
  peinfo2.innerHTML+=peinfo3.innerHTML;
  peinfo1.innerHTML=peinfo3.innerHTML+peinfo3.innerHTML;
  setInterval("scrollUp()",scrillspeed);
}

function init_peifo()
{
  peinfo5.innerHTML=peinfo4.innerHTML;
}
function scrollUp()
{
  if(stopscroll==true) return;
  currentTop+=1;
  if(currentTop==scrillHeight)
  {
   stoptime+=1;
   currentTop-=1;
   if(stoptime==stoptimes) { currentTop=0; stoptime=0; }
  }
  else
  {
   preTop=peinfo1.scrollTop;
   peinfo1.scrollTop+=1;
   if(preTop==peinfo1.scrollTop){ peinfo1.scrollTop=peinfo2.offsetHeight-marqueesHeight; peinfo1.scrollTop+=1; }
  }
}
init_peifo();
setInterval("",1000);
init_srolltext();
</script>
</body>
</html>
<%
Call CloseConn


Function ShowUnPassedInfo()
    Dim rsCount, UnPassed_Article, UnPassed_Soft, UnPassed_Photo, UnPassed_Message, strInfo
    Set rsCount = Conn.Execute("select count(ArticleID) from PE_Article where Deleted=" & PE_False & " and Status=0")
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Article = rsCount(0)
    Else
        UnPassed_Article = 0
    End If
    Set rsCount = Conn.Execute("select count(SoftID) from PE_Soft where Deleted=" & PE_False & " and Status=0")
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Soft = rsCount(0)
    Else
        UnPassed_Soft = 0
    End If
    Set rsCount = Conn.Execute("select count(PhotoID) from PE_Photo where Deleted=" & PE_False & " and Status=0")
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Photo = rsCount(0)
    Else
        UnPassed_Photo = 0
    End If
    Set rsCount = Conn.Execute("select count(GuestID) from PE_GuestBook where GuestIsPassed=" & PE_False)
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Message = rsCount(0)
    Else
        UnPassed_Message = 0
    End If
    Set rsCount = Nothing
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>待审文章："
    strInfo = strInfo & "<font color=red>" & UnPassed_Article & "</font>篇&nbsp;&nbsp;"
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>待审下载："
    strInfo = strInfo & "<font color=red>" & UnPassed_Soft & "</font>个<br>"
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>待审图片："
    strInfo = strInfo & "<font color=red>" & UnPassed_Photo & "</font>个&nbsp;&nbsp;"
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>待审留言："
    strInfo = strInfo & "<font color=red>" & UnPassed_Message & "</font>条"
    ShowUnPassedInfo = strInfo
End Function

Sub ShowObjectInstalled(strObjName)
    If IsObjInstalled(strObjName) Then
        response.write "<b>√</b>"
    Else
        response.write "<font color='red'><b>×</b></font>"
    End If
End Sub

Function ShowUnPassedChannel(ChannelModuleType)
    Dim rsChannel, rsChannelCount, ModuleName, NoneInfo
    
    Select Case PE_CLng(ChannelModuleType)
    Case 1
        ModuleName = "Article"
    Case 2
        ModuleName = "Soft"
    Case 3
        ModuleName = "Photo"
    Case Else
    
    End Select
    NoneInfo = True
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName from PE_Channel where ModuleType = " & PE_CLng(ChannelModuleType))
    Do While Not rsChannel.EOF
        Set rsChannelCount = Conn.Execute("select count(" & ModuleName & "ID) from PE_" & ModuleName & " where Deleted=" & PE_False & " and ChannelID = " & rsChannel("ChannelID") & " and Status=0")
        If rsChannelCount(0) > 0 Then
            response.write "<a href=Admin_" & ModuleName & ".asp?ChannelID=" & PE_CLng(rsChannel("ChannelID")) & "&Action=Manage&ClassID=0&SpecialID=0&Status=0 target=main>" & rsChannel("ChannelName") & "</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]  "
        NoneInfo = False
        End If
    rsChannel.movenext
    Loop
    If NoneInfo = True Then
        response.write "没有待审核信息"
    End If
    rsChannelCount.Close
    Set rsChannelCount = Nothing
    rsChannel.Close
    Set rsChannel = Nothing
    
End Function

Function ShowUnPassedComment()
    Dim rsChannelCount, rs, rsChannel
    Set rsChannel = Conn.Execute("select count(CommentID) from PE_Comment where Passed =" & PE_False)
    If rsChannel(0) > 0 Then
        response.write "待审核评论:[<span style='color:#ff0000'>" & rsChannel(0) & "</span>]"
    Else
        response.write "没有待审核评论"
    End If
    Set rsChannelCount = Conn.Execute("select top 20 * from PE_Comment where Passed = " & PE_False)
    rsChannel.Close
    Set rsChannel = Nothing
End Function


Function ShowUnPassedGuestBook()
    Dim rsChannelCount
    Set rsChannelCount = Conn.Execute("select count(GuestID) from PE_GuestBook where GuestIsPassed =" & PE_False)
    If rsChannelCount(0) > 0 Then
        response.write "<a href='Admin_GuestBook.asp?Passed=False' target=main>留言版</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]"
    Else
        response.write "没有待审核留言"
    End If
    rsChannelCount.Close
    Set rsChannelCount = Nothing
End Function

Function ShowUnPassedUser()
    Dim rsChannelCount
    Set rsChannelCount = Conn.Execute("select count(UserID) from PE_User where GroupID = 7")
    If rsChannelCount(0) > 0 Then
        response.write "<a href='Admin_User.asp?SearchType=11&GroupID=7' target=main>待审批会员</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]<br>"
    Else
        response.write "没有待审批会员<br>"
    End If
    Set rsChannelCount = Conn.Execute("select count(UserID) from PE_User where GroupID = 8")
    If rsChannelCount(0) > 0 Then
        response.write "<a href='Admin_User.asp?SearchType=11&GroupID=8' target=main>未验证会员</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]"
    Else
        response.write "没有未验证会员"
    End If
        
    rsChannelCount.Close
    Set rsChannelCount = Nothing
End Function

%>
