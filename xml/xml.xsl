<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<xsl:template match="/">
<html>
<head>
<title><xsl:value-of select="powereasy/SiteTitle"/> >> 首页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Skin/DefaultSkin.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="0" topmargin="0">
<div id="menuDiv" style="Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8"></div>

<table class="top_tdbgall" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
<!--顶部空距代码-->
    <tr>
      <td class="top_top" colSpan="2"></td>
    </tr>
<!--频道显示代码-->
    <tr>
      <td colSpan="2">
        <table class="top_Channel" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td align="left"><a href="Rss.asp" Title="Rss 2.0" Target="_blank"><img src="images/rss.gif" border="0" /></a><img src="images/Wap.gif" border="0" alt="WAP浏览支持" style="cursor:hand;"  onClick="window.open('/Wap.asp?ReadMe=Yes', 'Wap', 'width=160,height=257,resizable=0,scrollbars=no');" /></td>
            <td align="right"><xsl:apply-templates select="powereasy"/></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td align="middle"><a><xsl:attribute name="href"><xsl:value-of select="powereasy/SiteUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="powereasy/SiteName"/></xsl:attribute><img src="images/logo.gif" width="180" height="60" border="0" /></a></td>
      <td align="middle"></td>
    </tr>
    <tr>
      <td align="middle" colSpan="2">
      <!--导航、日期代码开始-->
        <table class="top_nav_menu" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
          <tr>
            <td align="middle" width="5%"><IMG src="/Images/arrow3.gif" align="absMiddle" /></td>
            <td width="35%">您现在的位置： <a><xsl:attribute name="href"><xsl:value-of select="powereasy/SiteUrl"/></xsl:attribute><xsl:value-of select="powereasy/SiteName"/></a>  >>  首页</td>
            <td align="right" width="60%">
            <MARQUEE onmouseover="this.stop()" onmouseout="this.start()" scrollAmount="1" scrollDelay="4" width="430" align="left"><p><xsl:apply-templates select="powereasy/AnnounceList"/></p></MARQUEE></td>
            <!--    <td width=70% align="right">今天是：
                        <IMG alt='#[script language="JavaScript" type="text/JavaScript" src="js/date.js"]
                                                [/script]#' src="editor/images/jscript.gif" border=0 $>　　</td> -->
          </tr>
        </table>
      </td>
    </tr>

    <tr>
      <td class="main_shadow" colSpan="2"></td>
    </tr>
</table>
  <!-- ********网页顶部代码结束******** -->
  <!-- ********网页中部代码开始******** -->
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td vAlign="top" width="180">
      <!--用户登录代码开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td><IMG src="skin/blue/login_01.gif" /></td>
          </tr>
          <tr>
            <td vAlign="center" align="middle" background="skin/blue/login_02.gif" height="151"><IFRAME id="UserLogin" src="UserLogin.asp?ShowType=1" frameBorder="0" width="170" scrolling="no" height="145"></IFRAME></td>
          </tr>
          <tr>
            <td><IMG src="skin/blue/login_03.gif" /></td>
          </tr>
        </table>
      <!--用户登录代码结束--></td>
      <td width="5"></td>
      <td vAlign="top" width="354">
      <!--本站最新文章代码开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_575"><A class="class" href="Article/ShowNew.asp"><B>本站最新文章</B></A></td>
          </tr>
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194">
<!-- ********循环输出频道全部文章列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='1']">
  	<xsl:for-each select="//Article">
		<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
	</xsl:for-each>
</xsl:for-each>
</td>
          </tr>
          <tr>
            <td class="main_shadow"></td>
          </tr>
        </table>
      <!--本站最新文章代码结束--></td>
      <td width="5"></td>
      <td vAlign="top" width="216">
      <!--特别推荐代码开始-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_575"><A class="class" href="Article/ShowElite.asp"><B>今日推荐</B></A></td>
          </tr>
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194"></td>
          </tr>
          <tr>
            <td class="main_shadow"></td>
          </tr>
        </table>
      <!--特别推荐代码结束-->
      </td>
    </tr>
  </table>
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
  <!--banner广告显示代码-->
    <tr>
      <td align="middle"></td>
    </tr>
    <tr>
      <td class="main_shadow"></td>
    </tr>
  </table>
  <!--外网搜索代码-->
  <table class="main_Search" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td align="middle"><IMG src="Images/search2.gif" /></td>
      <td align="middle">

	<table class="border" cellSpacing="0" cellPadding="0" width="100%">
        <tr vAlign="center">
<script>
<![CDATA[
function search4()
{
if(websearch.google.checked)
   window.open("http://www.google.com/search?hl=zh-CN&lr=lang_zh-CN&q="+websearch.key.value,"mspg6");
if(websearch.baidu.checked)
   window.open("http://www1.baidu.com/baidu?tn=site5566&word="+websearch.key.value,"mspg9");
if(websearch.sina.checked)
   window.open("http://search.sina.com.cn/cgi-bin/search/search.cgi?_searchkey="+websearch.key.value,"mspg0");
if(websearch.sohu.checked)
   window.open("http://site.search.sohu.com/sitesearch.jsp?page_index=0&key_word="+websearch.key.value,"mspg1");
if(websearch.yahoo.checked)
   window.open("http://cn.search.yahoo.com/search/cn?p="+websearch.key.value,"mspg2");
if(websearch.yeah.checked)
   window.open("http://nisearch.163.com/Search?q="+websearch.key.value,"mspg3");
return false;   
}
]]>
</script>
	<FORM name="websearch" onsubmit="return(search4())">
        <td align="middle" height="40">关键字： 
  	<Input size="18" name="key" /> 
  	<Input type="submit" value="搜索" name="submit" />
  	<Input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" type="checkbox" value="baidu" name="baidu" /> <A href="http://www.baidu.com/" target="_blank">百度</A> 
  	<Input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" type="checkbox" value="sina" name="sina" /> <A href="http://cha.sina.com.cn/" target="_blank">新浪</A> 
  	<Input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" type="checkbox" value="sohu" name="sohu" /> <A href="http://dir.sohu.com/" target="_blank">搜狐</A> 
  	<Input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" type="checkbox" value="yahoo" name="yahoo" /> <A href="http://cn.search.yahoo.com/" target="_blank">雅虎</A> 
  	<Input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" type="checkbox" value="yeah" name="yeah" /> <A href="http://so.163.com/" target="_blank">网易</A> 
  	<Input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" type="checkbox" value="google" name="google" /> <A href="http://www.google.com/intl/zh-CN/" target="_blank">google</A>
	</td></FORM>
        </tr>
        </table>
        
      </td>
    </tr>
    <tr>
      <td class="main_shadow" colSpan="2"></td>
    </tr>
  </table>
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td class="main_shadow"></td>
    </tr>
  </table>

 



 <!--文章频道显示代码-->
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td class="left_tdbgall" vAlign="top" width="180">
      <!--专栏作家开始-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">专 栏 作 家</td>
          </tr>
          <tr>
            <td class="left_tdbg1" vAlign="top" height="179"><xsl:apply-templates select="powereasy/AuthorList"/></td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        <!--专栏作家代码结束-->
      <!--用户排行代码开始-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">用 户 排 行</td>
          </tr>
          <tr>
            <td class="left_tdbg1" vAlign="top" height="126"><xsl:apply-templates select="powereasy/UserList"/></td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
      <!--用户排行代码结束-->
      <!--留言代码开始-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">最 新 留 言</td>
          </tr>
          <tr>
            <td class="left_tdbg1" vAlign="top" height="126">
<!-- ********循环输出留言列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='4']">
	<table width="100%">
  	<xsl:for-each select="//Guest">
		<tr><td width="100"><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></td><td>[<xsl:value-of select="@ReplyNum"/>]</td></tr>
	</xsl:for-each>
	</table>
</xsl:for-each>
</td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
      <!--用户排行代码结束-->
	</td>
      <td width="5"></td>
      <td vAlign="top">
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_shadow" colSpan="3"></td>
          </tr>
          <tr>
            <td vAlign="top">
            <!--栏目一最新文章代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
                <tr>
                  <td class="main_title_282i"><B>栏目一最新文章</B></td>
                </tr>
                <tr>
                  <td class="main_tdbg_282i" vAlign="top" height="136">
<!-- ********循环输出文章列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='1']">
  	<xsl:for-each select="Class[@ClassID='1']">
		<xsl:for-each select="*/Article">
			<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
		</xsl:for-each>
	</xsl:for-each>
</xsl:for-each>
		</td>
                </tr>
              </table>
            <!--栏目一最新文章代码结束--></td>
            <td width="4"></td>
            <td vAlign="top">
            <!--栏目二最新文章代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
                <tr>
                  <td class="main_title_282i"><B>栏目二最新文章</B></td>
                </tr>
                <tr>
                  <td class="main_tdbg_282i" vAlign="top" height="136">
<!-- ********循环输出文章列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='1']">
  	<xsl:for-each select="Class[@ClassID='3']">
		<xsl:for-each select="//Article">
			<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
		</xsl:for-each>
	</xsl:for-each>
</xsl:for-each>
</td>
                </tr>
              </table>
            <!--栏目二最新文章代码结束--></td>
          </tr>
          <tr>
            <td class="main_shadow" colSpan="3"></td>
          </tr>
          <tr>
            <td vAlign="top">
            <!--频道一最新文章代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
                <tr>
                  <td class="main_title_282i"><B>栏目一最新软件</B></td>
                </tr>
                <tr>
                  <td class="main_tdbg_282i" vAlign="top" height="136">
<!-- ********循环输出频道软件列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='2']">
  	<xsl:for-each select="//Soft">
		<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
	</xsl:for-each>
</xsl:for-each>
		</td>
                </tr>
              </table>
            <!--栏目一最新文章代码结束--></td>
            <td width="4"></td>
            <td vAlign="top">
            <!--频道二最新软件代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
                <tr>
                  <td class="main_title_282i"><B>栏目二最新软件</B></td>
                </tr>
                <tr>
                  <td class="main_tdbg_282i" vAlign="top" height="136">
<!-- ********循环输出频道软件列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='2']">
  	<xsl:for-each select="//Soft">
		<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
	</xsl:for-each>
</xsl:for-each>
</td>
                </tr>
              </table>
            <!--栏目二最新文章代码结束--></td>
          </tr>
          <tr>
            <td class="main_shadow" colSpan="3"></td>
          </tr>
          <tr>
            <td vAlign="top">
            <!--频道一最新文章代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
                <tr>
                  <td class="main_title_282i"><B>栏目一最新软件</B></td>
                </tr>
                <tr>
                  <td class="main_tdbg_282i" vAlign="top" height="136">
<!-- ********循环输出频道软件列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='2']">
  	<xsl:for-each select="//Soft">
		<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
	</xsl:for-each>
</xsl:for-each>
		</td>
                </tr>
              </table>
            <!--栏目一最新文章代码结束--></td>
            <td width="4"></td>
            <td vAlign="top">
            <!--频道二最新软件代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
                <tr>
                  <td class="main_title_282i"><B>栏目二最新软件</B></td>
                </tr>
                <tr>
                  <td class="main_tdbg_282i" vAlign="top" height="136">
<!-- ********循环输出频道软件列表******** -->
<xsl:for-each select="powereasy/Channel[@ChannelID='2']">
  	<xsl:for-each select="//Soft">
		<li><a><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/></xsl:attribute><xsl:value-of select="@Title"/></a></li>
	</xsl:for-each>
</xsl:for-each>
</td>
                </tr>
              </table>
            <!--栏目二最新文章代码结束--></td>
          </tr>
          <tr>
            <td class="main_shadow" colSpan="3"></td>
          </tr>
          <tr>
            <td class="main_Search" colSpan="3">
            <!--站内搜索代码开始-->
              <table cellSpacing="0" cellPadding="0" width="100%" border="0">
<FORM name="search" action="search.asp" method="post">
                <tr>
                  <td width="120"><IMG src="Images/search.gif" /></td>
                  <td align="middle">
  <Input type="radio" value="Article" name="ModuleName" /> 文章 
  <Input type="radio" value="Soft" name="ModuleName" /> 下载 
  <Input type="radio" value="Photo" name="ModuleName" /> 图片 
  <Input id="Keyword" maxLength="50" value="关键字" name="Keyword" /> 
  <Input id="Submit" type="submit" value="·搜索·" name="Submit" /> 
  <Input id="Field" type="hidden" value="Title" name="Field" /></td>
                </tr>
</FORM>
              </table>
            <!--站内搜索代码结束--></td>
          </tr>
          <tr>
            <td class="main_shadow" colSpan="3"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td class="main_shadow"></td>
    </tr>
  </table>
  <!--横向友情链接代码开始-->
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td class="main_title_760i"><B><A class="class" href="FriendSite/index.asp">友情链接</A></B></td>
    </tr>
    <tr>
      <td class="main_tdbg_760i" align="middle"></td>
    </tr>
    <tr>
      <td class="main_shadow"></td>
    </tr>
  </table>
  <!--横向友情链接代码结束-->
  <!-- ********网页中部代码结束******** -->
  <!-- ********网页底部代码开始******** -->
  <table class="Bottom_tdbgall" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr align="middle">
      <td class="Bottom_Adminlogo" colSpan="2">| <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{powereasy/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('http://localhost','动易网络');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="powereasy/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="Copyright.asp" target="_blank">版权申明</A> |  <a class="Bottom" href="Admin/Admin_Index.asp" target="_blank">管理登录</a>  </td>
    </tr>
    <tr class="Bottom_Copyright">
      <td width="20%"><IMG src="images/logo.gif" /></td>
      <td width="80%" align="center"> 站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="powereasy/WebmasterEmail"/></xsl:attribute><xsl:value-of select="powereasy/WebmasterName"/></a><br /><xsl:value-of select="powereasy/Copyright"/></td>
    </tr>
  </table>
</body>
</html>
</xsl:template>

<!-- ********循环输出频道列表******** -->
<xsl:template match="powereasy">
| <a class="Channel2" href="Index.html">首页</a>
	<xsl:for-each select="Channel">
	| <a class="Channel"><xsl:attribute name="href"><xsl:value-of select="@LinkUrl"/>/index.asp</xsl:attribute><xsl:value-of select="@ChannelName"/></a>
	</xsl:for-each> 
</xsl:template>

<!-- ********循环输出用户列表******** -->
<xsl:template match="powereasy/UserList">
	<table width="100%">
	<tr align="center"><td width="50">用户名称</td><td>发表文章</td></tr>
  	<xsl:for-each select="User">
		<tr align="center"><td><a href="{@UserID}"><xsl:value-of select="@NickName"/></a></td><td><xsl:value-of select="@PassedItems"/></td></tr>
	</xsl:for-each>
	</table>
</xsl:template>

<!-- ********循环输出作者列表******** -->
<xsl:template match="powereasy/AuthorList">
	<table width="100%">
  	<xsl:for-each select="Author">
		<tr align="center"><td width="50"><a href="{Author@AuthorID}"><xsl:value-of select="@AuthorName"/></a></td><td><xsl:value-of select="@NickName"/></td></tr>
	</xsl:for-each>
	</table>
</xsl:template>

<!-- ********循环输出公告******** -->
<xsl:template match="powereasy/AnnounceList">
  	<xsl:for-each select="Announce[@ShowType='0' or @ShowType='1']">
		<xsl:value-of select="@Title"/>  (<xsl:value-of select="@Author"/>)
	</xsl:for-each>
</xsl:template>
</xsl:stylesheet>