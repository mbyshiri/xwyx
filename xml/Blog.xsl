<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">

<xsl:element name="html">
<head>
<title><xsl:value-of select="body/Site/SiteTitle"/> >> 作品集列表页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Skin/DefaultSkin.css" rel="stylesheet" type="text/css" />
</head>

<body leftmargin="0" topmargin="0">
<table class="top_tdbgall" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td class="top_top" colSpan="2"></td>
    </tr>
    <tr>
      <td colSpan="2">
        <table class="top_Channel" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td align="left"><a href="Rss.asp" Title="Rss 2.0" Target="_blank"><img src="images/rss.gif" border="0" /></a><img src="images/Wap.gif" border="0" alt="WAP浏览支持" style="cursor:hand;"  onClick="window.open('/Wap.asp?ReadMe=Yes', 'Wap', 'width=160,height=257,resizable=0,scrollbars=no');" /></td>
            <td align="right">|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{body/Site/SiteUrl}" class="channel">网站首页</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:apply-templates select="body/ChannelList/Channelitem"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text></td>
          </tr>
        </table>
      </td>
    </tr>

    <tr>
      <td align="middle"><a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/SiteLogo}" width="180" height="60" border="0" /></a></td>
      <td align="middle"></td>
    </tr>
    <tr>
      <td align="middle" colSpan="2">
      <!--导航、日期代码开始-->
        <table class="top_nav_menu" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
          <tr>
            <td align="middle" width="50"><IMG src="/Images/arrow3.gif" align="absMiddle" /></td>
            <td width="40%">您现在的位置： <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> 作品集列表页</td>
            <td align="right">
            <MARQUEE onmouseover="this.stop()" onmouseout="this.start()" scrollAmount="1" scrollDelay="4" width="430" align="left"><p><xsl:apply-templates select="body/AnnounceList/Announceitem"/></p></MARQUEE></td>
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
      <!--本站最新Blog代码开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_575"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>最近更新作品集</B></td>
          </tr>
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194">
            <table width="100%"><xsl:apply-templates select="body/NewBlog/Blogitem"/></table>
            </td>
          </tr>
          <tr>
            <td class="main_shadow"></td>
          </tr>
        </table>
      <!--本站最新Blog代码开始--></td>
      <td width="5"></td>
      <td vAlign="top" width="216">
      <!--特别推荐代码开始-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_575"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>新加入作品集</B></td>
          </tr>
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194">
               <!-- ********循环输出最近加入BLOG列表******** -->
               <table width="100%"><xsl:apply-templates select="body/AddBlog/Blogitem"/></table>
            </td>
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
    <tr>
      <td class="main_shadow"></td>
    </tr>
  </table>

  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td vAlign="top" width="100%">
      <!--本站推荐作品集代码开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_760"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>推荐作品集</B></td>
          </tr>
          <tr>
            <td align="center"><xsl:apply-templates select="body/EliteBlog/Blogitem"/></td>
          </tr>
          <tr>
            <td class="main_shadow"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td vAlign="top" width="100%">
      <!--作品集分类开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_760"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>作品集分类列表</B></td>
          </tr>
          <tr>
            <td><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><a href="showblog.asp">全部分类</a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><xsl:apply-templates select="body/BlogClassList/item"/></td>
          </tr>
          <tr>
            <td class="main_shadow"></td>
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
  <!-- ********网页中部代码结束******** -->
  <!-- ********网页底部代码开始******** -->
  <table class="Bottom_tdbgall" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr align="middle">
      <td class="Bottom_Adminlogo" colSpan="2"> | <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="Copyright.asp" target="_blank">版权申明</A> |  <a class="Bottom" href="Admin/Admin_Index.asp" target="_blank">管理登录</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text></td>
    </tr>
    <tr class="Bottom_Copyright">
      <td width="20%"><IMG src="images/logo.gif" /></td>
      <td width="80%" align="center"> 站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a><br /><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/></td>
    </tr>
  </table>
</body>
</xsl:element>
</xsl:template>


<xsl:template match="Blogitem">
	<tr><td><a href="{link}"><xsl:value-of select="title"/></a></td><td width="100" align="right"><xsl:value-of select="BirthDay"/></td></tr>
</xsl:template>

<xsl:template match="body/AddBlog/Blogitem">
	<tr><td><a href="{link}"><xsl:value-of select="title"/></a></td><td width="100" align="right"><xsl:value-of select="author"/></td></tr>
</xsl:template>

<xsl:template match="body/EliteBlog/Blogitem">
	<a href="{link}">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="Photo"/></xsl:attribute>
		<xsl:attribute name="alt">名称:<xsl:value-of select="title"/><br/>日期:<xsl:value-of select="BirthDay"/></xsl:attribute>
		<xsl:attribute name="border">0</xsl:attribute>
		<xsl:attribute name="hight">160</xsl:attribute>
		<xsl:attribute name="width">120</xsl:attribute>
	</xsl:element>
        </a>
        <xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
</xsl:template>

<xsl:template match="body/BlogClassList/item">
	<a href="showblog.asp?TypeID={id}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
</xsl:template>

<xsl:template match="Channelitem">
        <xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{link}" class="channel"><xsl:value-of select="title"/></a>
</xsl:template>

<xsl:template match="Announceitem">
        <a href="{link}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:value-of select="DateAndTime"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
</xsl:template>

</xsl:stylesheet>