<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">
<xsl:element name="html">
<head>
<title><xsl:value-of select="body/MyBlog/BlogName"/></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Skin/DefaultSkin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<![CDATA[
/* 拖动面板CSS定义 */
div#divContainer {
	width: 100%;
}

div#itembody {
	float: left;
        width: 280px;

	margin: 2px;
	border: 1px solid #d2d3d9;
}

div#itemtitle {
        color: ffffff;background:url(skin/blue/main_title_575.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;height: 29;
}
]]>
</style>
<script src="JS/prototype.js"></script>
<script src="JS/scriptaculous.js"></script>
<script language="JavaScript">
<![CDATA[
function GetXmlData(iurl,divname,lnum)
{
    var url = "rssfeed.asp?l=" + lnum + "&url=" + iurl;
    var myAjax = new Ajax.Updater({success: "item_" + divname}, url, {method: 'get', onFailure: "read error"});
}
]]>
</script>
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
            <td align="middle" width="50"><IMG src="Images/arrow3.gif" align="absMiddle" /></td>
            <td width="40%">您现在的位置： <a class='LinkPath'><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> <xsl:value-of select="body/MyBlog/BlogName"/></td>
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
      <td vAlign="top">
      <!--显示我的 Blog栏目-->
<div id="divContainer">
        <xsl:apply-templates select="body/MyBlog/Blogitem"/>
</div>
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
      <td class="main_shadow"></td>
    </tr>
  </table>
  <!-- ********网页中部代码结束******** -->
  <!-- ********网页底部代码开始******** -->
  <table class="Bottom_tdbgall" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr align="middle">
      <td class="Bottom_Adminlogo" colSpan="2">| <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="Copyright.asp" target="_blank">版权申明</A> | <a class="Bottom" href="Admin/Admin_Index.asp" target="_blank">管理登录</a> |</td>
    </tr>
    <tr class="Bottom_Copyright">
      <td width="20%"><IMG src="images/logo.gif" /></td>
      <td width="80%" align="center"> 站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a><br /><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/></td>
    </tr>
  </table>
<script type="text/javascript">
// <![CDATA[
	Sortable.create('divContainer',{tag:'div',overlap:'horizontal',constraint:false});
// ]]>
</script>
</body>
</xsl:element>
</xsl:template>

<xsl:template match="body/MyBlog/Blogitem">
        <div id="itembody">
        <div id="itemtitle"><table width="100%"><tr valign="middle"><td width="40"><a href="#" onclick="new Element.toggle('item_{title}')"><img src="images/jiaodian_biao.gif" border="0" /></a></td><td><font color="white"><xsl:value-of select="title"/></font></td><td align="right"><a href="{link}" target="_blank"><img src="images/Rss.gif" border="0" align="absmiddle"/></a></td></tr></table></div>
        <div id="item_{title}">载入中...</div>
        <script language='JavaScript'>GetXmlData("<xsl:value-of select="link"/>","<xsl:value-of select="title"/>",<xsl:value-of select="listnum"/>);</script>
        </div>
</xsl:template>

<xsl:template match="Channelitem">
        <xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{link}" class="channel"><xsl:value-of select="title"/></a>
</xsl:template>

<xsl:template match="Announceitem">
        <a href="{link}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:value-of select="DateAndTime"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
</xsl:template>

</xsl:stylesheet>