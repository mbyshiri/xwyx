<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">
<xsl:element name="html">
<head>
<title><xsl:value-of select="body/Site/SiteTitle"/> >> 聚合空间首页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="OceanSkin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<![CDATA[
div#classelite {
	width: 100%;
	float: center;
	margin: 8px;
}
div#classlist {
	width: 100%;
	float: center;
	margin: 8px;
}
]]>
</style>
<script src="../JS/prototype.js"></script>
<script src="../JS/scriptaculous.js"></script>
<script src="../JS/checklogin.js"></script>
</head>
<body leftmargin="0" topmargin="0">
  <table height="114" cellSpacing="0" cellPadding="0" width="778" align="center" background="../Skin/Ocean/top_bg.jpg" border="0">
    <tr>
      <td width="213">
        <xsl:choose>  
            <xsl:when test="string-length(body/Site/SiteLogo) > 0">  
            <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/SiteLogo}" width="213" height="114" border="0" /></a>
            </xsl:when>
        </xsl:choose>
      </td>
      <td>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td align="right" colSpan="2">
              <table cellSpacing="0" cellPadding="0" align="right" border="0">
                <tr>
                  <td><IMG height="25" src="../Skin/Ocean/Announce_01.jpg" width="68" /></td>
                  <td class="showa" width="280" background="../Skin/Ocean/Announce_02.jpg">
                  <MARQUEE onmouseover="this.stop()" onmouseout="this.start()" scrollAmount="1" scrollDelay="4" width="430" align="left"><p><xsl:apply-templates select="body/AnnounceList/Announceitem"/></p></MARQUEE></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="83%" height="80">
              <xsl:choose>  
                <xsl:when test="string-length(body/Site/BannerUrl) > 0">  
                <a><xsl:attribute name="href"><xsl:value-of select="body/Site/BannerUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/BannerUrl}" width="468" height="60" border="0" /></a>
                </xsl:when>
              </xsl:choose> 
            </td>
            <td width="17%">
              <table height="89" cellSpacing="0" cellPadding="0" width="94" background="../Skin/Ocean/topr.gif" border="0">
                <tr>
                  <td align="middle" colSpan="2">
                    <table height="56" cellSpacing="0" cellPadding="0" width="79" border="0">
                      <tr>
                        <td align="middle" width="26"><IMG height="13" src="../Skin/Ocean/arrows.gif" width="13" /></td>
                        <td width="68"><a class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</a></td>
                      </tr>
                      <tr>
                        <td align="middle"><IMG height="13" src="../Skin/Ocean/arrows.gif" width="13" /></td>
                        <td><a class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" href="#">设为首页</a></td>
                      </tr>
                      <tr>
                        <td align="middle"><IMG height="13" src="../Skin/Ocean/arrows.gif" width="13" /></td>
                        <td><a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a></td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table cellSpacing="0" cellPadding="0" width="778" align="center" border="0">
    <tr>
      <td class="menu_s" align="middle">
            <xsl:choose>  
                <xsl:when test="body/Site/ShowSiteChannel = 'enable'">  
                |<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{body/Site/SiteUrl}" class="channel">网站首页</a><xsl:apply-templates select="body/ChannelList/Channelitem"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
                </xsl:when>
            </xsl:choose>
            <xsl:choose>  
                <xsl:when test="body/Site/EnableRss = 'enable'">  
                <a href="../Rss.asp" Title="Rss 2.0" Target="_blank" onmouseover="showmenu(event,'0','rss','http://localhost/Rss.asp');" onmouseout="delayhidemenu();"><img src="../images/rss.gif" border="0" /></a>
                </xsl:when>
            </xsl:choose>
     </td>
    </tr>
    <tr>
      <td><IMG height="7" src="../Skin/Ocean/menu_bg2.jpg" width="778" /></td>
    </tr>
    <tr>
      <td background="../Skin/Ocean/addr.jpg" height="21">
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td align="middle" width="5%"><IMG height="17" src="../Skin/Ocean/arrows2.gif" width="16" /></td>
            <td width="95%">您现在的位置： <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> 聚合空间首页</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td background="../Skin/Ocean/addr_line.jpg" height="4"></td>
    </tr>
  </table> 
  <!-- ********网页顶部代码结束******** -->
  <!-- ********网页中部代码开始******** -->
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td vAlign="top" width="180">
      <!--用户登录代码开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onClick="new Element.toggle('login')">
            <td><IMG src="../Skin/Ocean/login_01.jpg" /></td>
          </tr>
          <tbody id="login">
          <tr>
            <td vAlign="center" align="middle" background="../Skin/Ocean/login_02.gif" style="padding:1px">
            <div id="UserLogin">载入中<script language="JavaScript" type="text/JavaScript">LoadUserLogin("../",0,0);</script></div></td>
          </tr>
          </tbody>
          <tr>
            <td><IMG src="../Skin/Ocean/login_03.jpg" /></td>
          </tr>
        <tr> 
          <td Class="main_shadow"></td> 
        </tr> 
        </table>
      <!--用户登录代码结束--></td>
      <td width="5"></td>
      <td vAlign="top" width="354">
      <!--本站最新聚合代码开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('newblog')">
            <td class="main_title_575"><B>最近有更新的聚合空间</B></td>
          </tr>
          <tbody id="newblog">
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194">
            <table width="100%"><xsl:apply-templates select="body/NewBlog/Blogitem"/></table>
            </td>
          </tr>
          </tbody>
          <tr>
            <td class="main_shadow"></td>
          </tr>
        </table>
      <!--本站最新Blog代码开始--></td>
      <td width="5"></td>
      <td vAlign="top" width="216">
      <!--特别推荐代码开始-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('newjoinblog')">
            <td class="main_title_575"><B>新加入的聚合空间</B></td>
          </tr>
          <tbody id="newjoinblog">
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194">
               <!-- ********循环输出最近加入BLOG列表******** -->
               <table width="100%"><xsl:apply-templates select="body/AddBlog/Blogitem"/></table>
            </td>
          </tr>
          </tbody>
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
      <!--推荐聚合开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('eliteblog')">
            <td class="main_title_760">　　　<B>推荐聚合</B></td>
          </tr>
          <tbody id="eliteblog">
          <tr>
            <td align="center"><div id="classelite"><xsl:apply-templates select="body/EliteBlog/Blogitem"/></div></td>
          </tr>
          </tbody>
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
      <!--聚合分类开始-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('blogclass')">
            <td class="main_title_760">　　　<B>聚合分类列表</B></td>
          </tr>
          <tbody id="blogclass">
          <tr>
            <td><div id="classlist"><a href="index.asp">全部分类</a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><xsl:apply-templates select="body/BlogClassList/item"/></div></td>
          </tr>
          </tbody>
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
  <table cellSpacing="0" cellPadding="0" width="778" align="center" border="0">
    <tr>
      <td class="menu_bottombg" align="middle">| <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="../FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="../Copyright.asp" target="_blank">版权申明</A> | 
          <xsl:choose>  
              <xsl:when test="body/Site/ShowAdminLogin = 'enable'">  
              <a class="Bottom" href="../{body/Site/AdminDir}/Admin_Index.asp" target="_blank">管理登录</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
              </xsl:when>
          </xsl:choose>
      </td>
    </tr>
    <tr>
      <td class="bottom_bg" height="80">
        <table cellSpacing="0" cellPadding="0" width="90%" align="center" border="0">
          <tr>
            <td><IMG height="80" src="../Skin/Ocean/bottom_left.gif" width="9" /></td>
            <td align="middle" width="80%"><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/>　站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a></td>
            <td align="right"><IMG height="80" src="../Skin/Ocean/bottom_r.gif" width="9" /></td>
          </tr>
        </table>
      </td>
    </tr>
  </table> 
</body>
</xsl:element>
</xsl:template>


<xsl:template match="Blogitem">
	<tr><td><a href="{link}">
        <xsl:choose>  
            <xsl:when test="string-length(title)  >  20"  >  
            <xsl:value-of select="substring(title,1,20)"  /><xsl:text>...</xsl:text>  
            </xsl:when>
            <xsl:otherwise><xsl:value-of select="title" /></xsl:otherwise>  
        </xsl:choose>  
        </a></td><td width="100" align="right"><xsl:value-of select="BirthDay"/></td></tr>
</xsl:template>

<xsl:template match="body/AddBlog/Blogitem">
	<tr><td width="150"><a href="{link}">
        <xsl:choose>  
            <xsl:when test="string-length(title)  >  10"  >  
            <xsl:value-of select="substring(title,1,10)"  /><xsl:text>...</xsl:text>  
            </xsl:when>
            <xsl:otherwise><xsl:value-of select="title" /></xsl:otherwise>  
        </xsl:choose>  
        </a></td><td width="100" align="right"><xsl:value-of select="author"/></td></tr>
</xsl:template>

<xsl:template  match="/body/EliteBlog/Blogitem">  
    <xsl:apply-templates  select="Blogitem"/>  
</xsl:template> 

<xsl:template match="body/EliteBlog/Blogitem[position() mod 5 = 1]"> 
	<a href="{link}">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="Photo"/></xsl:attribute>
		<xsl:attribute name="alt">名称:<xsl:value-of select="title"/>|日期:<xsl:value-of select="BirthDay"/></xsl:attribute>
		<xsl:attribute name="border">0</xsl:attribute>
		<xsl:attribute name="hight">160</xsl:attribute>
		<xsl:attribute name="width">120</xsl:attribute>
	</xsl:element>
        </a>
        <xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
	<a href="{following-sibling::Blogitem[1]/link}">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="following-sibling::Blogitem[1]/Photo"/></xsl:attribute>
		<xsl:attribute name="alt">名称:<xsl:value-of select="following-sibling::Blogitem[1]/title"/>|日期:<xsl:value-of select="following-sibling::Blogitem[1]/BirthDay"/></xsl:attribute>
		<xsl:attribute name="border">0</xsl:attribute>
		<xsl:attribute name="hight">160</xsl:attribute>
		<xsl:attribute name="width">120</xsl:attribute>
	</xsl:element>
        </a>
        <xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
	<a href="{following-sibling::Blogitem[2]/link}">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="following-sibling::Blogitem[2]/Photo"/></xsl:attribute>
		<xsl:attribute name="alt">名称:<xsl:value-of select="following-sibling::Blogitem[2]/title"/>|日期:<xsl:value-of select="following-sibling::Blogitem[2]/BirthDay"/></xsl:attribute>
		<xsl:attribute name="border">0</xsl:attribute>
		<xsl:attribute name="hight">160</xsl:attribute>
		<xsl:attribute name="width">120</xsl:attribute>
	</xsl:element>
        </a>
        <xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
	<a href="{following-sibling::Blogitem[3]/link}">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="following-sibling::Blogitem[3]/Photo"/></xsl:attribute>
		<xsl:attribute name="alt">名称:<xsl:value-of select="following-sibling::Blogitem[3]/title"/>|日期:<xsl:value-of select="following-sibling::Blogitem[3]/BirthDay"/></xsl:attribute>
		<xsl:attribute name="border">0</xsl:attribute>
		<xsl:attribute name="hight">160</xsl:attribute>
		<xsl:attribute name="width">120</xsl:attribute>
	</xsl:element>
        </a>
        <xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
	<a href="{following-sibling::Blogitem[4]/link}">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="following-sibling::Blogitem[4]/Photo"/></xsl:attribute>
		<xsl:attribute name="alt">名称:<xsl:value-of select="following-sibling::Blogitem[4]/title"/>|日期:<xsl:value-of select="following-sibling::Blogitem[4]/BirthDay"/></xsl:attribute>
		<xsl:attribute name="border">0</xsl:attribute>
		<xsl:attribute name="hight">160</xsl:attribute>
		<xsl:attribute name="width">120</xsl:attribute>
	</xsl:element>
        </a>
        <br />
</xsl:template>

<xsl:template match="body/BlogClassList/item">
	<a href="index.asp?TypeID={id}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text>
</xsl:template>

<xsl:template match="Channelitem">
        <xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{link}" class="channel"><xsl:value-of select="title"/></a>
</xsl:template>

<xsl:template match="Announceitem">
        <a href="{link}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:value-of select="DateAndTime"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
</xsl:template>

</xsl:stylesheet>