<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">
<xsl:element name="html">
<head>
<title><xsl:value-of select="body/MyBlog/BlogName"/></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="index.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<![CDATA[
/* 拖动面板CSS定义 */
div#divTitle {
	width: 564px;
	float: left;
	margin: 2px;
	border: 1px solid #d2d3d9;
}

div#titlebanner {
        color: ffffff;background:url(../../skin/blue/main_title_575.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;padding-left:30;height: 29;
}

div#titletext {
        font-family:宋体;text-align: left;padding-left:5;font-size: 9pt;line-height: 15pt;text-indent: 20px
}

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
        color: ffffff;background:url(../../skin/blue/main_title_575.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;height: 29;
}
]]>
</style>
<script src="../../JS/prototype.js"></script>
<script src="../../JS/scriptaculous.js"></script>
<script src="../../JS/checklogin.js"></script>
<script language="JavaScript">
<![CDATA[
function GetXmlData(iurl,divname,lnum)
{
    var Feedurl = "../../rssfeed.asp";

    var RssReadDOM = new ActiveXObject("Microsoft.XMLDOM");
    RssReadDOM.async=false;

    var p = RssReadDOM.createProcessingInstruction("xml","version=\"1.0\" encoding=\"gb2312\""); 
    //添加文件头 
    RssReadDOM.appendChild(p); 

    //创建根节点
    var objRoot = RssReadDOM.createElement("root");

   //创建子节点
    var objField = RssReadDOM.createNode(1,"listnum",""); 
    objField.text = lnum;
    objRoot.appendChild(objField);

    objField = RssReadDOM.createNode(1,"titlelength",""); 
    objField.text = 35;
    objRoot.appendChild(objField);

    objField = RssReadDOM.createNode(1,"feedurl",""); 
    objField.text = iurl;
    objRoot.appendChild(objField);

    //添加根节点
    RssReadDOM.appendChild(objRoot);

    //查询开始
    var RssFeedHttp = getHTTPObject();
    RssFeedHttp.open("POST",Feedurl,true);
    RssFeedHttp.onreadystatechange = function () 
    {
	if (RssFeedHttp.readyState == 4 && RssFeedHttp.status==200){
            var rstr = "";
            var rssroot = RssFeedHttp.responseXml.getElementsByTagName("item");
            for(i = 0; i < rssroot.length; i++){
                rstr += "<li>";
                rstr += "<a href=\"" + rssroot.item(i).getElementsByTagName("link").item(0).text + "\">" + rssroot.item(i).getElementsByTagName("title").item(0).text + "</a>";
                rstr += "</li>";
            }
            $("item_" + divname).innerHTML=rstr;	
	}else{
            $("item_" + divname).innerHTML="loading...";
        }
    }
    RssFeedHttp.send(RssReadDOM);
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
            <td align="left">
            <xsl:choose>  
                <xsl:when test="body/Site/EnableRss = 'enable'">  
                <a href="../../Rss.asp" Title="Rss 2.0" Target="_blank"><img src="../../images/rss.gif" border="0" /></a>
                </xsl:when>
            </xsl:choose>
            <xsl:choose>  
                <xsl:when test="body/Site/EnableWap = 'enable'">  
                <img src="../../images/Wap.gif" border="0" alt="WAP浏览支持" style="cursor:hand;"  onClick="window.open('../../Wap.asp?ReadMe=Yes', 'Wap', 'width=160,height=257,resizable=0,scrollbars=no');" />
                </xsl:when>
            </xsl:choose>
            </td>
            <td align="right">
            <xsl:choose>  
                <xsl:when test="body/Site/ShowSiteChannel = 'enable'">  
                |<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{body/Site/SiteUrl}" class="channel">网站首页</a><xsl:apply-templates select="body/ChannelList/Channelitem"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
                </xsl:when>
            </xsl:choose>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td align="middle">
        <xsl:choose>  
            <xsl:when test="string-length(body/Site/SiteLogo) > 0"  >  
            <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/SiteLogo}" width="180" height="60" border="0" /></a>
            </xsl:when>
        </xsl:choose>
        <xsl:choose>  
            <xsl:when test="string-length(body/Site/BannerUrl) > 0"  >  
            <a><xsl:attribute name="href"><xsl:value-of select="body/Site/BannerUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/BannerUrl}" width="580" height="60" border="0" /></a>
            </xsl:when>
        </xsl:choose> 
      </td>
      <td align="middle"></td>
    </tr>
    <tr>
      <td align="middle" colSpan="2">
      <!--导航、日期代码开始-->
        <table class="top_nav_menu" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
          <tr>
            <td align="middle" width="50"><IMG src="../../Images/arrow3.gif" align="absMiddle" /></td>
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
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('login')">
            <td><IMG src="../../skin/blue/login_01.gif" /></td>
          </tr>
          <tbody id="login">
          <tr>
            <td vAlign="center" align="middle" background="../../skin/blue/login_02.gif" height="151"><div id="UserLogin">载入中...<script language="JavaScript" type="text/JavaScript">LoadUserLogin("../../",0,0);</script></div></td>
          </tr>
          </tbody>
          <tr>
            <td><IMG src="../../skin/blue/login_03.gif" /></td>
          </tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('plinfo')">
            <td class="left_title" align="middle">空间档案</td>
          </tr>
          <tbody id="plinfo">
          <tr>
            <td class="left_tdbg1" vAlign="top" height="179">
                <center><img src="{body/MyBlog/Photo}" width="150" height="160" border="1"></img></center>
                <li>空间主人:<xsl:value-of select="body/MyBlog/UserName"/></li>
                <li>创建日期:<xsl:value-of select="body/MyBlog/BirthDay"/></li>
            </td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('plmenu')">
            <td class="left_title" align="middle">快速连接</td>
          </tr>
          <tbody id="plmenu">
          <tr>
            <td class="left_tdbg1" vAlign="top">
                <center><a href="../index.asp">= 更多空间 =</a></center>
            </td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
<xsl:choose>  
<xsl:when test="body/MyLink/linkitem != ''">  
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('pllink')">
            <td class="left_title" align="middle">关注连接</td>
          </tr>
          <tbody id="pllink">
          <tr>
            <td class="left_tdbg1" vAlign="top">
<xsl:for-each select="body/MyLink/linkitem">
                <li><a href="{link}"><xsl:value-of select="title"/></a></li>
</xsl:for-each>
            </td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
</xsl:when>
</xsl:choose>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('plfangke')">
            <td class="left_title" align="middle">最近访客</td>
          </tr>
          <tbody id="plfangke">
          <tr>
            <td class="left_tdbg1" vAlign="top" height="50">
<xsl:for-each select="body/NewVisitor/visitor">
                <li><a href="../{username}{userid}/"><xsl:value-of select="username"/>(<xsl:value-of select="num"/>)</a></li>
</xsl:for-each></td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
</td>
      <td width="5"></td>
      <td vAlign="top">
      <!--显示我的 Blog栏目-->
<div id="divTitle">
        <div id="titlebanner" onclick="new Element.toggle('titletext')"><table width="100%"><tr valign="middle"><td><font color="white">简介</font></td></tr></table></div>
        <div id="titletext"><xsl:value-of select="body/MyBlog/BlogIntro" disable-output-escaping="yes"/></div>
</div>
<div id="divContainer">
<xsl:for-each select="body/MyBlog/Blogitem">
    <xsl:if test="type='rss' or type='diary' or type='music' or type='book' or type='photo'">
        <div id="itembody">
            <div id="itemtitle"><table width="100%"><tr valign="middle"><td width="30"><img src="../../images/jiaodian_biao.gif" border="0" onclick="new Element.toggle('item_{title}')"/></td><td><font color="white"><xsl:value-of select="title"/></font></td><td align="right">
            <xsl:choose>
                <xsl:when test="type = 'diary'"><a href="../Showdiary.asp?BlogID={ClassID}" target="_blank"><img src="../images/diarylist.gif" border="0" align="absmiddle" alt="查看全部"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;</xsl:text></xsl:when>
                <xsl:when test="type = 'music'"><a href="../Showmusic.asp?BlogID={ClassID}" target="_blank"><img src="../images/musiclist.gif" border="0" align="absmiddle" alt="查看全部"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;</xsl:text></xsl:when>
                <xsl:when test="type = 'book'"><a href="../Showbook.asp?BlogID={ClassID}" target="_blank"><img src="../images/booklist.gif" border="0" align="absmiddle" alt="查看全部"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;</xsl:text></xsl:when>
                <xsl:when test="type = 'photo'"><a href="../Showphoto.asp?BlogID={ClassID}" target="_blank"><img src="../images/photolist.gif" border="0" align="absmiddle" alt="查看全部"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;</xsl:text></xsl:when>
                <xsl:when test="type = 'link'"><a href="../Showlink.asp?BlogID={ClassID}" target="_blank"><img src="../images/linklist.gif" border="0" align="absmiddle" alt="查看全部"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;</xsl:text></xsl:when>
            </xsl:choose>
            <a href="{link}" target="_blank"><img src="../../images/Rss.gif" border="0" align="absmiddle"/></a>
            </td></tr></table></div>
            <div id="item_{title}">loading...</div>
        </div></xsl:if>
</xsl:for-each>
<xsl:for-each select="body/MyBlog/Blogitem">
    <xsl:if test="type=1 or type=2 or type=3">
        <div id="itembody">
          <div id="itemtitle"><table width="100%"><tr valign="middle"><td width="30"><img src="../../images/jiaodian_biao.gif" border="0" onclick="new Element.toggle('item_{title}')"/></td><td><font color="white"><xsl:value-of select="title"/></font></td><td align="right">
          <a href="{link}" target="_blank"><img src="../../images/Rss.gif" border="0" align="absmiddle"/></a>
          </td></tr></table></div>
          <div id="item_{title}">loading...</div>
    </div></xsl:if>
</xsl:for-each>
</div>
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
      <td class="Bottom_Adminlogo" colSpan="2"> | <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="../../FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="../../Copyright.asp" target="_blank">版权申明</A> | 
          <xsl:choose>  
              <xsl:when test="body/Site/ShowAdminLogin = 'enable'">  
              <a class="Bottom" href="../../{body/Site/AdminDir}/Admin_Index.asp" target="_blank">管理登录</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
              </xsl:when>
          </xsl:choose>
      </td>
    </tr>
    <tr class="Bottom_Copyright">
      <td width="20%"><IMG src="../../images/logo.gif" /></td>
      <td width="80%" align="center"> 站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a><br /><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/></td>
    </tr>
  </table>
<script type="text/javascript">
    Sortable.create('divContainer',{tag:'div',overlap:'horizontal',constraint:false});
    setTimeout("addfangke(<xsl:value-of select="body/MyBlog/BlogID"/>,0)",6000);
    <xsl:for-each select="body/MyBlog/Blogitem">
        <xsl:if test="type='rss' or type='diary' or type='music' or type='book' or type='photo' or type=1 or type=2 or type=3">
        GetXmlData('<xsl:value-of select="link"/>','<xsl:value-of select="title"/>',<xsl:value-of select="listnum"/>);
        </xsl:if>
    </xsl:for-each>
</script>
</body>
</xsl:element>
</xsl:template>

<xsl:template match="Channelitem">
        <xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{link}" class="channel"><xsl:value-of select="title"/></a>
</xsl:template>

<xsl:template match="Announceitem">
        <a href="../../{link}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:value-of select="DateAndTime"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
</xsl:template>

</xsl:stylesheet>