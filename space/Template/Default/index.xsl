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
        color: ffffff;background:url(../../skin/Ocean/main_bs1.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;padding-left:30;height: 29;
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
        color: ffffff;background:url(../../skin/Ocean/main_bs1.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;height: 29;
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
  <table height="114" cellSpacing="0" cellPadding="0" width="778" align="center" background="../../Skin/Ocean/top_bg.jpg" border="0">
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
                  <td><IMG height="25" src="../../Skin/Ocean/Announce_01.jpg" width="68" /></td>
                  <td class="showa" width="280" background="../../Skin/Ocean/Announce_02.jpg">
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
              <table height="89" cellSpacing="0" cellPadding="0" width="94" background="../../Skin/Ocean/topr.gif" border="0">
                <tr>
                  <td align="middle" colSpan="2">
                    <table height="56" cellSpacing="0" cellPadding="0" width="79" border="0">
                      <tr>
                        <td align="middle" width="26"><IMG height="13" src="../../Skin/Ocean/arrows.gif" width="13" /></td>
                        <td width="68"><a class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</a></td>
                      </tr>
                      <tr>
                        <td align="middle"><IMG height="13" src="../../Skin/Ocean/arrows.gif" width="13" /></td>
                        <td><a class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" href="#">设为首页</a></td>
                      </tr>
                      <tr>
                        <td align="middle"><IMG height="13" src="../../Skin/Ocean/arrows.gif" width="13" /></td>
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
                <a href="../../Rss.asp" Title="Rss 2.0" Target="_blank" onmouseover="showmenu(event,'0','rss','http://localhost/Rss.asp');" onmouseout="delayhidemenu();"><img src="../../images/rss.gif" border="0" /></a>
                </xsl:when>
            </xsl:choose>
     </td>
    </tr>
    <tr>
      <td><IMG height="7" src="../../Skin/Ocean/menu_bg2.jpg" width="778" /></td>
    </tr>
    <tr>
      <td background="../../Skin/Ocean/addr.jpg" height="21">
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td align="middle" width="5%"><IMG height="17" src="../../Skin/Ocean/arrows2.gif" width="16" /></td>
            <td width="95%">您现在的位置： <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> 聚合空间首页</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td background="../../Skin/Ocean/addr_line.jpg" height="4"></td>
    </tr>
  </table> 
  <!-- ********网页顶部代码结束******** -->
  <!-- ********网页中部代码开始******** -->
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td vAlign="top" width="180">
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onClick="new Element.toggle('login')">
            <td><IMG src="../../Skin/Ocean/login_01.jpg" /></td>
          </tr>
          <tbody id="login">
          <tr>
            <td vAlign="center" align="middle" background="../../Skin/Ocean/login_02.gif" style="padding:1px">
            <div id="UserLogin">载入中<script language="JavaScript" type="text/JavaScript">LoadUserLogin("../../",0,0);</script></div></td>
          </tr>
          </tbody>
          <tr>
            <td><IMG src="../../Skin/Ocean/login_03.jpg" /></td>
          </tr>
        <tr> 
          <td Class="main_shadow"></td> 
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
            <div id="itemtitle"><table width="100%"><tr valign="middle" onclick="new Element.toggle('item_{title}')"><td>　　<font color="white"><xsl:value-of select="title"/></font></td><td align="right">
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
          <div id="itemtitle"><table width="100%"><tr valign="middle" onclick="new Element.toggle('item_{title}')"><td>　　<font color="white"><xsl:value-of select="title"/></font></td><td align="right">
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
  <table cellSpacing="0" cellPadding="0" width="778" align="center" border="0">
    <tr>
      <td class="menu_bottombg" align="middle">| <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="../../FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="../../Copyright.asp" target="_blank">版权申明</A> | 
          <xsl:choose>  
              <xsl:when test="body/Site/ShowAdminLogin = 'enable'">  
              <a class="Bottom" href="../../{body/Site/AdminDir}/Admin_Index.asp" target="_blank">管理登录</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
              </xsl:when>
          </xsl:choose>
      </td>
    </tr>
    <tr>
      <td class="bottom_bg" height="80">
        <table cellSpacing="0" cellPadding="0" width="90%" align="center" border="0">
          <tr>
            <td><IMG height="80" src="../../Skin/Ocean/bottom_left.gif" width="9" /></td>
            <td align="middle" width="80%"><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/>　站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a></td>
            <td align="right"><IMG height="80" src="../../Skin/Ocean/bottom_r.gif" width="9" /></td>
          </tr>
        </table>
      </td>
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