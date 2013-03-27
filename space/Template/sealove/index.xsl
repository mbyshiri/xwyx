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
	width: 780px;
	float: left;
	margin: 2px;
	border: 1px solid #d2d3d9;
}

div#titlebanner {
        color: ffffff;background:url(../../skin/sealove/Top_01BG.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;padding-left:30;height: 29;
}

div#titletext {
        font-family:宋体;text-align: left;padding-left:5;font-size: 9pt;line-height: 15pt;text-indent: 20px
}

div#divContainer {
	width: 100%;
}

div#itembody {
	float: left;
        width: 388px;

	margin: 2px;
	border: 1px solid #d2d3d9;
}

div#itemtitle {
        color: ffffff;background:url(../../skin/sealove/Top_01BG.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;height: 29;
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
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td colspan="3" height="6"><img src="../../Skin/sealove/space.gif" /></td></tr>
	<tr><td width="14" height="34"><img src="../../Skin/sealove/Top_01Left.gif" /></td>
	    <td background="../../Skin/sealove/Top_01BG.gif" align="right" style="color:#FFFFFF">
            <xsl:choose>  
                <xsl:when test="body/Site/ShowSiteChannel = 'enable'">  
                |<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{body/Site/SiteUrl}" class="channel">网站首页</a><xsl:apply-templates select="body/ChannelList/Channelitem"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
                </xsl:when>
            </xsl:choose>
            </td>
		<td width="14"><img src="../../Skin/sealove/Top_01Right.gif" /></td>
	</tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="4" background="../../Skin/sealove/Top_02Left.gif"><img src="../../Skin/sealove/space.gif" /></td>
		<td background="../../Skin/sealove/Top_02BG.gif">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="280" height="90" align="center">
        <xsl:choose>  
            <xsl:when test="string-length(body/Site/SiteLogo) > 0"  >  
            <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/SiteLogo}" border="0" /></a>
            </xsl:when>
        </xsl:choose>
                            </td>
			    <td align="center">
        <xsl:choose>  
            <xsl:when test="string-length(body/Site/BannerUrl) > 0"  >  
            <a><xsl:attribute name="href"><xsl:value-of select="body/Site/BannerUrl"/></xsl:attribute><xsl:attribute name="Title"><xsl:value-of select="body/Site/SiteName"/></xsl:attribute><img src="{body/Site/BannerUrl}" width="580" height="60" border="0" /></a>
            </xsl:when>
        </xsl:choose> 
                            </td>
			</tr>
		</table>
		</td>
		<td width="4" background="../../Skin/sealove/Top_02Right.gif"><img src="../../Skin/sealove/space.gif" /></td>
	</tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="17" background="../../Skin/sealove/Top_03Left.gif"><img src="../../Skin/sealove/space.gif" /></td>
		<td background="../../Skin/sealove/Top_03BG.gif">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="75" height="33">
            <xsl:choose>  
                <xsl:when test="body/Site/EnableRss = 'enable'">  
                <a href="../../Rss.asp" Title="Rss 2.0" Target="_blank"><img src="../../images/rss.gif" border="0" /></a>
                </xsl:when>
            </xsl:choose>
                            </td>
				<td width="20"><img src="../../Skin/sealove/icon01.gif" /></td>
				<td width="60">最新公告：</td>
				<td width="400"><MARQUEE onmouseover="this.stop()" onmouseout="this.start()" scrollAmount="1" scrollDelay="4" width="400" align="left"><p><xsl:apply-templates select="body/AnnounceList/Announceitem"/></p></MARQUEE></td>
				<td align="right"></td>				
			</tr>
		</table>
		</td>
		<td width="17" background="../../Skin/sealove/Top_03Right.gif"><img src="../../Skin/sealove/space.gif" /></td>
	</tr>
	<tr><td colspan="3" height="5"><img src="../../Skin/sealove/space.gif" /></td></tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="15"><img src="../../Skin/sealove/Main_TopLeft.gif" /></td>
		<td height="11" background="../../Skin/sealove/Main_TopBG.gif"><img src="../../Skin/sealove/space.gif" /></td>
		<td width="15"><img src="../../Skin/sealove/Main_TopRight.gif" /></td>
	</tr>
</table>

<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td width="8" background="../../Skin/sealove/Main_Left.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="135" height="60"><img src="../../Skin/sealove/Main_Search.gif" alt="站内搜索" /></td>
				<td><table cellSpacing="0" cellPadding="0" border="0">
					<FORM name="search" action="../../search.asp" method="post">
					<tr><td align="middle"><Input id="Keyword" maxLength="50" value="关键字" name="Keyword" /></td>
						<td align="center" width="55"><input name="Submit" id="Submit" type="image" src="../../Skin/sealove/Icon_Search.gif" style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" /></td>
						<td align="middle">
				<input type="radio" value="Article" name="ModuleName" checked="True" /> 文章
				<Input type="radio" value="Soft" name="ModuleName" /> 下载
				<Input type="radio" value="Photo" name="ModuleName" /> 图片
				<Input id="Field" type="hidden" value="Title" name="Field" /></td>
					</tr>
					</FORM>
					</table>
				</td>
				<td width="166" align="right"><img src="../../Skin/sealove/Main_girl01.gif" /></td>
			</tr>
		</table>


		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td valign="top">
				<table width="98%" align="center" border="0" cellpadding="0" cellspacing="0" background="../../Skin/sealove/Path_BG.gif">
					<tr><td width="9"><img src="../../Skin/sealove/Path_Left.gif" /></td>
						<td width="20"><img src="../../Skin/sealove/icon02.gif" /></td>
						<td>您现在的位置：<a class="LinkPath"><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> <xsl:value-of select="body/MyBlog/BlogName"/></td>
						<td width="84"><a href="/Reg/User_Reg.asp" target="_blank"><img src="../../Skin/sealove/Button_Reg.gif" alt="会员注册" border="0" /></a></td>
						<td width="9"><img src="../../Skin/sealove/Path_Right.gif" /></td>
					</tr>
				</table>
				</td>
				<td width="92" height="48" align="right" valign="top"><img src="../../Skin/sealove/Main_girl02.gif" /></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td height="2" bgcolor="#0099BB"></td></tr>
			<tr><td height="1"></td></tr>
			<tr><td height="1" bgcolor="#0099BB"></td></tr>
			<tr><td height="2"></td></tr>
			<tr><td background="../../Skin/sealove/AD02.gif" align="right">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                           <td><img src="../../Skin/sealove/AD01.gif" border="0" alt="通档广告位：请自行修改为JS调用代码" /></td>
                           <td align="right"><img src="../../Skin/sealove/AD08.gif" border="0" /></td>
                        </tr>
                        </table>			  
			</td>
			</tr>
			<tr><td height="8"></td></tr>
		</table>
	</td>
	<td width="8" background="../../Skin/sealove/Main_Right.gif"></td>
	</tr>
</table>
  <!-- ********网页顶部代码结束******** -->

  <!-- ********网页中部代码开始******** -->
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td width="8" background="../../Skin/sealove/Main_Left.gif"></td>
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
          <a href="{link}" target="_blank"><img src="../../images/Rss.gif" border="0" align="absmiddle" /></a>
          </td></tr></table></div>
          <div id="item_{title}">loading...</div>
    </div></xsl:if>
</xsl:for-each>
</div>
      </td>

      <td width="5"></td>
	<td width="180" valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr onClick="new Element.toggle('login')"><td><img src="../../Skin/sealove/Login_Top.gif" alt="会员登录" /></td></tr>
			<tr><td background="../../Skin/sealove/Login_BG1.gif">
				<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" style="background-image: url(../../Skin/sealove/Login_BG.gif); background-repeat: no-repeat; background-position: center top">
					<tr><td height="6"></td></tr>
					<tr><td><div id="UserLogin">载入中...<script language="JavaScript" type="text/JavaScript">LoadUserLogin("../../",0,0);</script></div></td></tr>
					<tr><td height="6"></td></tr>
				</table>
			</td></tr>
			<tr><td><img src="../../Skin/sealove/Room_bottom.gif" /></td></tr>
		</table><br />


		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../../Skin/sealove/Column01_L.gif" /></td>
				<td background="../../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>空间档案</b></td>
				<td width="10"><img src="../../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="150">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center">
		<center><img src="{body/MyBlog/Photo}" width="150" height="160" border="1"></img></center>
                <li>空间主人:<xsl:value-of select="body/MyBlog/UserName"/></li>
                <li>创建日期:<xsl:value-of select="body/MyBlog/BirthDay"/></li>
                                        </td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../../Skin/sealove/space.gif" /></td></tr>
		</table><br />

		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../../Skin/sealove/Column01_L.gif" /></td>
				<td background="../../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>快速连接</b></td>
				<td width="10"><img src="../../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center"><center><a href="../index.asp">= 更多空间 =</a></center></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../../Skin/sealove/space.gif" /></td></tr>
		</table><br />
<xsl:choose>  
<xsl:when test="body/MyLink/linkitem != ''">  
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../../Skin/sealove/Column01_L.gif" /></td>
				<td background="../../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>关注连接</b></td>
				<td width="10"><img src="../../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center">
				<xsl:for-each select="body/MyLink/linkitem">
                				<li><a href="{link}"><xsl:value-of select="title"/></a></li>
				</xsl:for-each></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../../Skin/sealove/space.gif" /></td></tr>
		</table><br />
</xsl:when>
</xsl:choose>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../../Skin/sealove/Column01_L.gif" /></td>
				<td background="../../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>最近访客</b></td>
				<td width="10"><img src="../../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center">
				<xsl:for-each select="body/NewVisitor/visitor">
              				  <li><a href="../{username}{userid}/"><xsl:value-of select="username"/>(<xsl:value-of select="num"/>)</a></li>
				</xsl:for-each></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../../Skin/sealove/space.gif" /></td></tr>
		</table>
</td>
<td width="5"></td>
<td width="8" background="/Skin/sealove/Main_Right.gif"></td>
</tr>
</table>
  <!-- ********网页中部代码结束******** -->
  <!-- ********网页底部代码开始******** -->
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td width="8" background="../../Skin/sealove/Main_Left.gif"></td>
		<td valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td height="1" bgcolor="#0099BB"></td></tr>
			<tr><td height="1"></td></tr>
			<tr><td height="9" bgcolor="#0099BB"></td>
			</tr>
			<tr><td height="1"></td></tr>
			<tr><td height="1" bgcolor="#0099BB"></td></tr>
		</table>
		<table width="90%" align="center" border="0" cellpadding="0" cellspacing="0">
			<tr><td height="25" align="center" style="color:#000000"> | <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="../../FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="../../Copyright.asp" target="_blank">版权申明</A> | 
          <xsl:choose>  
              <xsl:when test="body/Site/ShowAdminLogin = 'enable'">  
              <a class="Bottom" href="../../{body/Site/AdminDir}/Admin_Index.asp" target="_blank">管理登录</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
              </xsl:when>
          </xsl:choose>
          </td></tr>
	  <tr><td height="1" background="../../Skin/sealove/line01.gif"></td></tr>
	  </table>
	  <table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="210" align="center"><a href="http://www.asp163.net" target="_blank"><img src="../../Skin/sealove/PElogo_sealove.gif" border="0" alt="动易网络" /></a></td>
				<td> 站长：<A href="mailto:info@powereasy.net"></A><br />
				  模板设计：<a href="http://www.mz25.net/" target="_blank">梅子</a></td>
				<td width="20"></td>
				<td width="120" height="80">
					<a href="http://www.miibeian.gov.cn" target="_blank">
					<img src="../../Skin/sealove/mii.gif" border="0" alt="信息产业部备案" /><br />*ICP备********号</a></td>
			</tr>
	  </table>
	  </td><td width="8" background="../../Skin/sealove/Main_Right.gif"></td></tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="15"><img src="../../Skin/sealove/Main_BottomLeft.gif" /></td>
		<td height="11" background="../../Skin/sealove/Main_BottomBG.gif"><img src="../../Skin/sealove/space.gif" /></td>
		<td width="15"><img src="../../Skin/sealove/Main_BottomRight.gif" /></td>
	</tr>
	<tr><td colspan="3" height="5"></td></tr>
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