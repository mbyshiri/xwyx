<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">
<xsl:element name="html">
<head>
<title><xsl:value-of select="body/MyBlog/BlogName"/></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="BlueSkin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<![CDATA[
/* 拖动面板CSS定义 */
div#diarycontent {
	width: 100%;
	float: left;
	margin: 0px;
}
div#diarybody {
	width: 564px;
	float: left;
	margin: 2px;
	border: 1px solid #d2d3d9;
}

div#diarytitle {
        color: ffffff;background:url(../skin/blue/main_title_575.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;padding-left:30;height: 29;
}

div#diarytext {
        font-family:宋体;text-align: center;padding-left:5;font-size: 9pt;line-height: 15pt;text-indent: 20px
}
div#diaryfoot {
        color: 0000FF;background:#d2d3d9;text-align: right;padding-right:20;height: 20;
}

div#commentbody {
        margin: 10px;
        border: 1px solid #ffffff;
}
div#commenttitle {
        text-align: left;padding-left:10;height: 10;
}
div#commentcontent {
        border: 1px solid #d2d3d9;text-align: left;padding-left:10;height: 20;text-indent: 20px
}
div#showpage {
	width: 564px;
	float: left;
	margin: 2px;
	border: 1px solid #d2d3d9;
        text-align: right;
        padding-right:5;
}
]]>
</style>
<script src="../JS/prototype.js"></script>
<script src="../JS/scriptaculous.js"></script>
<script src="../JS/checklogin.js"></script>
<script language="JavaScript">
<![CDATA[
var tempdivcontent;
var isroot;
var s1;
var s2;
var s3;
var s4;
var s5;

function initPage(totalpage,current,totalitem,iBid,iUid,iroot)
{
    s1 = totalpage;
    s2 = current;
    s3 = totalitem;
    s4 = iBid;
    s5 = iUid;
    isroot = iroot;
    if(isroot==1){
        ShowPage(s1,s2,s3,s4,isroot);
    }else{
        ShowPage(s1,s2,s3,s5,isroot);
    }
}

function ShowPage(totalpage,current,totalitem,iBid,iroot)
{
    if(totalpage>1){
        var temppage = "<b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + ",1," + totalitem + "," + iBid + ",1);\">首页</b>";
        if(current>1){
            temppage += " <b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + "," + (current-1) + "," + totalitem + "," + iBid + ",1);\"><</b> ";
        }
        for (var i = 1; i <= totalpage; i++) {
            if(i==current){
                temppage += " [<font color=\"red\"><b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + "," + i + "," + totalitem + "," + iBid + ",1);\">" + i + "</b></font>] ";
            }else{
                temppage += " <b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + "," + i + "," + totalitem + "," + iBid + ",1);\">" + i + "</b> ";
            }
        }
        if(current<totalpage){
            temppage += " <b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + "," + (current+1) + "," + totalitem + "," + iBid + ",1);\">></b> ";
        }
        if(totalpage>1){
            temppage += " <b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + "," + totalpage + "," + totalitem + "," + iBid + ",1);\">尾页</b> 共[" + totalpage + "]页 ";
        }else{
            temppage += " 尾页</b> 共[" + totalpage + "]页 ";
        }
        $('showpage').innerHTML = temppage;
    }else{
        Element.hide('showpage');
    }
}

function ChangePage(iTotalPage,iPage,iTotal,Bid2,iType)
{
    s2 = iPage;
    $('diarycontent').innerHTML = "数据更新中...";
    var url = "Showphoto.asp";
    if(iType==1){
        var pars = "BlogID=" + Bid2 + "&page=" + iPage;
    }else{
        var pars = "ID=" + Bid2;
    }
    var myAjax = new Ajax.Request(url, {method: 'get', parameters: pars, onComplete: PageResponse});
    ShowPage(iTotalPage,iPage,iTotal,Bid2,isroot);
}

function PageResponse(originalRequest)
{
    var tempstr;
    tempstr = "";
　　var xml = new ActiveXObject("Microsoft.XMLDOM");
　　xml.async = false;
　　xml.load(originalRequest.responseXml);
    var root = xml.getElementsByTagName("Diary");
    for(i = 0; i < root.length; i++){
        tempstr += "<div id=\"diarybody\">";
        tempstr += "<div id=\"diarytitle\"><table width=\"100%\"><tr valign=\"middle\"><td><font color=\"#ffffff\">" + root.item(i).getElementsByTagName("Title").item(0).text + "</font></td><td align=\"right\">浏览<font color=\"red\">" + root.item(i).getElementsByTagName("Hits").item(0).text + "</font>次</td></tr></table></div>";
        tempstr += "<div id=\"diarytext\"><img src=\"" + root.item(i).getElementsByTagName("Content").item(0).text + "\" width=\"500\"/></div>";
        tempstr += "<div id=\"diaryfoot\">[<b style=\"cursor:hand;\" onclick=\"new Element.toggle('comment_" + root.item(i).getElementsByTagName("Title").item(0).text + "')\">查看评论</b>(<font color=\"red\">" + root.item(i).getElementsByTagName("Comment").item(0).text + "</font>)][<b style=\"cursor:hand;\" onclick=\"showComment(" + root.item(i).getElementsByTagName("ID").item(0).text
        tempstr += "," + xml.getElementsByTagName("MyBlog/BlogID").item(0).text
        tempstr += ");\">发表评论</b>] [发布时间" + root.item(i).getElementsByTagName("Datetime").item(0).text + "]</div>";
        tempstr += "<div id=\"comment_" + root.item(i).getElementsByTagName("Title").item(0).text + "\" style=\"display:none\">";
            var commentstr = root.item(i).getElementsByTagName("CommentList");
            for(j = 0; j < commentstr.length; j++){
                tempstr += "<div id=\"commentbody\">";
                tempstr += "<div id=\"commenttitle\">" + commentstr.item(j).getElementsByTagName("name").item(0).text + "在" + commentstr.item(j).getElementsByTagName("datetime").item(0).text + "评论说:<b>" + commentstr.item(j).getElementsByTagName("title").item(0).text + "</b></div>";
                tempstr += "<div id=\"commentcontent\">" + commentstr.item(j).getElementsByTagName("content").item(0).text + "</div>";
                tempstr += "</div>";
            }
        tempstr += "</div>";
        tempstr += "</div>";
    }
    $('diarycontent').innerHTML = tempstr;
}

function showComment(itemID,blogid)
{
    tempdivcontent = $('diarycontent').innerHTML;
    var templ = "<div id=\"diarybody\">";
    templ = templ + "<div id=\"diarytitle\">发表评论</div>";
    templ = templ + "<div id=\"pltext\">";
    templ = templ + "<input name=\"plname\" type=\"hidden\" value=\"" + username + "\">"
    templ = templ + "<input name=\"plpass\" type=\"hidden\" value=\"" + userpass + "\">";
    if(userstat=='login'){
        templ = templ + "匿名评论 <input type=\"checkbox\" name=\"noname\" value='1'><br />";
    }else{
        templ = templ + "匿名评论 <input type=\"checkbox\" name=\"noname\" value='1' Checked Disabled><br />";
    }
    templ = templ + "评论标题 <input name=\"pltitle\" id=\"pltitle\" type=\"text\"><br />";
    templ = templ + "评论内容 <textarea name=\"plcontent\" cols=\"50\" rows=\"4\"></textarea><br />";
    templ = templ + "<input name=\"plid\" id=\"plid\" type=\"hidden\" value=" + itemID + ">";
    templ = templ + "<input name=\"blogid\" id=\"blogid\" type=\"hidden\" value=" + blogid + ">";
    templ = templ + "<center><a href=\"#\" onclick=\"SaveComment();\">保存</a> <a href=\"#\" onclick=\"CancelComment();\">取消</a></center></div>";
    templ = templ + "</div>";
    $('diarycontent').innerHTML = templ;
    Field.focus('pltitle');
}

function CancelComment()
{
    $('diarycontent').innerHTML = tempdivcontent;
    Sortable.create('diarycontent',{tag:'div'});
}

function SaveComment()
{
    var saveurl = "Showphoto.asp?Action=savepl";
    var name = $F('plname');
    var noname = $F('noname');
    var plpass = $F('plpass');
    var title = $F('pltitle');
    var content = $F('plcontent').stripTags();
    var pid = $F('plid');
    var blogid = $F('blogid');
    //$('pltext').innerHTML = "<center>保存数据中...</center>";
    if((noname!='1')&&(name=='')&&(userstat=='login')){
        alert("您尚未登录,请选择匿名发表!");
    }else{
        if((noname!='1')&(plpass=='')&&(userstat=='login')){
            alert("您尚未登录,请选择匿名发表!");
        }else{
            if(title==''){
                 alert("标题不能为空!");
                 Field.focus('pltitle');
            }else{
                if(content==''){
                     alert("内容不能为空!");
                     Field.focus('plcontent');
                }else{
                    // 创建返回信息XML文档
                    var checkurl = "Showphoto.asp?Action=savepl";

                    var pl_dom = new ActiveXObject("Microsoft.XMLDOM");
                    pl_dom.async=false;

                    var p = pl_dom.createProcessingInstruction("xml","version=\"1.0\" encoding=\"gb2312\""); 
                    //添加文件头 
                    pl_dom.appendChild(p); 

                    //创建根节点
                    var objRoot = pl_dom.createElement("root");

                    //创建子节点
                    var objField = pl_dom.createNode(1,"username",""); 
                    objField.text = name;
                    objRoot.appendChild(objField);

                    objField = pl_dom.createNode(1,"password",""); 
                    objField.text = plpass;
                    objRoot.appendChild(objField);

                    objField = pl_dom.createNode(1,"noname",""); 
                    if(noname!='1'){
                        objField.text = 0;
                    }else{
                        objField.text = 1;
                    }
                    objRoot.appendChild(objField);

                    objField = pl_dom.createNode(1,"title",""); 
                    objField.text = title;
                    objRoot.appendChild(objField);
                    objField = pl_dom.createNode(1,"content",""); 
                    objField.text = content;
                    objRoot.appendChild(objField);
                    objField = pl_dom.createNode(1,"type",""); 
                    objField.text = 6;
                    objRoot.appendChild(objField);
                    objField = pl_dom.createNode(1,"id",""); 
                    objField.text = pid;
                    objRoot.appendChild(objField);
                    objField = pl_dom.createNode(1,"blogid",""); 
                    objField.text = blogid;
                    objRoot.appendChild(objField);

                    //添加根节点
                    pl_dom.appendChild(objRoot);

                    // 把XML文档发送到Web服务器
                    var plhttp = getHTTPObject();
                    plhttp.open("POST",checkurl,false);
                    plhttp.send(pl_dom);
                    // 显示服务器返回的信息
                    if(plhttp.readyState == 4 && plhttp.status==200){
                        CommentReponse(plhttp);
                    }else{
                        CancelComment();
                    }
                }
            }
        }
    }
}

function CommentReponse(backRequest)
{
    var tempstr;
    empstr = "";
　　var xml = new ActiveXObject("Microsoft.XMLDOM");
　　xml.async = false;
　　xml.load(backRequest.responseXml);
    //$('diarycontent').innerHTML = backRequest.responseText;
    var root = xml.getElementsByTagName("body/serverbackinfo");
    //$('diarycontent').innerHTML = root.item(0).getElementsByTagName("infomation").item(0).text;
    //$('diarycontent').innerHTML = tempdivcontent;
    if(isroot=='1'){ 
        ChangePage(s1,s2,s3,s4,1);
    }else{
        ChangePage(s1,s2,s3,s5,0);
    }
    Sortable.create('diarycontent',{tag:'div'});
}
]]>
</script>
</head>
<body leftmargin="0" topmargin="0" onload="initPage('{body/MyBlog/TotalPage}','{body/MyBlog/CurrentPage}','{body/MyBlog/totalPut}','{body/MyBlog/BlogID}','{body/MyBlog/Diary/ID}','{body/MyBlog/IsRoot}');">
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
                <a href="../Rss.asp" Title="Rss 2.0" Target="_blank"><img src="../images/rss.gif" border="0" /></a>
                </xsl:when>
            </xsl:choose>
            <xsl:choose>  
                <xsl:when test="body/Site/EnableWap = 'enable'">  
                <img src="../images/Wap.gif" border="0" alt="WAP浏览支持" style="cursor:hand;"  onClick="window.open('../Wap.asp?ReadMe=Yes', 'Wap', 'width=160,height=257,resizable=0,scrollbars=no');" />
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
            <td align="middle" width="50"><IMG src="../Images/arrow3.gif" align="absMiddle" /></td>
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
          <tr onclick="new Element.toggle('login')">
            <td><IMG src="../skin/blue/login_01.gif" /></td>
          </tr>
          <tbody id="login">
          <tr>
            <td vAlign="center" align="middle" background="../skin/blue/login_02.gif" height="151"><div id="UserLogin">载入中...<script language="JavaScript" type="text/JavaScript">LoadUserLogin("../",0,0);</script></div></td>
          </tr>
          </tbody>
          <tr>
            <td><IMG src="../skin/blue/login_03.gif" /></td>
          </tr>
        </table>
      <!--用户登录代码结束-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('plopti')">
            <td class="left_title" align="middle">操 作 列 表</td>
          </tr>
          <tbody id="plopti">
          <tr>
            <td class="left_tdbg1" align="middle"><a href="Showphoto.asp?BlogID={body/MyBlog/BlogID}">= 全部照片 =</a><br /><a href="{body/MyBlog/BlogDir}/">= 返回首页 =</a><br /><a href="index.asp">= 空间列表 =</a></td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('plinfo')">
            <td class="left_title" align="middle">广 而 告 之</td>
          </tr>
          <tbody id="plinfo">
          <tr>
            <td class="left_tdbg1" vAlign="top"><xsl:value-of select="body/MyBlog/BlogIntro" disable-output-escaping="yes"/></td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('plfangke')">
            <td class="left_title" align="middle">最近访客</td>
          </tr>
          <tbody id="plfangke">
          <tr>
            <td class="left_tdbg1" vAlign="top" height="50">
<xsl:for-each select="body/NewVisitor/visitor">
                <li><a href="{username}/"><xsl:value-of select="username"/>(<xsl:value-of select="num"/>)</a></li>
</xsl:for-each></td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr onclick="new Element.toggle('newpl')">
            <td class="left_title" align="middle">最 新 评 论</td>
          </tr>
          <tbody id="newpl">
          <tr>
            <td class="left_tdbg1" vAlign="top"><xsl:apply-templates select="body/NewCommentList/Commentitem"/>
            </td>
          </tr>
          </tbody>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
</td>
      <td width="5"></td>
      <td vAlign="top">
      <!--显示我的照片-->
<div id="diarycontent"><xsl:apply-templates select="body/MyBlog/Diary"/></div>
<div id="showpage">载入分页信息...</div>
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
      <td class="Bottom_Adminlogo" colSpan="2"> | <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">设为首页</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">加入收藏</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>联系站长</a> | <A class="Bottom" href="../FriendSite/Index.asp" target="_blank">友情链接</A> | <A class="Bottom" href="../Copyright.asp" target="_blank">版权申明</A> | 
          <xsl:choose>  
              <xsl:when test="body/Site/ShowAdminLogin = 'enable'">  
              <a class="Bottom" href="../{body/Site/AdminDir}/Admin_Index.asp" target="_blank">管理登录</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
              </xsl:when>
          </xsl:choose>
      </td>
    </tr>
    <tr class="Bottom_Copyright">
      <td width="20%"><IMG src="../images/logo.gif" /></td>
      <td width="80%" align="center"> 站长：<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a><br /><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/></td>
    </tr>
  </table>
<script type="text/javascript">
	Sortable.create('diarycontent',{tag:'div'});
        setTimeout("addfangke(<xsl:value-of select="body/MyBlog/BlogID"/>,'<xsl:value-of select="body/MyBlog/BlogDir"/>')",6000);
</script>
</body>
</xsl:element>
</xsl:template>

<xsl:template match="Diary">
        <div id="diarybody">
        <div id="diarytitle"><table width="100%"><tr valign="middle"><td><font color="#ffffff"><xsl:value-of select="Title" disable-output-escaping="yes"/></font></td><td align="right">浏览<font color="red"><xsl:value-of select="Hits"/></font>次</td></tr></table></div>
        <div id="diarytext"><img src="{Content}" width="500"/></div>
        <div id="diaryfoot">[<b style="cursor:hand;" onclick="new Element.toggle('comment_{Title}')">查看评论</b>(共<font color="red"><xsl:value-of select="Comment"/></font>条)]<xsl:text> </xsl:text>[<b style="cursor:hand;" onclick="showComment({ID},{/body/MyBlog/BlogID});">发表评论</b>]<xsl:text> </xsl:text>[发布时间<xsl:value-of select="Datetime"/>]</div>
        <div id="comment_{Title}" style="display:none">
            <xsl:for-each select="CommentList">  
                <div id="commentbody">
                <div id="commenttitle"><xsl:value-of select="name"/>在<xsl:value-of select="datetime"/>评论说:<b><xsl:value-of select="title"/></b></div>
                <div id="commentcontent"><xsl:value-of select="content"/></div>
                </div>
            </xsl:for-each>
        </div>
        </div>
</xsl:template>

<xsl:template match="Commentitem">
        <li><xsl:value-of select="title"/></li>
</xsl:template>

<xsl:template match="Channelitem">
        <xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{link}" class="channel"><xsl:value-of select="title"/></a>
</xsl:template>

<xsl:template match="Announceitem">
        <a href="../{link}"><xsl:value-of select="title"/></a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:value-of select="DateAndTime"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
</xsl:template>

</xsl:stylesheet>