<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">
<xsl:element name="html">
<head>
<title><xsl:value-of select="body/MyBlog/BlogName"/></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="sealove.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<![CDATA[
/* �϶����CSS���� */
div#diarycontent {
	width: 100%;
	float: left;
	margin: 0px;
}
div#diarybody {
	width: 780px;
	float: left;
	margin: 2px;
	border: 1px solid #d2d3d9;
}

div#diarytitle {
        color: ffffff;background:url(../skin/sealove/Top_01BG.gif);border-top: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;text-align: left;padding-left:30;height: 29;
}

div#diarytext {
        font-family:����;text-align: left;padding-left:5;font-size: 9pt;line-height: 15pt;text-indent: 20px
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
        border: 1px solid #d2d3d9;text-align: left;padding-left:10;height: 20;text-indent: 20px;
}
div#pltext {
        text-align: top;
}
div#showpage {
	width: 780px;
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
        var temppage = "<b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + ",1," + totalitem + "," + iBid + ",1);\">��ҳ</b>";
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
            temppage += " <b style=\"cursor:hand;\" onclick=\"ChangePage(" + totalpage + "," + totalpage + "," + totalitem + "," + iBid + ",1);\">βҳ</b> ��[" + totalpage + "]ҳ ";
        }else{
            temppage += " βҳ</b> ��[" + totalpage + "]ҳ ";
        }
        $('showpage').innerHTML = temppage;
    }else{
        Element.hide('showpage');
    }
}

function ChangePage(iTotalPage,iPage,iTotal,Bid2,iType)
{
    s2 = iPage;
    $('diarycontent').innerHTML = "���ݸ�����...";
    var url = "Showdiary.asp";
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
����var xml = new ActiveXObject("Microsoft.XMLDOM");
����xml.async = false;
����xml.load(originalRequest.responseXml);
    var root = xml.getElementsByTagName("Diary");
    for(i = 0; i < root.length; i++){
        tempstr += "<div id=\"diarybody\">";
        tempstr += "<div id=\"diarytitle\"><table width=\"100%\"><tr valign=\"middle\"><td><font color=\"#ffffff\">" + root.item(i).getElementsByTagName("Title").item(0).text + "</font></td><td align=\"right\">���<font color=\"red\">" + root.item(i).getElementsByTagName("Hits").item(0).text + "</font>��</td></tr></table></div>";
        tempstr += "<div id=\"diarytext\">" + root.item(i).getElementsByTagName("Content").item(0).text + "</div>";
        tempstr += "<div id=\"diaryfoot\">[<b style=\"cursor:hand;\" onclick=\"new Element.toggle('comment_" + root.item(i).getElementsByTagName("Title").item(0).text + "')\">�鿴����</b>(<font color=\"red\">" + root.item(i).getElementsByTagName("Comment").item(0).text + "</font>)][<b style=\"cursor:hand;\" onclick=\"showComment(" + root.item(i).getElementsByTagName("ID").item(0).text
        tempstr += "," + xml.getElementsByTagName("MyBlog/BlogID").item(0).text
        tempstr += ");\">��������</b>] [����ʱ��" + root.item(i).getElementsByTagName("Datetime").item(0).text + "]</div>";
        tempstr += "<div id=\"comment_" + root.item(i).getElementsByTagName("Title").item(0).text + "\" style=\"display:none\">";
            var commentstr = root.item(i).getElementsByTagName("CommentList");
            for(j = 0; j < commentstr.length; j++){
                tempstr += "<div id=\"commentbody\">";
                tempstr += "<div id=\"commenttitle\">" + commentstr.item(j).getElementsByTagName("name").item(0).text + "��" + commentstr.item(j).getElementsByTagName("datetime").item(0).text + "����˵:<b>" + commentstr.item(j).getElementsByTagName("title").item(0).text + "</b></div>";
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
    templ += "<div id=\"diarytitle\">��������</div>";
    templ += "<div id=\"pltext\">";
    templ += "<input name=\"plname\" type=\"hidden\" value=\"" + username + "\">"
    templ += "<input name=\"plpass\" type=\"hidden\" value=\"" + userpass + "\">";
    if(userstat=='login'){
        templ += "�������� <input type=\"checkbox\" name=\"noname\" value='1'><br />";
    }else{
        templ += "�������� <input type=\"checkbox\" name=\"noname\" value='1' Checked Disabled><br />";
    }
    templ += "���۱��� <input name=\"pltitle\" id=\"pltitle\" type=\"text\"><br />";
    templ += "�������� <textarea name=\"plcontent\" cols=\"50\" rows=\"4\"></textarea><br />";
    templ += "<input name=\"plid\" id=\"plid\" type=\"hidden\" value=" + itemID + ">";
    templ += "<input name=\"blogid\" id=\"blogid\" type=\"hidden\" value=" + blogid + ">";
    templ += "<center><a href=\"#\" onclick=\"SaveComment();\">����</a> <a href=\"#\" onclick=\"CancelComment();\">ȡ��</a></center></div>";
    templ += "</div>";
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
    var saveurl = "Showdiary.asp?Action=savepl";
    var name = $F('plname');
    var noname = $F('noname');
    var plpass = $F('plpass');
    var title = $F('pltitle');
    var content = $F('plcontent').stripTags();
    var pid = $F('plid');
    var blogid = $F('blogid');
    //$('pltext').innerHTML = "<center>����������...</center>";
    if((noname!='1')&&(name=='')&&(userstat=='login')){
        alert("����δ��¼,��ѡ����������!");
    }else{
        if((noname!='1')&(plpass=='')&&(userstat=='login')){
            alert("����δ��¼,��ѡ����������!");
        }else{
            if(title==''){
                 alert("���ⲻ��Ϊ��!");
                 Field.focus('pltitle');
            }else{
                if(content==''){
                     alert("���ݲ���Ϊ��!");
                     Field.focus('plcontent');
                }else{
                    // ����������ϢXML�ĵ�
                    var checkurl = "Showdiary.asp?Action=savepl";

                    var pl_dom = new ActiveXObject("Microsoft.XMLDOM");
                    pl_dom.async=false;

                    var p = pl_dom.createProcessingInstruction("xml","version=\"1.0\" encoding=\"gb2312\""); 
                    //����ļ�ͷ 
                    pl_dom.appendChild(p); 

                    //�������ڵ�
                    var objRoot = pl_dom.createElement("root");

                    //�����ӽڵ�
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
                    objField.text = 3;
                    objRoot.appendChild(objField);
                    objField = pl_dom.createNode(1,"id",""); 
                    objField.text = pid;
                    objRoot.appendChild(objField);
                    objField = pl_dom.createNode(1,"blogid",""); 
                    objField.text = blogid;
                    objRoot.appendChild(objField);

                    //��Ӹ��ڵ�
                    pl_dom.appendChild(objRoot);

                    // ��XML�ĵ����͵�Web������
                    var plhttp = getHTTPObject();
                    plhttp.open("POST",checkurl,false);
                    plhttp.send(pl_dom);
                    // ��ʾ���������ص���Ϣ
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
����var xml = new ActiveXObject("Microsoft.XMLDOM");
����xml.async = false;
����xml.load(backRequest.responseXml);
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
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td colspan="3" height="6"><img src="../Skin/sealove/space.gif" /></td></tr>
	<tr><td width="14" height="34"><img src="../Skin/sealove/Top_01Left.gif" /></td>
	    <td background="../Skin/sealove/Top_01BG.gif" align="right" style="color:#FFFFFF">
            <xsl:choose>  
                <xsl:when test="body/Site/ShowSiteChannel = 'enable'">  
                |<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{body/Site/SiteUrl}" class="channel">��վ��ҳ</a><xsl:apply-templates select="body/ChannelList/Channelitem"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
                </xsl:when>
            </xsl:choose>
            </td>
		<td width="14"><img src="../Skin/sealove/Top_01Right.gif" /></td>
	</tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="4" background="../Skin/sealove/Top_02Left.gif"><img src="../Skin/sealove/space.gif" /></td>
		<td background="../Skin/sealove/Top_02BG.gif">
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
		<td width="4" background="../Skin/sealove/Top_02Right.gif"><img src="../Skin/sealove/space.gif" /></td>
	</tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="17" background="../Skin/sealove/Top_03Left.gif"><img src="../Skin/sealove/space.gif" /></td>
		<td background="../Skin/sealove/Top_03BG.gif">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="75" height="33">
            <xsl:choose>  
                <xsl:when test="body/Site/EnableRss = 'enable'">  
                <a href="../Rss.asp" Title="Rss 2.0" Target="_blank"><img src="../images/rss.gif" border="0" /></a>
                </xsl:when>
            </xsl:choose>
                            </td>
				<td width="20"><img src="../Skin/sealove/icon01.gif" /></td>
				<td width="60">���¹��棺</td>
				<td width="400"><MARQUEE onmouseover="this.stop()" onmouseout="this.start()" scrollAmount="1" scrollDelay="4" width="400" align="left"><p><xsl:apply-templates select="body/AnnounceList/Announceitem"/></p></MARQUEE></td>
				<td align="right"></td>				
			</tr>
		</table>
		</td>
		<td width="17" background="../Skin/sealove/Top_03Right.gif"><img src="../Skin/sealove/space.gif" /></td>
	</tr>
	<tr><td colspan="3" height="5"><img src="../Skin/sealove/space.gif" /></td></tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="15"><img src="../Skin/sealove/Main_TopLeft.gif" /></td>
		<td height="11" background="../Skin/sealove/Main_TopBG.gif"><img src="../Skin/sealove/space.gif" /></td>
		<td width="15"><img src="../Skin/sealove/Main_TopRight.gif" /></td>
	</tr>
</table>

<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td width="8" background="../Skin/sealove/Main_Left.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="135" height="60"><img src="../Skin/sealove/Main_Search.gif" alt="վ������" /></td>
				<td><table cellSpacing="0" cellPadding="0" border="0">
					<FORM name="search" action="../search.asp" method="post">
					<tr><td align="middle"><Input id="Keyword" maxLength="50" value="�ؼ���" name="Keyword" /></td>
						<td align="center" width="55"><input name="Submit" id="Submit" type="image" src="../Skin/sealove/Icon_Search.gif" style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" /></td>
						<td align="middle">
				<input type="radio" value="Article" name="ModuleName" checked="True" /> ����
				<Input type="radio" value="Soft" name="ModuleName" /> ����
				<Input type="radio" value="Photo" name="ModuleName" /> ͼƬ
				<Input id="Field" type="hidden" value="Title" name="Field" /></td>
					</tr>
					</FORM>
					</table>
				</td>
				<td width="166" align="right"><img src="../Skin/sealove/Main_girl01.gif" /></td>
			</tr>
		</table>


		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td valign="top">
				<table width="98%" align="center" border="0" cellpadding="0" cellspacing="0" background="../Skin/sealove/Path_BG.gif">
					<tr><td width="9"><img src="../Skin/sealove/Path_Left.gif" /></td>
						<td width="20"><img src="../Skin/sealove/icon02.gif" /></td>
						<td>�����ڵ�λ�ã�<a class="LinkPath"><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> <xsl:value-of select="body/MyBlog/BlogName"/></td>
						<td width="84"><a href="/Reg/User_Reg.asp" target="_blank"><img src="../Skin/sealove/Button_Reg.gif" alt="��Աע��" border="0" /></a></td>
						<td width="9"><img src="../Skin/sealove/Path_Right.gif" /></td>
					</tr>
				</table>
				</td>
				<td width="92" height="48" align="right" valign="top"><img src="../Skin/sealove/Main_girl02.gif" /></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td height="2" bgcolor="#0099BB"></td></tr>
			<tr><td height="1"></td></tr>
			<tr><td height="1" bgcolor="#0099BB"></td></tr>
			<tr><td height="2"></td></tr>
			<tr><td background="../Skin/sealove/AD02.gif" align="right">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                           <td><img src="../Skin/sealove/AD01.gif" border="0" alt="ͨ�����λ���������޸�ΪJS���ô���" /></td>
                           <td align="right"><img src="../Skin/sealove/AD08.gif" border="0" /></td>
                        </tr>
                        </table>			  
			</td>
			</tr>
			<tr><td height="8"></td></tr>
		</table>
	</td>
	<td width="8" background="../Skin/sealove/Main_Right.gif"></td>
	</tr>
</table>
  <!-- ********��ҳ�����������******** -->
  <!-- ********��ҳ�в����뿪ʼ******** -->
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td width="8" background="../Skin/sealove/Main_Left.gif"></td>
      <td vAlign="top">
      <!--��ʾ�ҵ���־-->
<div id="diarycontent"><xsl:apply-templates select="body/MyBlog/Diary"/></div>
<div id="showpage">�����ҳ��Ϣ...</div>
      </td>
      <td width="5"></td>
	<td width="180" valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr onClick="new Element.toggle('login')"><td><img src="../Skin/sealove/Login_Top.gif" alt="��Ա��¼" /></td></tr>
			<tr><td background="../Skin/sealove/Login_BG1.gif">
				<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" style="background-image: url(../Skin/sealove/Login_BG.gif); background-repeat: no-repeat; background-position: center top">
					<tr><td height="6"></td></tr>
					<tr><td><div id="UserLogin">������...<script language="JavaScript" type="text/JavaScript">LoadUserLogin("../",0,0);</script></div></td></tr>
					<tr><td height="6"></td></tr>
				</table>
			</td></tr>
			<tr><td><img src="../Skin/sealove/Room_bottom.gif" /></td></tr>
		</table><br />


		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../Skin/sealove/Column01_L.gif" /></td>
				<td background="../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>�� �� �� ��</b></td>
				<td width="10"><img src="../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center"><a href="Showdiary.asp?BlogID={body/MyBlog/BlogID}">= ȫ����־ =</a><br /><a href="{body/MyBlog/BlogDir}/">= ������ҳ =</a><br /><a href="index.asp">= �ռ��б� =</a></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../Skin/sealove/space.gif" /></td></tr>
		</table><br />
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../Skin/sealove/Column01_L.gif" /></td>
				<td background="../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>�� ־ �� ��</b></td>
				<td width="10"><img src="../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center"><xsl:value-of select="body/MyBlog/BlogIntro" disable-output-escaping="yes"/></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../Skin/sealove/space.gif" /></td></tr>
		</table><br />
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../Skin/sealove/Column01_L.gif" /></td>
				<td background="../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>����ÿ�</b></td>
				<td width="10"><img src="../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center">
				<xsl:for-each select="body/NewVisitor/visitor">
               				 <li><a href="{username}/"><xsl:value-of select="username"/>(<xsl:value-of select="num"/>)</a></li>
				</xsl:for-each></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../Skin/sealove/space.gif" /></td></tr>
		</table><br />
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="33" height="28"><img src="../Skin/sealove/Column01_L.gif" /></td>
				<td background="../Skin/sealove/Column01_BG.gif" style="color:#FFFFFF"><b>�� �� �� ��</b></td>
				<td width="10"><img src="../Skin/sealove/Column01_R.gif" /></td></tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="1" bgcolor="#AAAAAA"></td>
				<td valign="top" height="50">
				<table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr><td align="center"><xsl:apply-templates select="body/NewCommentList/Commentitem"/></td></tr>
				</table>
				</td>
				<td width="1" bgcolor="#AAAAAA"></td>
			</tr>
		</table>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="6" height="23"><img src="../Skin/sealove/Column01_Lb.gif" /></td>
				<td background="../Skin/sealove/Column01_BGb.gif" align="right"></td>
				<td width="6"><img src="../Skin/sealove/Column01_Rb.gif" /></td>
			</tr>
			<tr><td colspan="3" height="8"><img src="../Skin/sealove/space.gif" /></td></tr>
		</table><br />

</td>
<td width="5"></td>
<td width="8" background="/Skin/sealove/Main_Right.gif"></td>
</tr>
</table>

  <!-- ********��ҳ�в��������******** -->
  <!-- ********��ҳ�ײ����뿪ʼ******** -->
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr><td width="8" background="../Skin/sealove/Main_Left.gif"></td>
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
			<tr><td height="25" align="center" style="color:#000000"> | <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">��Ϊ��ҳ</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">�����ղ�</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>��ϵվ��</a> | <A class="Bottom" href="../FriendSite/Index.asp" target="_blank">��������</A> | <A class="Bottom" href="../Copyright.asp" target="_blank">��Ȩ����</A> | 
          <xsl:choose>  
              <xsl:when test="body/Site/ShowAdminLogin = 'enable'">  
              <a class="Bottom" href="../{body/Site/AdminDir}/Admin_Index.asp" target="_blank">�����¼</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
              </xsl:when>
          </xsl:choose>
          </td></tr>
	  <tr><td height="1" background="../Skin/sealove/line01.gif"></td></tr>
	  </table>
	  <table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr><td width="210" align="center"><a href="http://www.asp163.net" target="_blank"><img src="../Skin/sealove/PElogo_sealove.gif" border="0" alt="��������" /></a></td>
				<td> վ����<A href="mailto:info@powereasy.net"></A><br />
				  ģ����ƣ�<a href="http://www.mz25.net/" target="_blank">÷��</a></td>
				<td width="20"></td>
				<td width="120" height="80">
					<a href="http://www.miibeian.gov.cn" target="_blank">
					<img src="../Skin/sealove/mii.gif" border="0" alt="��Ϣ��ҵ������" /><br />*ICP��********��</a></td>
			</tr>
	  </table>
	  </td><td width="8" background="../Skin/sealove/Main_Right.gif"></td></tr>
</table>
<table width="1000" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="15"><img src="../Skin/sealove/Main_BottomLeft.gif" /></td>
		<td height="11" background="../Skin/sealove/Main_BottomBG.gif"><img src="../Skin/sealove/space.gif" /></td>
		<td width="15"><img src="../Skin/sealove/Main_BottomRight.gif" /></td>
	</tr>
	<tr><td colspan="3" height="5"></td></tr>
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
        <div id="diarytitle"><table width="100%"><tr valign="middle"><td><font color="#ffffff"><xsl:value-of select="Title" disable-output-escaping="yes"/></font></td><td align="right">���<font color="red"><xsl:value-of select="Hits"/></font>��</td></tr></table></div>
        <div id="diarytext"><xsl:value-of select="Content" disable-output-escaping="yes"/></div>
        <div id="diaryfoot">[<b style="cursor:hand;" onclick="new Element.toggle('comment_{Title}')">�鿴����</b>(��<font color="red"><xsl:value-of select="Comment"/></font>��)]<xsl:text> </xsl:text>[<b style="cursor:hand;" onclick="showComment({ID},{/body/MyBlog/BlogID});">��������</b>]<xsl:text> </xsl:text>[����ʱ��<xsl:value-of select="Datetime"/>]</div>
        <div id="comment_{Title}" style="display:none">
            <xsl:for-each select="CommentList">  
                <div id="commentbody">
                <div id="commenttitle"><xsl:value-of select="name"/>��<xsl:value-of select="datetime"/>����˵:<b><xsl:value-of select="title"/></b></div>
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