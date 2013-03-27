<?xml version="1.0" encoding="GB2312"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:variable name="title" select="/body/Site/SiteTitle"/>
<xsl:template match="/">

<xsl:element name="html">
<head>
<title><xsl:value-of select="body/Site/SiteTitle"/> >> ��Ʒ���б�ҳ</title>
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
            <td align="left"><a href="Rss.asp" Title="Rss 2.0" Target="_blank"><img src="images/rss.gif" border="0" /></a><img src="images/Wap.gif" border="0" alt="WAP���֧��" style="cursor:hand;"  onClick="window.open('/Wap.asp?ReadMe=Yes', 'Wap', 'width=160,height=257,resizable=0,scrollbars=no');" /></td>
            <td align="right">|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><a href="{body/Site/SiteUrl}" class="channel">��վ��ҳ</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:apply-templates select="body/ChannelList/Channelitem"/><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text></td>
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
      <!--���������ڴ��뿪ʼ-->
        <table class="top_nav_menu" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
          <tr>
            <td align="middle" width="50"><IMG src="/Images/arrow3.gif" align="absMiddle" /></td>
            <td width="40%">�����ڵ�λ�ã� <a><xsl:attribute name="href"><xsl:value-of select="body/Site/SiteUrl"/></xsl:attribute><xsl:value-of select="body/Site/SiteName"/></a> >> ��Ʒ���б�ҳ</td>
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
  <!-- ********��ҳ�����������******** -->
  <!-- ********��ҳ�в����뿪ʼ******** -->
  <table class="center_tdbgall" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr>
      <td vAlign="top" width="180">
      <!--�û���¼���뿪ʼ-->
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
      <!--�û���¼�������--></td>
      <td width="5"></td>
      <td vAlign="top" width="354">
      <!--��վ����Blog���뿪ʼ-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_575"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>���������Ʒ��</B></td>
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
      <!--��վ����Blog���뿪ʼ--></td>
      <td width="5"></td>
      <td vAlign="top" width="216">
      <!--�ر��Ƽ����뿪ʼ-->
        <table style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_575"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>�¼�����Ʒ��</B></td>
          </tr>
          <tr>
            <td class="main_tdbg_575" vAlign="top" height="194">
               <!-- ********ѭ������������BLOG�б�******** -->
               <table width="100%"><xsl:apply-templates select="body/AddBlog/Blogitem"/></table>
            </td>
          </tr>
          <tr>
            <td class="main_shadow"></td>
          </tr>
        </table>
      <!--�ر��Ƽ��������-->
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
      <!--��վ�Ƽ���Ʒ�����뿪ʼ-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_760"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>�Ƽ���Ʒ��</B></td>
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
      <!--��Ʒ�����࿪ʼ-->
        <table cellSpacing="0" cellPadding="0" width="100%" border="0">
          <tr>
            <td class="main_title_760"><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><B>��Ʒ�������б�</B></td>
          </tr>
          <tr>
            <td><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><a href="showblog.asp">ȫ������</a><xsl:text disable-output-escaping="yes">&amp;nbsp;&amp;nbsp;&amp;nbsp;</xsl:text><xsl:apply-templates select="body/BlogClassList/item"/></td>
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
  <!-- ********��ҳ�в��������******** -->
  <!-- ********��ҳ�ײ����뿪ʼ******** -->
  <table class="Bottom_tdbgall" style="WORD-BREAK: break-all" cellSpacing="0" cellPadding="0" width="760" align="center" border="0">
    <tr align="middle">
      <td class="Bottom_Adminlogo" colSpan="2"> | <A class="Bottom" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('{body/Site/SiteUrl}');" style="cursor:hand;">��Ϊ��ҳ</A> | <A class="Bottom" href="javascript:window.external.addFavorite('{body/Site/SiteUrl}','{body/Site/SiteName}');">�����ղ�</A> | <a class="Bottom"><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute>��ϵվ��</a> | <A class="Bottom" href="FriendSite/Index.asp" target="_blank">��������</A> | <A class="Bottom" href="Copyright.asp" target="_blank">��Ȩ����</A> |  <a class="Bottom" href="Admin/Admin_Index.asp" target="_blank">�����¼</a><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>|<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text></td>
    </tr>
    <tr class="Bottom_Copyright">
      <td width="20%"><IMG src="images/logo.gif" /></td>
      <td width="80%" align="center"> վ����<a><xsl:attribute name="href">mailto:<xsl:value-of select="body/Site/WebmasterEmail"/></xsl:attribute><xsl:value-of select="body/Site/WebmasterName"/></a><br /><xsl:value-of select="body/Site/Copyright" disable-output-escaping="yes"/></td>
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
		<xsl:attribute name="alt">����:<xsl:value-of select="title"/><br/>����:<xsl:value-of select="BirthDay"/></xsl:attribute>
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