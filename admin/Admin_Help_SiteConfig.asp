<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 0      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��    
Call CloseConn
%>
<html>
<head>
<title><%=SiteName & "--��̨������ҳ"%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="392" rowspan="2"><img src="Images/adminmain01.gif" width="392" height="126"></td>
    <td height="114" valign="top" background="Images/adminmain0line2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20"></td>
      </tr>
      <tr>
        <td><%=AdminName%>���ã�������
          <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">ϵͳ���ù���</font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_SiteConfig.asp">��վ��Ϣ����</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Template.asp?ChannelID=0" target="main">��ҳģ�����</a></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">������վ������Ϣ���ã�����վ�����ơ���ַ��LOGO����Ȩ��Ϣ�ȣ���վ����ѡ�����ã�����ʾƵ��������Զ��ͼƬ�ȣ��û�ѡ�����ã����Ƿ��������û�ע�ᡢ�Ƿ���Ҫ��֤�ȣ������ʼ�������ѡ��ȡ�<br>
    ������ݲ˵���<a href="Admin_SiteConfig.asp" target="main"><u><font color="#FF0000">��վ��Ϣ����</font></u></a> | <a href="Admin_Article.asp?ChannelID=1&amp;Action=Manage&amp;Passed=True&amp;ManageType=HTML" target=main><font color="#FF0000"><u></u></font></a><a href="Admin_SiteConfig.asp#SiteOption" target="main"><u><font color="#FF0000">��վѡ������</font></u></a> | <a href="Admin_Article.asp?ChannelID=1&amp;Action=Manage&amp;Passed=True&amp;ManageType=HTML" target=main><font color="#FF0000"><u></u></font></a><a href="Admin_SiteConfig.asp#User" target="main"><u><font color="#FF0000">�û�ѡ��</font></u></a>��</td>
    <td width="40">&nbsp;</td>
    <td width="400">������һ�ΰ�װϵͳ��<font color="#FF0000">������վ��ҳ</font>����ҳ����Ŀҳ������ҳ��ר��ҳ����������������ȫ��HTMLҳ�棨���ۺ͵����ͳ�Ƴ��⣩����Ƶ���������ɹ�������<a href=Admin_Channel.asp target=main title="������վ����ϵͳ�У�Ƶ����ָĳһ����ģ��ļ��ϡ�ĳһƵ�������Ǿ߱�����ϵͳ���ܣ���߱�����ϵͳ��ͼƬϵͳ�Ĺ��ܡ�"><U>��վƵ������</U></a>�����á�<br>
      ������ݲ˵���<a href="Admin_Template.asp?ChannelID=0"><font color="#FF0000"><u>������վ��ҳģ��</u></font></a> | <a href="Admin_CreateSiteIndex.asp"><font color="#FF0000"><u>������վ��ҳ</u></font></a>��</td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Channel.asp" target="main">��վƵ������</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Skin.asp" target="main">��վ������</a></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">����������վ�ĸ���Ƶ���Ĺ���ģ�飬�����¡����ء�ͼƬ�����Ե�Ƶ����Ƶ���ɷ�Ϊ<a href="#" title="ϵͳ�ڲ�Ƶ��ָ������MY�������й���ģ�飨���š����¡�ͼƬ�ȣ�����������µ�Ƶ������Ƶ���߱�����ʹ�ù���ģ����ȫ��ͬ�Ĺ��ܡ�"><U>ϵͳ�ڲ�Ƶ��</U></a>��<a href="#" title="�ⲿƵ��ָ���ӵ�MY����ϵͳ����ĵ�ַ�С�����Ƶ��׼�����ӵ���վ�е�����ϵͳʱ"><U>�ⲿƵ��</U></a>���ࡣϵͳ��һЩ��Ҫ���ܣ�������HTML���ܡ�Ƶ������˹��ܡ��ϴ��ļ����͡�����������ÿ����ʾ����Ŀ�����ײ���Ŀ��������ʾ��ʽ�����ڴ˽������á�<br>
    ������ݲ˵���<A href="Admin_Channel.asp?Action=Add" target=main><font color="#FF0000"><u>�����վƵ��</u></font></A> | <A href="Admin_Channel.asp" target=main><font color="#FF0000"><u>������վƵ��</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400">�������ģ���ǿ���������վ��ǰ̨��ʾʱ�����ĵ����塢���ͼƬ�ȣ�ͨ������css��ҳ��ʽ�����������ƺͿ��Ƶġ�������ҳ�����еĲ����ʽ��(CSS)��ʽ�������ض���HTML��ǩ�԰����ض���ʽ�����ı���ʽ��ϵͳ�����Զ���CSS��ʽ�Ĺ��ܣ�����ʱ�����޸���ʽ��<br>
      ������ݲ˵���<A href="Admin_Skin.asp?Action=Add" target=main><font color="#FF0000"><u>�����վ���</u></font></A> | <A href="Admin_Skin.asp" target=main><font color="#FF0000"><u>������վ���</u></font></A>��</td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Announce.asp" target="main">��վ�������</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Vote.asp" target="main">��վ�������</a></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">�������Է������޸ĺ�ɾ����վ���档��������ΪƵ�����ù��棬Ҳ���Է�����Ƶ����ͬ�Ĺ��档����֧��<a href="#" title="UBB ������ Infopop ��˾Ϊ�� Ultimate Bulletin Board ��̳������ר�� HTML ���롣"><u>UBB����</u></a>����ȫ�����������������������͵���������ʾ���ͣ�ֻ�н�������Ϊ����ʱ�Ż���ǰ̨��ʾ��<br>
      ������ݲ˵���<A href="Admin_Announce.asp?Action=Add" target=main><u><font color="#FF0000">������վ����</font></u></A> | <A href="Admin_Announce.asp" target=main><font color="#FF0000"><u>������վ����</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400">�������Է������޸ĺ�ɾ����վ���顣��������ΪƵ�����õ��飬Ҳ���Է�����Ƶ����ͬ�ĵ��顣���Է�����������飬Ҳ���Է�����������顣�е�ѡ�Ͷ�ѡ���ֵ������ͣ�ֻ�н�������Ϊ���µ����Ż���ǰ̨��ʾ��<br>
      ������ݲ˵���<A href="Admin_Vote.asp?Action=Add" target=main><u><font color="#FF0000">������վ����</font></u></A> | <A href="Admin_Vote.asp" target=main><font color="#FF0000"><u>������վ����</u></font></A>��</td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Advertisement.asp" target="main">��վ������</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_FriendSite.asp" target="main">�������ӹ���</a></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">�������Է������޸ĺ�ɾ����վ��档�������Ϊ��̬JS�ļ����Է�����ºͼӿ���ʾ�ٶȡ��������ù���λ������λ֧��������ͨ�ñ�׼��ͬһ��λ�������ö����棬ͬһ���������ڲ�ͬ��λ�����ж�����ʾ��ʽ�������ϴ����ͼƬ�����ô�С�����֧��flash��ʽ��ֻ�н�����λ��Ϊ�ʱ�Ż���ǰ̨��ʾ��<br>
      ������ݲ˵���<A href="Admin_Advertisement.asp" target=main><u><font color="#FF0000">������վ���</font></u></A> | <A href="Admin_Advertisement.asp?Action=AddZone" target=main><u><font color="#FF0000">��ӹ���λ</font></u></A> | <A href="Admin_Advertisement.asp?Action=AddAD" target=main><u><font color="#FF0000">����¹��</font></u></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400"> ����ϵͳ�߱��������������վ������������ӹ��ܣ���ִ����ӡ��޸ġ�ɾ���Ȳ������������Ƽ������������ӷֳ�<A href="Admin_FriendSite.asp?LinkType=2" title="��ʾ����վ��������Ϊ����������ʽ��"><u>��������</u></A>��<A href="Admin_FriendSite.asp?LinkType=1" title="��ʾ����վlogoͼƬΪ����������ʽ��"><u>LOGO����</u></A>������ʾ��ʽ��<br>
      ������ݲ˵���<A href="Admin_FriendSite.asp?Action=Add" target=main><font color="#FF0000"><u>�����������</u></font></A> | <A href="Admin_FriendSite.asp" target=main><font color="#FF0000"><u>������������</u></font></A>��<br>
    </td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center">&nbsp;</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Counter.asp" target=main>��վͳ�Ʒ���</A></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="1" colspan="2"></td>
    <td colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">&nbsp;</td>
    <td width="40">&nbsp;</td>
    <td width="400">������ʾ��ϸ����վͳ����Ϣ���ɲ鿴��վ�ۺ�ͳ����Ϣ��������ʼ�¼�����ʴ���������ҳ�桢����ϵͳ�ȷ�����Ϣ��<br>
      ������ݲ˵���<A href="Admin_Counter.asp" target=main><font color="#FF0000"><u>��վͳ�Ʒ���</u></font></A> | <A href="Admin_Counter.asp?Action=FVisitor" target=main><font color="#FF0000"><u>������ʼ�¼</u></font></A>��<br>
    </td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 class="topbg"><span class="Glow">Copyright 2003-2006 &copy; <%=SiteName%> All Rights Reserved.</span>
  </tr>
</table>
</body>
</html>
