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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">���ݿ����</font></td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Backup" target=main>�������ݿ�</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=SpaceSize" target=main>ϵͳ�ռ�ռ��</A></td>
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
    <td width="400">����ϵͳ���������ݿ⣬�Ա����ݿ��������ʱ�ܽ��лָ�������������Ҫ�������ݿ����·��Ŀ¼����Ŀ¼�����ڣ����Զ��������������뱸���ļ�������չ����Ĭ��Ϊ��.asa����������ͬ���ļ���ϵͳ���Զ����ǡ�<br>
      ������ݲ˵���<A href="Admin_Database.asp?Action=Backup" target=main><font color="#FF0000"><U>�������ݿ�</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�����鿴��վϵͳռ�ÿռ������������Բ鿴����ϵͳ����̨����ϵͳͼƬ����Ƶ���������ļ�ռ�ÿռ�������<br>
      ������ݲ˵���<A href="Admin_Database.asp?Action=SpaceSize" target=main><font color="#FF0000"><u>ϵͳ�ռ�ռ��</u></font></A>��</td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Restore" target=main>�ָ����ݿ�</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Init" target=main>ϵͳ��ʼ��</A></td>
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
    <td width="400" valign="top">����������<a href="Admin_CreateSiteIndex.asp" target="main">��</a><a href="Admin_CreateSiteIndex.asp" target="main">��ϵͳ���ݵ����ݿ�</a>�лָ����ݿ⡣��ע�ⱸ�����ݿ�·�����·������������ȷ�����ݿ�����<br>
      ������ݲ˵���<A href="Admin_Database.asp?Action=Restore" target=main><font color="#FF0000"><U>�ָ����ݿ�</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">��������վϵͳ���г�ʼ������ָ�����ݿ����ݵ����ݽ��ᱻ��ա�<FONT color=#0000FF>�����ô˹��ܣ���Ϊһ��������޷��ָ���</FONT>��<br>
      ������ݲ˵���<A href="Admin_Database.asp?Action=Init" target=main><font color="#FF0000"><U>ϵͳ��ʼ��</u></font></A>��</td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Database.asp?Action=Compact" target=main>ѹ�����ݿ�</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center">&nbsp;</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="1" colspan="2" class="topbg2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400">�����ھ�����ɾ�����ݵȲ����󣬿���ʹ�ñ����ܱ�֤���ݿ��������š�ѹ��ǰ�������ȱ������ݿ⣬���ⷢ���������<br>
    ������ݲ˵���<A href="Admin_Database.asp?Action=Compact" target=main><font color="#FF0000"><U>ѹ�����ݿ�</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400">&nbsp;</td>
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
