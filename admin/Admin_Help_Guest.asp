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
<style type="text/css">
<!--
.STYLE4 {color: #000000}
-->
</style>
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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">���԰����</font></td>
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
    <td width="100" align="center" class="topbg">���԰����</td>
    <td width="200">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_GuestBook.asp?Passed=False" target=main>�������</A></td>
    <td width="400">&nbsp;</td>
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
    <td width="300">�������԰�����վ������һ����ʽ��֧��<a href="#" title="UBB ������ Infopop ��˾Ϊ�� Ultimate Bulletin Board ��̳������ר�� HTML ���롣"><u>UBB����</u></a>���߱�ͷ������鹦�ܡ��û���������վ�����ԣ�Ҳ�ɲ鿴�������ͻظ�������Ϣ������Ա����ˡ��ظ����޸ġ�ɾ��������Ϣ������Ա�����ں�̨�����Ƿ���������˹��ܡ�<br>
      ��������ģʽ�ж��֣�<a href="#" title="ָ����Ҫע���Ϊ����վ��ע���û��Ϳ��Բ鿴�������ͻظ���Ϣ"><u>�ο�ģʽ</u></a>��<a href="#" title="ָע���Ϊ����վ��ע���û�����в鿴�������ͻظ���Ϣ���в鿴�Լ������Թ��ܡ�"><u>�û�ģʽ</u></a>��</td>
    <td width="40">&nbsp;</td>
    <td width="500" valign="top">����Ϊ��ֹ�û�������ʱ�����������ۣ��ɿ���ϵͳ��������˹��ܡ���������˹��ܺ��û���������ͨ������Ա��˺������ǰ̨��ʾ������������˹�������<a href="Admin_Channel.asp?Action=Modify&iChannelID=4"><u>����Ƶ������</u></a>�����á�<br>
      ��������Ա���Զ��û����Խ����޸ġ�ɾ�����ظ���ͨ����˵Ĳ�����<br>
      ������ݲ˵���<A href="Admin_GuestBook.asp?Passed=False" target=main><font color="#FF0000"><U>��վ�������</U></font></A>��</td>
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
    <td width="200">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_GuestBook.asp?Passed=All" target=main>���Թ���</A></td>
    <td width="400">&nbsp;</td>
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
    <td width="300" valign="top">&nbsp;</td>
    <td width="40">&nbsp;</td>
    <td width="500" height="100" valign="top">��������Ա���Զ��û����Խ����޸ġ�ɾ�����ظ���ȡ����˵Ĳ�����ȡ����˺�����Բ�����ǰ̨��ʾ��<br>
    ������ݲ˵���<A href="Admin_GuestBook.asp?Passed=All" target=main><font color="#FF0000"><U>��վ���Թ���</U></font></A>��</td>
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
