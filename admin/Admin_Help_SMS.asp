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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�ֻ����Ź���</font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="3">
  <tr>
    <td width="20">&nbsp;</td>
    <td>������ӭ��ʹ���ֻ����Ź��ܣ��������롰���׶���ͨ�����ܼ��ɣ�Ϊ���ṩ��һ����Ч�桢�ͳɱ����ƶ���������ƽ̨�������׶���ͨ���Ƕ��׹�˾���й����ź�����һ����ҵ�񣬶���ͨ�û�����WEB��ʽͨ�������׶���ͨ������ƽ̨���й��ƶ����й���ͨ���й����ź��й���ͨ�û�ʵʱ��ʱ���Ͷ���Ϣ����ҵ��ɹ㷺Ӧ������Ʒ������������֪ͨ��������Ϣ��������ѯ���񡢻��飨������֪ͨ������ף�����²�Ʒ�������ͻ���ͨ�ȷ��棬ʵ���ƶ��칫���ƶ�����
<br />��������ʹ�����ȵ������׶���ͨ������ƽ̨ע���Ա����ֵ��Ȼ���ڱ�ϵͳ�н��С��ֻ��������á���</td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="http://sms.powereasy.net/" target="_blank">���׶���ͨ</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_SiteConfig.asp" target=main>�ֻ���������</A></td>
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
    <td width="400" valign="top">������ʹ���ֻ����Ź���֮ǰ�������뵽�����׶���ͨ������ƽ̨ע�ᣬ�Եõ��û��������롢���ʶ��ź�MD5˽Կ����Ϣ�������г�ֵ��<br>
      ������ݲ˵���<A href="http://sms.powereasy.net/Service.aspx" target="_blank"><font color="#FF0000"><u>���񵼺�</u></font></A> | <A href="http://sms.powereasy.net/Register.aspx" target="_blank"><font color="#FF0000"><u>ע���»�Ա</u></font></A> | <A href="http://sms.powereasy.net/Member/Recharge.aspx" target="_blank"><font color="#FF0000"><u>����ͨ��ֵ</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400">�����ڡ���վ��Ϣ���á��ġ���վѡ��п��������Ƿ����á��ֻ����š����ܣ��ڡ��ֻ��������á��У���д��Ӧ�����ڡ����׶���ͨ��ƽ̨�е�ע���û�����MD5˽Կ����Ԥ����ز������ֻ��������ݡ�<br>
      ������ݲ˵���<A href="Admin_SiteConfig.asp" target="main"><font color="#FF0000"><u>��վ��Ϣ����</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_SMS.asp?SendTo=Member" target="main">�����ֻ�����</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_SMS.asp?SendTo=Other" target="main"></A>�鿴���ͽ��</td>
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
    <td width="400" valign="top">���������Ը���վ�е�ע���Ա��ע�����ϵ�ˡ������е��ջ��˺������˵��ֻ���С��ͨ���Ͷ��š�Ҳ�������úõ��ͻ��ύ����ʱϵͳ�Զ������ֻ�����֪ͨ����Ա��<br>
      ������ݲ˵������͸�<A href="Admin_SMS.asp?SendTo=Member" target="main"><font color="#FF0000"><u>��Ա</u></font></A> | <A href="Admin_SMS.asp?SendTo=Contacter" target="main"><font color="#FF0000"><u>��ϵ��</u></font></A> | <A href="Admin_SMS.asp?SendTo=Consignee" target="main"><font color="#FF0000"><u>�����е��ջ���</u></font></A> | <A href="Admin_SMS.asp?SendTo=Other" target="main"><font color="#FF0000"><u>������</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�������á����׶���ͨ�����͵�ÿһ���������۳ɹ��������ϸ�ķ��ͼ�¼�������ŷ��Ͳ��ɹ�����Ʒѡ������Բ鿴ÿ�η����ֻ����ŵķ��ͽ����<br>
      ������ݲ˵���<A href="Admin_SMSLog.asp" target="main"><font color="#FF0000"><u>�鿴���ŷ��ͽ��</u></font></A></td>
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
