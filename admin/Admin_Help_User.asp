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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�û�����</font></td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_User.asp" target=main>ע���û�����</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Admin.asp" target=main>����Ա����</A></td>
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
    <td width="400">���������ܿ�����ϸ������������վע���û�����Ϣ��Ȩ�ޡ����Զ��û������޸ġ�������ɾ�������ѵĲ�����Ҳ���Զ��û�����ɾ���������ͽ����Ĳ����������ƶ��û�����Ӧ��<a href="Admin_UserGroup.asp" target="main" title="�û������û��˻��ļ��ϣ�ͨ�������û��飬��������û������������Ȩ����Ȩ�ޡ������Ȩ�������ڡ�Ƶ����������Ƶ���ġ���Ŀ�����С�"><U>�û���</U></a>��<br>
      ������ݲ˵���<A href="Admin_User.asp" target=main><font color="#FF0000"><u>ע���û�����</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">����ϵͳ����ǿ�����վȨ�޹��������ù���Ա��ϸȨ�ޣ�����ɾ����Ա��ָ����ϸ�Ĺ���Ȩ�ޣ�ʹ��վ�Ĺ���ּ�������˹�ͬ����������վ<a href="Admin_Admin.asp?Action=Add" target="main" title="��������Ա��ӵ������Ȩ�ޡ�ĳЩȨ�ޣ������Ա������վ��Ϣ���á���վѡ�����õȹ���Ȩ�ޣ�ֻ�г�������Ա���С�"><U>��������Ա</U></a>��<a href="Admin_Admin.asp?Action=Add" target="main" title="��ͨ����Ա��ͱ��ָ��������վ�����ܣ���Ҫ��ϸָ��ÿһ�����Ȩ�ޡ�"><U>��ͨ����Ա</U></a>��ͬһ�˺ſ������Ƿ��������ͬʱʹ�ô��ʺŵ�¼��<br>
      ������ݲ˵���<A href="Admin_Admin.asp?Action=Add" target=main><font color="#FF0000"><u>����Ա���</u></font></A> | <A href="Admin_Admin.asp" target=main><font color="#FF0000"><u>����</u></font></A>��</td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_UserGroup.asp" target=main>�û������</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Maillist.asp" target=main>�ʼ��б����</A></td>
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
    <td width="400" valign="top">�����û������û��˻��ļ��ϣ�ͨ�������û��飬��������û������������Ȩ����Ȩ�ޡ��û���Ȩ�޵�����ԽС��˵�����е�Ȩ��Խ�󣨵ȼ�Խ�ߣ���Ȩ�����ò��õȼ��ƣ����ߵȼ����û�����е͵ȼ��û�������Ȩ�ޡ������ʹ��Ȩ�������ڡ�Ƶ����������Ƶ���ġ���Ŀ�����С�<br>
      ������ݲ˵���<A href="Admin_UserGroup.asp" target=main><font color="#FF0000"><u>�û������</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�������û����͡����û������Ͱ��û�Email�����ʼ�����Ϣ�����͵�����ע��ʱ������д��������û����ʼ��б��ʹ�ý����Ĵ����ķ�������Դ��������ʹ�á��������ܿɽ��ʼ��б����������ݿ���ı���<br>
      ������ݲ˵���<A href="Admin_Maillist.asp" target=main><font color="#FF0000"><u>�ʼ��б�</u></font></A> | <A href="Admin_Maillist.asp?Action=Export" target=main><font color="#FF0000"><u>�б���</u></font></A><font color="#FF0000">&nbsp;</font>��</td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_User.asp?Action=Update" target=main>�����û�����</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">�������Ϣ</td>
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
    <td width="400">���������������¼����û��ķ��������������������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С��޸���ʼID�ŵ�����ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����<br>
      ������ݲ˵���<A href="Admin_User.asp?Action=Update" target=main><font color="#FF0000"><u>�����û�����</u></font></A>��</td>
    <td width="40">&nbsp;</td>
    <td width="400">����ϵͳ�ṩ�˶���Ϣ���ܣ���Ҳ����׫д����Ϣ���뱾վ�ڵ�ע���û����н�����������<a href="#" title="�ռ���ֻ�����뱾վע���û���ע�������ռ��˿�����Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������5���û���"><u>�ռ���</u></a>��<a href="#" title="���50���ַ�"><u>����</u></a>��<a href="#" title="���1000���ַ�"><u>����</u></a>�������Թ������Ϣ����ʱ�鿴�Լ��ķ����䣬ɾ�����ڵĶ���Ϣ�Խ�ʡ�������Ŀռ䡣<br>
    <%if AdminPurview=1 or CheckPurview_Other(AdminPurview_Others,"Message")=True then%>
������ݲ˵���<A href="Admin_Message.asp" target=main><font color="#FF0000"><u>�������Ϣ</u></font></A>��</td>
    <%end if%>
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
