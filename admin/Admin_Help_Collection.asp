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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�ɼ�����</font></td>
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
    <td width="20">&nbsp;</td>
    <td>������ӭ������<%=SiteName%>�ɼ�����ģ�飡��ϵͳ�ǻ����Ƚ���Internet�ɼ����������������ɼ���վ��Ϣ��������Ϣ�ɼ����ȸ��Ի����á�������ͨ����ģ�鼰ʱ�ɼ�����������Ϣ���ݣ�������Ϣ�洢��������վ���ݿ��С��������Խ��ɼ�����Ϣ����˵ķ�ʽɸѡ������</td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_Collection.asp?Action=Main" target=main>���²ɼ�</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_CollectionHistory.asp?Action=main" target=main>�ɼ���ʷ��¼</A></td>
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
    <td width="400" valign="top">������������������вɼ����� ���鿴�ɼ�����Ŀ���ơ��ɼ���ַ������Ƶ����������Ŀ������ר�⡢״̬�Լ��ϴβɼ�����ȡ�<br>
    ������ݲ˵���<A href="Admin_Collection.asp?Action=Main" target=main><font color="#FF0000"><u>���²ɼ�</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">������������ʱ��ѯ�����вɼ�����ʷ��<br>
������ݲ˵���<A href="Admin_CollectionHistory.asp?Action=main" target=main><font color="#FF0000"><u>�ɼ���ʷ��¼</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_CollectionManage.asp?Action=ItemManage" target=main>��Ŀ����</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">��Ŀ����</td>
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
    <td width="400" valign="top">������������Ӳɼ���Ŀ�������²��������Ŀ���ã������Ŀ &gt;&gt; �������� &gt;&gt; �б����� &gt;&gt; �������� &gt;&gt; �������� &gt;&gt; �������� &gt;&gt; �������� &gt;&gt; ��ɡ�<br>
    ������ݲ˵���<A href="Admin_CollectionManage.asp?Action=ItemManage" target=main><font color="#FF0000"><u>��Ŀ����</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�������ñ����ܣ������Խ����Ĳɼ����ݿ�����ϵͳ�ṩ�Ĺ��ܽ��е��롢������<br>
    ������ݲ˵���<A href="Admin_CollectionManage.asp?Action=Import" target=main><font color="#FF0000"><u>��Ŀ����</u></font></A> | <A
href="Admin_CollectionManage.asp?Action=Export" target=main><font color="#FF0000"><u>��Ŀ����</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="Admin_Filter.asp?Action=main" target=main>���˹���</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">&nbsp;</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">�������ɶԲɼ������ݽ��й��˹�������������Զ��������Ŀ�����趨�������ơ����˶��󡢹������͡������������滻���ݣ������� ��ʱ���û�رչ�����Ŀ��<br>
    ������ݲ˵���<A href="Admin_Filter.asp?Action=main" target=main><font color="#FF0000"><u>���˹���</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;</td>
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
