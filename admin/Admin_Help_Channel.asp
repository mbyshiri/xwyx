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
        <td><span class="STYLE4">Ƶ����������</span></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000"><%= ChannelName %>����</font></td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Add" target="main">�������</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_Template.asp?ChannelID=<%=ChannelID%>" target="main">ģ�����</a></td>
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
    <td width="400">����������ʹ��ϵͳ�ṩ��ǿ��<a href="#" title="���߱༭���ܹ�����ҳ��ʵ���������༭������磺Word�������е�ǿ����ӱ༭���ݵĹ��ܡ�"><u>���߱༭��</u></a>�����վ���ݣ�����ѡ����ģʽ�͸߼�ģʽ���߼�ģʽ�ܽ��и���߼������ã��磺<FONT color=#ff0000><a href="#" title="����������ϵͳ����ĵ�ַ�С����˱���׼�����ӵ�������վ�е�����ʱ����ʹ�����ַ�ʽ��"><u>ת������</u></a></FONT>��<u>������</u>��<a href="#" title="�������ݵ��Ķ��ȼ���ֻ�о�����ӦȨ�޵��˲����Ķ�������"><u>�Ķ��ȼ�</u></a>��<a href="#" title="�����û����Ķ�������ʱ��������Ӧ�����������οͺ͹���Ա��Ч��"><u>�Ķ�����</u></a>��<a href="#" title="�����ù̶����ȵ㡢�Ƽ������ݵ�����"><u>��������</u></a>��<a href="#" title="������ɫ������ģ���а���CSS����ɫ��ͼƬ����Ϣ���Ͱ������ģ�壨���ģ���а����˰�����Ƶİ�ʽ����Ϣ��"><u>ѡ��ģ��</u></a>�ȡ�<br>
      ������ݲ˵���<a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Add&AddType=1" target=main><font color="#FF0000"><u>�������</u></font></a>
      <%If ModuleType = 1 Then%>
      | <a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Add&AddType=3" target=main><font color="#FF0000"><u>���ǩ������</u></font></a>
      <%End If%>
    </td>
    <td width="40">&nbsp;</td>
    <td width="400">����ϵͳ�ṩ��ʽģ������ܣ�������ʾǰ̨ʱ����������ҳ�Ľ��沼����ʽ�����������񲼾֡�ͼƬ������Ҫ��ʾ��λ�õȵȡ���������ǰ̨��ʾ��ҳ���ʽ���ڴ��޸������á�<br>
    ������ݲ˵���<a href="Admin_Template.asp?ChannelID=<%=ChannelID%>" target=main><u><font color="#FF0000">����ģ��</font></u></a> | <a href="Admin_Template.asp?ChannelID=<%=ChannelID%>&Action=Add&TemplateType=1"><u><font color="#FF0000">���ģ��</font></u></a> | <a href="Admin_Template.asp?ChannelID=<%=ChannelID%>&Action=Import"><u><font color="#FF0000">����ģ��</font></u></a> | <a href="Admin_Template.asp?ChannelID=<%=ChannelID%>&Action=Export"><u><font color="#FF0000">����ģ��</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&Passed=All" target="main">��������</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">��������</td>
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
    <td width="400">��������ӵ������ṩ��ݵĹ�����ӦȨ�޹���Ա�������ע���û���������ݣ��޸ġ��ƶ���ɾ���Ѿ���������ݣ��������޸����ݵ����ԣ�Ҳ�ɽ�ָ�������������ƶ�����һ��Ŀ�С�<br>
      ������ݲ˵���<a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&Passed=All" target=main><font color="#FF0000"><u>��������</u></font></a> | <a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&ManageType=Check&Passed=False" target=main><font color="#FF0000"><u>�������</u></font></a> | <a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&Passed=True&ManageType=HTML" target=main><font color="#FF0000"><u>����HTML</u></font></a> | <a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&ManageType=MyArticle&Passed=All" target=main><font color="#FF0000"><u>�Ҽӵ�����</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400">���������һЩ��Ŀ������Ҫ�޸���ͬ�����ã��������ϵͳ�ṩ���������ù��ܽ��й����������ƶ����ݡ������޸����ݡ���Ŀר�����ݵȡ�<br>
    ������ݲ˵���<a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&Passed=All"><font color="#FF0000"><u>�����޸�����</u></font></a> | <a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=BatchMove" target=main><font color="#FF0000"><u>�����ƶ�����</u></font></a> | <a href="Admin_Class.asp?ChannelID=<%=ChannelID%>&Action=Batch"><font color="#FF0000"><u>����������Ŀ</u></font></a></td>
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
    <td width="100" align="center" class="topbg"><a  class='Class' href="Admin_Class.asp?ChannelID=<%=ChannelID%>" target="main">������Ŀ</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_<%=ModuleName%>JS.asp?ChannelID=<%=ChannelID%>" target=main>����JS�ļ�</a></td>
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
    <td width="400">��������Ƶ�������õĸ�����Ŀ����Ŀ�����޼����๦�ܡ�����ӡ�ɾ���������ƶ�����λ���ϲ�������������Ŀ�ȡ�<br>
      ������ݲ˵���<a href="Admin_Class.asp?ChannelID=<%=ChannelID%>" target="main"><u><font color="#FF0000">��Ŀ����</font></u></a> | <a href="Admin_Class.asp?ChannelID=<%=ChannelID%>&Action=Add" target="main"><u><font color="#FF0000">�����Ŀ</font></u></a> | <a href="Admin_Class.asp?ChannelID=<%=ChannelID%>&Action=Batch" target="main"><u><font color="#FF0000">��������</font></u></a> | <a href="Admin_Class.asp?ChannelID=<%=ChannelID%>&Action=Patch" target="main"><u><font color="#FF0000">�޸���Ŀ�ṹ</font></u></a></td>
    <td width="40">&nbsp;</td>
    <td width="400">����JS������Ϊ�˼ӿ�����ٶ��ر����ɵġ���������ز�����ɾ����Ԥ��Ч�������Զ����ֶ�ˢ���й�JS�ļ���<br>
      ������ݲ˵���<a href="Admin_<%=ModuleName%>JS.asp?ChannelID=<%=ChannelID%>" target=main><font color="#FF0000"><u>����JS�ļ�</u></font></a></td>
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
    <td width="100" align="center" class="topbg">ר��������</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">�ϴ������վ</td>
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
    <td width="400">���������ͬ��Ŀ�е���������ͬһ���⣬����Խ�����Ӧר�⣬�Ա���������ר����Ŀ���Խ����޸ġ�ɾ������յȲ�����<br>
      �����û����Զ���վ�����ݷ���������ۣ�����Ա���Իظ����޸ġ�ɾ����������ۡ�<br>
      ������ݲ˵���<a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&ManageType=Special&Passed=All" target="main"><u><font color="#FF0000">ר�����ݹ���</font></u></a> | <a href="Admin_Special.asp?ChannelID=<%=ChannelID%>" target=main><font color="#FF0000"><u>ר������</u></font></a> | <a href="Admin_Comment.asp?ChannelID=<%=ChannelID%>" target=main><font color="#FF0000"><u>���۹���</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400"> ����ϵͳ����<a href="#" title="ָ����Ҫ���֧�־Ϳ����ϴ�ָ���ļ����͵Ĺ��ܡ�"><u>������ϴ�</u></a>���ܣ����ϴ�Ƶ�����趨���ļ����͡������õ��ϴ��ļ�����ʹ����������ļ����ܶ��ڽ�������<br>
      ����ע���û��͹���Աɾ�����õ�����ʱ����ɾ��������վ���Է�ֹ�����������վ�ڵ����ݿ���ʱ�ָ��������<br>
��ݲ˵�:
    <%
    Dim strUpload
    Select Case ModuleType
    Case 1, 5
        strUpload = "UploadFiles"
    Case 2
        strUpload = "UploadSoft"
    Case 3
        strUpload = "UploadPhotos"
    End Select
    %>
    <a href="Admin_UploadFile.asp?ChannelID=<%=ChannelID%>&UploadDir=<%=strUpload%>" target=main><font color="#FF0000"><u>�ϴ��ļ�����</u></font></a> | <a href="Admin_UploadFile.asp?ChannelID=<%=ChannelID%>&Action=Clear&UploadDir=<%=strUpload%>" target=main><font color="#FF0000"><u>���������ϴ��ļ�</u></font></a> | <a href="Admin_<%=ModuleName%>.asp?ChannelID=<%=ChannelID%>&Action=Manage&ManageType=Recyclebin&Passed=All" target=main><font color="#FF0000"><u>����վ����</u></font></a></td>
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
