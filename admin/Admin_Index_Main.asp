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
%>
<html>
<head>
<title><%=SiteName & "--��̨������ҳ"%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
<style type="text/css">
<!--
body {
    background-color: #FFFFFF;
    margin-left: 0px;
}
.STYLE4 {color: #000000}
-->
</style>
</head>
<body topmargin="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="392" rowspan="2"><img src="Images/adminmain01.gif" width="392" height="126"></td>
    <td height="114" valign="top" background="Images/adminmain0line2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20"></td>
      </tr>
      <tr>
        <td><span class="STYLE4">��������</span></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><div id="peinfo1">���ڶ�ȡ������...</div><div id="peinfo2" style="z-index: 1; visibility: hidden; position: absolute"></div>
          <div id="peinfo5" style=" visibility: hidden;"></div>
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellpadding="3" cellspacing="0">
      
      <tr>
        <td width="31%" height="87" align="right"><table width="94%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><%=AdminName%>���ã� </td>
            </tr>
            <tr>
              <td valign="top">������
                <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script>
                �����У�</td>
            </tr>

            <tr>
              <td valign="top"><%=ShowUnPassedInfo%></td>
            </tr>
    </table></td>
        <td width="1%"><table width="3" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="3" height="65" bgcolor="#1890CC"></td>
            </tr>
        </table></td>
        <td width="68%">��ӭ������<%=SiteName%>��վ��̨����ϵͳ������������������ϵͳ�ṩ��ǿ���HTML���ɹ��ܣ���ݵĺ�̨�����ܣ���Ŀ���޼����࣬���������վƵ�����ܣ���Ŀ�������á������ƶ��ȹ�����Ч�ع�����վ����������ʱʹ�ö�����<font color="#FF0000">�ر�����</font>���ܹرջ����ߵĹ�����������չ�������档���μ�����վ������������Ϣ��</td>
      </tr>
      <tr>
        <td height="5" colspan="3"></td>
        </tr>
    </table></td>
  </tr>
</table>
<%
Dim rsArticleInfo, sqlArticleInfo, rsSoftInfo, sqlSoftInfo, rsGuestBookInfo, sqlGuestBookInfo, rsPhotoInfo, sqlPhotoInfo, rsChannelInfo, sqlChannelInfo, rsCommentInfo, sqlCommentInfo, rsUserInfo, sqlUserInfo
Dim channelIDinfo, channelIDinfonew, a1, a2, vNoApproveUser, vTestUser, c1, c2, cdir
Dim vArticleID, vChannelID, vChannelIDAry(50), vChannelName(50), vCommentCount(50), x, y
%>
<%
If ShowUnpass = True Then
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">��������</td>
    <td width="100"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">�������</td>
    <td width="100"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">����ͼƬ</td>
    <td width="100"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">��������</td>
    <td width="40"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">��������</td>
    <td width="40"></td>
    <td width="20" rowspan="2"></td>
    <td width="80" align="center" class="topbg">�����Ա</td>
    <td width="40"></td>
    <td width="20" rowspan="2"></td>
  </tr>
  <tr>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
    <td height="1" colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20"></td>
    <td width="180" valign="top"><% ShowUnPassedChannel(1) %></td>
    <td width="20"></td>
    <td width="180" valign="top"><% ShowUnPassedChannel(2) %></td>
    <td width="20"></td>
    <td width="180" valign="top"><% ShowUnPassedChannel(3) %></td>
    <td width="20"></td>
    <td width="120" valign="top"><% ShowUnPassedGuestBook() %></td>
    <td width="20"></td>
    <td width="120" valign="top"><% ShowUnPassedComment() %></td>
    <td width="20"></td>
    <td width="120" valign="top"><% ShowUnPassedUser() %></td>
  </tr>
</table>

<%
End If
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="400" align="center" class="topbg"><span class="Glow">�� վ �� �� �� �� �� ��</span></td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="400" align="center" class="topbg"><span class="Glow">�� �� �� �� �� �� �� ��</span></td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1"></td>
    <td></td>
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
        <td width="100" align="center" class="topbg"><a class='Class' href="Admin_SiteConfig.asp">��վ��Ϣ����</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
        <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Class.asp?ChannelID=1" target="main">��վ��Ŀ����</a></td>
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
      ������ݲ˵���<a href="Admin_SiteConfig.asp" target="main"><u><font color="#FF0000">��վ��Ϣ����</font></u></a>
      | <a href="Admin_SiteConfig.asp#SiteOption" target="main"><u><font color="#FF0000">��վѡ������</font></u></a>
    | <a href="Admin_SiteConfig.asp#User" target="main"><u><font color="#FF0000">�û�ѡ��</font></u></a></td>
    <td width="40">&nbsp;</td>
    <td width="400">����������վ��Ƶ���������õĸ�����Ŀ����Ŀ�����޼����๦�ܡ��ɶ���Ŀ������ӡ�ɾ�������򡢸�λ���ϲ����������õȹ���<br>
      ����<font color="#0000FF">���ΰ�װ���ڸ�Ƶ�����������Ŀ</font>��<br>
      ������ݲ˵���<a href="Admin_Class.asp?ChannelID=1" target="main"><u><font color="#FF0000">������Ŀ����</font></u></a>
      | <a href="Admin_Class.asp?ChannelID=2" target="main"><u><font color="#FF0000">������Ŀ����</font></u></a>
      | <a href="Admin_Class.asp?ChannelID=3" target="main"><u><font color="#FF0000">ͼƬ��Ŀ����</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_CreateSiteIndex.asp" target="main">��ҳ���ɹ���</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Article.asp?ChannelID=1&Action=Add&AddType=2" target="main">��վ�������</a></td>
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
    <td width="400">�����״ΰ�װϵͳ��<a href="Admin_CreateSiteIndex.asp" target="main"><font color="#FF0000">������վ��ҳ</font></a>����ҳ����Ŀҳ������ҳ��ר��ҳ����������������ȫ��HTMLҳ�棨���ۺ͵����ͳ�Ƴ��⣩����Ƶ���������ɹ�������<a href=Admin_Channel.asp target=main title="������վ����ϵͳ�У�Ƶ����ָĳһ����ģ��ļ��ϡ�ĳһƵ�������Ǿ߱�����ϵͳ���ܣ���߱�����ϵͳ��ͼƬϵͳ�Ĺ��ܡ�"><U><font color="#FF0000">��վƵ������</font></U></a>�����á�</td>
    <td width="40">&nbsp;</td>
    <td width="400">����ϵͳ�ṩǿ���<a href="#" title="���߱༭���ܹ�����ҳ��ʵ���������༭������磺Word�������е�ǿ����ӱ༭���ݵĹ��ܡ�"><u>���߱༭��</u></a>���������ģ������ӡ�ɾ�����޸���վ����Ƶ���¸���Ŀ��������ݣ����֡������ͼƬ�ȣ��������������ݵ�������Եȡ�<br>
      ������ݲ˵���<a href="Admin_Article.asp?ChannelID=1&Action=Add&AddType=2" target="main"><u><font color="#FF0000">�������</font></u>
      </a><a href="Admin_Article.asp?ChannelID=1&Action=Manage&Passed=All"><u><font color="#FF0000">����</font></u></a>
      | <a href="Admin_Soft.asp?ChannelID=2&Action=Add&AddType=2" target="main"><u><font color="#FF0000">������</font></u></a>
      <a href="Admin_Article.asp?ChannelID=2&Action=Manage&Passed=All"><u><font color="#FF0000">����</font></u></a> | <a href="Admin_Photo.asp?ChannelID=3&Action=Add&AddType=2" target="main"><u><font color="#FF0000">���ͼƬ</font></u></a>
      <a href="Admin_Article.asp?ChannelID=3&Action=Manage&Passed=All"><u><font color="#FF0000">����</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Admin.asp" target="main">����Ա����</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_GuestBook.asp?Passed=All" target="main">��վ���Թ���</a></td>
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
    <td width="400">����ǿ���<a href="Admin_Admin.asp" target="main"><U><font color="#FF0000">��վȨ�޹���</font></U></a>������ɾ����Ա��ָ����ϸ�Ĺ���Ȩ�ޣ�ʹ��վ�Ĺ���ּ�������˹�ͬ����������վ���<a href="Admin_Admin.asp?Action=Add" target="main" title="��������Ա��ӵ������Ȩ�ޡ�ĳЩȨ�ޣ������Ա������վ��Ϣ���á���վѡ�����õȹ���Ȩ�ޣ�ֻ�г�������Ա���С�"><U>��������Ա</U></a>��<a href="Admin_Admin.asp?Action=Add" target="main" title="��ͨ����Ա��ͱ��ָ��������վ�����ܣ���Ҫ��ϸָ��ÿһ�����Ȩ�ޡ�"><U>��ͨ����Ա</U></a>����Ҳ���������趨<a href="Admin_UserGroup.asp" target="main" title="�û������û��˻��ļ��ϣ�ͨ�������û��飬��������û������������Ȩ����Ȩ�ޡ������Ȩ�������ڡ�Ƶ����������Ƶ���ġ���Ŀ�����С�"><font color="#FF0000"><U>�û���</U></font></a>�Թ���ע���û�����</td>
    <td width="40">&nbsp;</td>
    <td width="400">�������û������Խ�����ˡ��޸ġ�ɾ�����ظ��Ȳ��������Ե���˹�������<a href=Admin_Channel.asp target=main title="��վ����ϵͳ�У�Ƶ����ָĳһ����ģ��ļ��ϡ�ĳһƵ�������Ǿ߱�����ϵͳ���ܣ���߱�����ϵͳ��ͼƬϵͳ�Ĺ��ܡ�"><U>��վƵ������</U></a>�����á�<br>
      ������ݲ˵���<a href="Admin_GuestBook.asp?Passed=False" target="main"><u><font color="#FF0000">�������</font></u></a>
      | <a href="Admin_GuestBook.asp?Passed=All" target="main"><u><font color="#FF0000">��������</font></u></a></td>
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
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Channel.asp" target="main" title="��վ����ϵͳ�У�Ƶ����ָĳһ����ģ��ļ��ϡ�ĳһƵ�������Ǿ߱�����ϵͳ���ܣ���߱�����ϵͳ��ͼƬϵͳ�Ĺ��ܡ�">��վƵ������</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class="Class" href="Admin_Advertisement.asp" target="main">��վ������</a></td>
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
    <td width="400">����<a href="Admin_Channel.asp" target="main">������վ�ĸ���Ƶ���Ĺ���ģ�飬�����¡����ء�ͼƬ�����Ե�Ƶ����</a>Ƶ���ɷ�Ϊ<a href="#" title="ϵͳ�ڲ�Ƶ��ָ������MY�������й���ģ�飨���š����¡�ͼƬ�ȣ�����������µ�Ƶ������Ƶ���߱�����ʹ�ù���ģ����ȫ��ͬ�Ĺ��ܡ�"><U>ϵͳ�ڲ�Ƶ��</U></a>��<a href="#" title="�ⲿƵ��ָ���ӵ�MY����ϵͳ����ĵ�ַ�С�����Ƶ��׼�����ӵ���վ�е�����ϵͳʱ"><U>�ⲿƵ��</U></a>��Ƶ������������ѡ��Ƶ��һ����Ӻ�Ͳ����ٸ���Ƶ�����͡�</td>
    <td width="40">&nbsp;</td>
    <td width="400">����ϵͳ�ṩǿ���<a href="Admin_Advertisement.asp" target=main title="���߱༭���ܹ�����ҳ��ʵ���������༭������磺Word�������е�ǿ����ӱ༭���ݵĹ��ܡ�"><u>��վ������</u></a>���ܣ�����վ�Ĺ���λ�����������ӡ��޸ġ�ɾ���Ȳ�����<br>
    ��ݲ˵���<a href="Admin_Advertisement.asp?Action=ZoneList" target="main"><u><font color="#FF0000">����λ����</font></u></a> | <a href="Admin_Advertisement.asp?Action=AddZone" target="main"><u><font color="#FF0000">����°�λ</font></u></a> | <a href="Admin_Advertisement.asp?Action=ADList" target="main"><u><font color="#FF0000">��վ������</font></u></a> | <a href="Admin_GuestBook.asp?Passed=All" target="main"></a><a href="Admin_Advertisement.asp?Action=AddAD" target="main"><u><font color="#FF0000">����¹��</font></u></a></td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 colspan=2 class="topbg"><span class="Glow">�� �� �� �� Ϣ</span>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">���������ͣ�      <%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td width="50%">�ű��������棺
    <%
    response.write ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion
    If CSng(ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion) < 5.6 Then
        response.write "&nbsp;&nbsp;<a href='http://www.microsoft.com/downloads/release.asp?ReleaseID=33136' target='_blank'><font color='green'>�汾���ͣ����˸���</font></a>"
    End If
    %>
    </td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">վ������·����      <%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td width="50%">���ݿ�ʹ�ã�<%ShowObjectInstalled("adodb.connection")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">FSO�ı���д��<%ShowObjectInstalled(objName_FSO)%></td>
    <td width="50%">��������д��<%ShowObjectInstalled("Adodb.Stream")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">XMLHTTP���֧�֣�<%ShowObjectInstalled("Microsoft.XMLHTTP")%></td>
    <td width="50%">XMLDOM���֧�֣�<%ShowObjectInstalled("Microsoft.XMLDOM")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">XML���֧�֣�<%ShowObjectInstalled("MSXML2.XMLHTTP")%></td>
    <td width="50%">AspJpeg���֧�֣�<%ShowObjectInstalled("Persits.Jpeg")%></td>
  </tr>
  
  <tr class="tdbg" height=23>
    <td width="50%">Jmail���֧�֣�<%ShowObjectInstalled("JMail.SMTPMail")%></td>
    <td width="50%">CDONTS���֧�֣�<%ShowObjectInstalled("CDONTS.NewMail")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%">ASPEMAIL���֧�֣�<%ShowObjectInstalled("Persits.MailSender")%></td>
    <td width="50%">WebEasyMail���֧�֣�<%ShowObjectInstalled("easymail.MailSend")%></td>
  </tr>
  <tr class="tdbg" height=23>
    <td width="50%"> </td>
    <td width="50%" align="right"><a href="Admin_ServerInfo.asp">��˲鿴����ϸ�ķ�������Ϣ&gt;&gt;&gt;</a></td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 class="topbg"><span class="Glow">Copyright 2003-2006 &copy; <%=SiteName%> All Rights Reserved.</span>
  </tr>
</table>
<div id="peinfo3" style="height:1;overflow=auto;visibility:hidden;">
<script src="http://www.powereasy.net/PowerEasy2006_Info.asp?Action=ShowAnnounce"></script>
</div>
<div id="peinfo4" style="height:1;overflow=auto;visibility:hidden;">
<script src="http://www.powereasy.net/PowerEasy2006_Info.asp?Action=ShowPatchAnnounce"></script>
</div>
<script language="JavaScript">
marqueesHeight=36;
scrillHeight=20;
scrillspeed=60;
stoptimes=50;
stopscroll=false;preTop=0;currentTop=0;stoptime=0;
peinfo1.scrollTop=0;
with (peinfo1)
{
  style.width=0;
  style.height=marqueesHeight;
  style.overflowX='visible';
  style.overflowY='hidden';
  noWrap=true;
  onmouseover=new Function("stopscroll=true");
  onmouseout=new Function("stopscroll=false");
}
function init_srolltext()
{
  peinfo2.innerHTML='';
  peinfo2.innerHTML+=peinfo3.innerHTML;
  peinfo1.innerHTML=peinfo3.innerHTML+peinfo3.innerHTML;
  setInterval("scrollUp()",scrillspeed);
}

function init_peifo()
{
  peinfo5.innerHTML=peinfo4.innerHTML;
}
function scrollUp()
{
  if(stopscroll==true) return;
  currentTop+=1;
  if(currentTop==scrillHeight)
  {
   stoptime+=1;
   currentTop-=1;
   if(stoptime==stoptimes) { currentTop=0; stoptime=0; }
  }
  else
  {
   preTop=peinfo1.scrollTop;
   peinfo1.scrollTop+=1;
   if(preTop==peinfo1.scrollTop){ peinfo1.scrollTop=peinfo2.offsetHeight-marqueesHeight; peinfo1.scrollTop+=1; }
  }
}
init_peifo();
setInterval("",1000);
init_srolltext();
</script>
</body>
</html>
<%
Call CloseConn


Function ShowUnPassedInfo()
    Dim rsCount, UnPassed_Article, UnPassed_Soft, UnPassed_Photo, UnPassed_Message, strInfo
    Set rsCount = Conn.Execute("select count(ArticleID) from PE_Article where Deleted=" & PE_False & " and Status=0")
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Article = rsCount(0)
    Else
        UnPassed_Article = 0
    End If
    Set rsCount = Conn.Execute("select count(SoftID) from PE_Soft where Deleted=" & PE_False & " and Status=0")
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Soft = rsCount(0)
    Else
        UnPassed_Soft = 0
    End If
    Set rsCount = Conn.Execute("select count(PhotoID) from PE_Photo where Deleted=" & PE_False & " and Status=0")
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Photo = rsCount(0)
    Else
        UnPassed_Photo = 0
    End If
    Set rsCount = Conn.Execute("select count(GuestID) from PE_GuestBook where GuestIsPassed=" & PE_False)
    If Not (rsCount.EOF And rsCount.bof) Then
        UnPassed_Message = rsCount(0)
    Else
        UnPassed_Message = 0
    End If
    Set rsCount = Nothing
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>�������£�"
    strInfo = strInfo & "<font color=red>" & UnPassed_Article & "</font>ƪ&nbsp;&nbsp;"
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>�������أ�"
    strInfo = strInfo & "<font color=red>" & UnPassed_Soft & "</font>��<br>"
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>����ͼƬ��"
    strInfo = strInfo & "<font color=red>" & UnPassed_Photo & "</font>��&nbsp;&nbsp;"
    strInfo = strInfo & "<img src='Images/img_u.gif' align='absmiddle'>�������ԣ�"
    strInfo = strInfo & "<font color=red>" & UnPassed_Message & "</font>��"
    ShowUnPassedInfo = strInfo
End Function

Sub ShowObjectInstalled(strObjName)
    If IsObjInstalled(strObjName) Then
        response.write "<b>��</b>"
    Else
        response.write "<font color='red'><b>��</b></font>"
    End If
End Sub

Function ShowUnPassedChannel(ChannelModuleType)
    Dim rsChannel, rsChannelCount, ModuleName, NoneInfo
    
    Select Case PE_CLng(ChannelModuleType)
    Case 1
        ModuleName = "Article"
    Case 2
        ModuleName = "Soft"
    Case 3
        ModuleName = "Photo"
    Case Else
    
    End Select
    NoneInfo = True
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName from PE_Channel where ModuleType = " & PE_CLng(ChannelModuleType))
    Do While Not rsChannel.EOF
        Set rsChannelCount = Conn.Execute("select count(" & ModuleName & "ID) from PE_" & ModuleName & " where Deleted=" & PE_False & " and ChannelID = " & rsChannel("ChannelID") & " and Status=0")
        If rsChannelCount(0) > 0 Then
            response.write "<a href=Admin_" & ModuleName & ".asp?ChannelID=" & PE_CLng(rsChannel("ChannelID")) & "&Action=Manage&ClassID=0&SpecialID=0&Status=0 target=main>" & rsChannel("ChannelName") & "</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]  "
        NoneInfo = False
        End If
    rsChannel.movenext
    Loop
    If NoneInfo = True Then
        response.write "û�д������Ϣ"
    End If
    rsChannelCount.Close
    Set rsChannelCount = Nothing
    rsChannel.Close
    Set rsChannel = Nothing
    
End Function

Function ShowUnPassedComment()
    Dim rsChannelCount, rs, rsChannel
    Set rsChannel = Conn.Execute("select count(CommentID) from PE_Comment where Passed =" & PE_False)
    If rsChannel(0) > 0 Then
        response.write "���������:[<span style='color:#ff0000'>" & rsChannel(0) & "</span>]"
    Else
        response.write "û�д��������"
    End If
    Set rsChannelCount = Conn.Execute("select top 20 * from PE_Comment where Passed = " & PE_False)
    rsChannel.Close
    Set rsChannel = Nothing
End Function


Function ShowUnPassedGuestBook()
    Dim rsChannelCount
    Set rsChannelCount = Conn.Execute("select count(GuestID) from PE_GuestBook where GuestIsPassed =" & PE_False)
    If rsChannelCount(0) > 0 Then
        response.write "<a href='Admin_GuestBook.asp?Passed=False' target=main>���԰�</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]"
    Else
        response.write "û�д��������"
    End If
    rsChannelCount.Close
    Set rsChannelCount = Nothing
End Function

Function ShowUnPassedUser()
    Dim rsChannelCount
    Set rsChannelCount = Conn.Execute("select count(UserID) from PE_User where GroupID = 7")
    If rsChannelCount(0) > 0 Then
        response.write "<a href='Admin_User.asp?SearchType=11&GroupID=7' target=main>��������Ա</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]<br>"
    Else
        response.write "û�д�������Ա<br>"
    End If
    Set rsChannelCount = Conn.Execute("select count(UserID) from PE_User where GroupID = 8")
    If rsChannelCount(0) > 0 Then
        response.write "<a href='Admin_User.asp?SearchType=11&GroupID=8' target=main>δ��֤��Ա</a>:[<span style='color:#ff0000'>" & rsChannelCount(0) & "</span>]"
    Else
        response.write "û��δ��֤��Ա"
    End If
        
    rsChannelCount.Close
    Set rsChannelCount = Nothing
End Function

%>
