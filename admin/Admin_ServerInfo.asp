<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 1      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim ObjTotest(26, 4)

ObjTotest(0, 0) = "MSWC.AdRotator"
ObjTotest(1, 0) = "MSWC.BrowserType"
ObjTotest(2, 0) = "MSWC.NextLink"
ObjTotest(3, 0) = "MSWC.Tools"
ObjTotest(4, 0) = "MSWC.Status"
ObjTotest(5, 0) = "MSWC.Counters"
ObjTotest(6, 0) = "IISSample.ContentRotator"
ObjTotest(7, 0) = "IISSample.PageCounter"
ObjTotest(8, 0) = "MSWC.PermissionChecker"
ObjTotest(9, 0) = "Scripting.FileSystemObject"
ObjTotest(9, 1) = "(FSO �ı��ļ���д)"
ObjTotest(10, 0) = "adodb.connection"
ObjTotest(10, 1) = "(ADO ���ݶ���)"
    
ObjTotest(11, 0) = "SoftArtisans.FileUp"
ObjTotest(11, 1) = "(SA-FileUp �ļ��ϴ�)"
ObjTotest(12, 0) = "SoftArtisans.FileManager"
ObjTotest(12, 1) = "(SoftArtisans �ļ�����)"
ObjTotest(13, 0) = "LyfUpload.UploadFile"
ObjTotest(13, 1) = "(���Ʒ���ļ��ϴ����)"
ObjTotest(14, 0) = "Persits.Upload.1"
ObjTotest(14, 1) = "(ASPUpload �ļ��ϴ�)"
ObjTotest(15, 0) = "w3.upload"
ObjTotest(15, 1) = "(Dimac �ļ��ϴ�)"

ObjTotest(16, 0) = "JMail.SmtpMail"
ObjTotest(16, 1) = "(Dimac JMail �ʼ��շ�) <a href='http://www.ajiang.net'>�����ֲ�����</a>"
ObjTotest(17, 0) = "CDONTS.NewMail"
ObjTotest(17, 1) = "(���� SMTP ����)"
ObjTotest(18, 0) = "Persits.MailSender"
ObjTotest(18, 1) = "(ASPemail ����)"
ObjTotest(19, 0) = "SMTPsvg.Mailer"
ObjTotest(19, 1) = "(ASPmail ����)"
ObjTotest(20, 0) = "DkQmail.Qmail"
ObjTotest(20, 1) = "(dkQmail ����)"
ObjTotest(21, 0) = "Geocel.Mailer"
ObjTotest(21, 1) = "(Geocel ����)"
ObjTotest(22, 0) = "IISmail.Iismail.1"
ObjTotest(22, 1) = "(IISmail ����)"
ObjTotest(23, 0) = "SmtpMail.SmtpMail.1"
ObjTotest(23, 1) = "(SmtpMail ����)"
    
ObjTotest(24, 0) = "SoftArtisans.ImageGen"
ObjTotest(24, 1) = "(SA ��ͼ���д���)"
ObjTotest(25, 0) = "W3Image.Image"
ObjTotest(25, 1) = "(Dimac ��ͼ���д���)"

Public IsObj, VerObj

'���Ԥ�����֧��������汾

Dim i
For i = 0 To 25
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Dim TestObj
    Set TestObj = server.CreateObject(ObjTotest(i, 0))
    If -2147221005 <> Err Then      '��л����iAmFisher�ı�����
        IsObj = True
        VerObj = TestObj.version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
    ObjTotest(i, 2) = IsObj
    ObjTotest(i, 3) = VerObj
Next

'�������Ƿ�֧�ּ�����汾���ӳ���
Sub ObjTest(strObj)
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Dim TestObj
    Set TestObj = server.CreateObject(strObj)
    If -2147221005 <> Err Then      '��л����iAmFisher�ı�����
        IsObj = True
        VerObj = TestObj.version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
End Sub
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
<TITLE>��������Ϣ</TITLE>
</HEAD>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�鿴��������Ϣ</font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=22 class="topbg"><strong><%=SiteName%>----��������Ϣ</strong></td>
  </tr>
  <tr>
    <td class="tdbg"><div align="right">�ر��л<a href="http://www.ajiang.net">�������غ�</a>�ṩ��̽�����</div>
      <font class=fonts>�Ƿ�֧��ASP</font> <br>
      ���������������ʾ���Ŀռ䲻֧��ASP�� <br>
       1�����ʱ��ļ�ʱ��ʾ���ء� <br>
       2�����ʱ��ļ�ʱ�������ơ�&lt;%@ Language="VBScript" %&gt;�������֡� </td>
  </tr>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>���������йز���</strong></td>
  </tr>
  <tr class="tdbg">
    <td width='350' align=left>&nbsp;��������</td>
    <td>&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;������IP</td>
    <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;�������˿�</td>
    <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;������ʱ��</td>
    <td>&nbsp;<%=now%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;IIS�汾</td>
    <td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;�ű���ʱʱ��</td>
    <td>&nbsp;<%=Server.ScriptTimeout%> ��</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;���ļ�·��</td>
    <td>&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;������CPU����</td>
    <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;��������������</td>
    <td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;����������ϵͳ</td>
    <td>&nbsp;<%=Request.ServerVariables("OS")%></td>
  </tr>
</table>

<br>
<font class=fonts>���֧�����</font>
<%
Dim strClass
    strClass = Trim(Request.Form("classname"))
    If "" <> strClass Then
    Response.Write "<br>��ָ��������ļ������"
    ObjTest (strClass)
      If Not IsObj Then
        Response.Write "<br><font color=red>���ź����÷�������֧�� " & strClass & " �����</font>"
      Else
        Response.Write "<br><font class=fonts>��ϲ���÷�����֧�� " & strClass & " �����������汾�ǣ�" & VerObj & "</font>"
      End If
      Response.Write "<br>"
    End If
    %>


<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>IIS�Դ���ASP���</strong></td>
  </tr>
    <%For i=0 to 10%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>��</b></font>"
        Else
            Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>�������ļ��ϴ��͹������</strong></td>
  </tr>
    <%For i=11 to 15%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>��</b></font>"
        Else
            Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>�������շ��ʼ����</strong></td>
  </tr>
    <%For i=16 to 23%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>��</b></font>"
        Else
            Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>ͼ�������</strong></td>
  </tr>
    <%For i=24 to 25%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>��</b></font>"
        Else
            Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>

<br>
<font class=fonts>�������֧��������</font><br>
��������������������Ҫ���������ProgId��ClassId?
<table width=100% border="0" cellpadding="0" cellspacing="0" class="border" style="border-collapse: collapse">
<FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
    <tr height="18" class=tdbg>
        
      <td height=30 align="center">&nbsp;
        <input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value=" ȷ �� " class=backc id=submit1 name=submit1>
<INPUT type=reset value=" �� �� " class=backc id=reset1 name=reset1>
</td>
    </tr>
</FORM>
</table>
<br>
<font class=fonts>ASP�ű����ͺ������ٶȲ���</font><br>
�����÷�����ִ��50��Ρ�1��1���ļ��㣬��¼����ʹ�õ�ʱ�䡣
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td height=22 class="topbg"><strong>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</strong></td>
    <td height=22 class="topbg"><strong>���ʱ��</strong></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;�й�Ƶ������������2002-08-06 9:29��</td><td>&nbsp;610.9 ����</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;��������west263������2002-08-06 9:29��</td><td>&nbsp;357.8 ����</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;�����й�����������2002-08-06 9:29��</td><td>&nbsp;353.1 ����</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;����Ƽ�tonydns������2002-10-13 14:19��</td><td>&nbsp;303.2 ����</td>
  </tr>
  <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
<%

    '��л����ͬѧ¼ http://www.5719.net �Ƽ�ʹ��timer����
    '��Ϊֻ����50��μ��㣬����ȥ�����Ƿ����ѡ���ֱ�Ӽ��
    
    Dim t1, t2, lsabc, thetime
    t1 = Timer
    For i = 1 To 500000
        lsabc = 1 + 1
    Next
    t2 = Timer

    thetime = CStr(Int(((t2 - t1) * 10000) + 0.5) / 10)
%>
  <tr class="tdbg">
    <td align=left>&nbsp;<font color=red>������ʹ�õ���̨������</font>&nbsp;</td><td>&nbsp;<font color=red><%=thetime%> ����</font></td>
  </tr>
  </form>
</table>
<br>
<div align="center"><a href="Admin_Index_Main.asp">�����ع�����ҳ��</a></div>
</BODY>
</HTML>