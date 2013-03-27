<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

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
ObjTotest(9, 1) = "(FSO 文本文件读写)"
ObjTotest(10, 0) = "adodb.connection"
ObjTotest(10, 1) = "(ADO 数据对象)"
    
ObjTotest(11, 0) = "SoftArtisans.FileUp"
ObjTotest(11, 1) = "(SA-FileUp 文件上传)"
ObjTotest(12, 0) = "SoftArtisans.FileManager"
ObjTotest(12, 1) = "(SoftArtisans 文件管理)"
ObjTotest(13, 0) = "LyfUpload.UploadFile"
ObjTotest(13, 1) = "(刘云峰的文件上传组件)"
ObjTotest(14, 0) = "Persits.Upload.1"
ObjTotest(14, 1) = "(ASPUpload 文件上传)"
ObjTotest(15, 0) = "w3.upload"
ObjTotest(15, 1) = "(Dimac 文件上传)"

ObjTotest(16, 0) = "JMail.SmtpMail"
ObjTotest(16, 1) = "(Dimac JMail 邮件收发) <a href='http://www.ajiang.net'>中文手册下载</a>"
ObjTotest(17, 0) = "CDONTS.NewMail"
ObjTotest(17, 1) = "(虚拟 SMTP 发信)"
ObjTotest(18, 0) = "Persits.MailSender"
ObjTotest(18, 1) = "(ASPemail 发信)"
ObjTotest(19, 0) = "SMTPsvg.Mailer"
ObjTotest(19, 1) = "(ASPmail 发信)"
ObjTotest(20, 0) = "DkQmail.Qmail"
ObjTotest(20, 1) = "(dkQmail 发信)"
ObjTotest(21, 0) = "Geocel.Mailer"
ObjTotest(21, 1) = "(Geocel 发信)"
ObjTotest(22, 0) = "IISmail.Iismail.1"
ObjTotest(22, 1) = "(IISmail 发信)"
ObjTotest(23, 0) = "SmtpMail.SmtpMail.1"
ObjTotest(23, 1) = "(SmtpMail 发信)"
    
ObjTotest(24, 0) = "SoftArtisans.ImageGen"
ObjTotest(24, 1) = "(SA 的图像读写组件)"
ObjTotest(25, 0) = "W3Image.Image"
ObjTotest(25, 1) = "(Dimac 的图像读写组件)"

Public IsObj, VerObj

'检查预查组件支持情况及版本

Dim i
For i = 0 To 25
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Dim TestObj
    Set TestObj = server.CreateObject(ObjTotest(i, 0))
    If -2147221005 <> Err Then      '感谢网友iAmFisher的宝贵建议
        IsObj = True
        VerObj = TestObj.version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
    ObjTotest(i, 2) = IsObj
    ObjTotest(i, 3) = VerObj
Next

'检查组件是否被支持及组件版本的子程序
Sub ObjTest(strObj)
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Dim TestObj
    Set TestObj = server.CreateObject(strObj)
    If -2147221005 <> Err Then      '感谢网友iAmFisher的宝贵建议
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
<TITLE>服务器信息</TITLE>
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
        <td><%=AdminName%>您好，今天是
          <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">查看服务器信息</font></td>
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
    <td height=22 class="topbg"><strong><%=SiteName%>----服务器信息</strong></td>
  </tr>
  <tr>
    <td class="tdbg"><div align="right">特别感谢<a href="http://www.ajiang.net">【阿江守候】</a>提供此探针程序！</div>
      <font class=fonts>是否支持ASP</font> <br>
      出现以下情况即表示您的空间不支持ASP： <br>
       1、访问本文件时提示下载。 <br>
       2、访问本文件时看到类似“&lt;%@ Language="VBScript" %&gt;”的文字。 </td>
  </tr>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>服务器的有关参数</strong></td>
  </tr>
  <tr class="tdbg">
    <td width='350' align=left>&nbsp;服务器名</td>
    <td>&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;服务器IP</td>
    <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;服务器端口</td>
    <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;服务器时间</td>
    <td>&nbsp;<%=now%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;IIS版本</td>
    <td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;脚本超时时间</td>
    <td>&nbsp;<%=Server.ScriptTimeout%> 秒</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;本文件路径</td>
    <td>&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;服务器CPU数量</td>
    <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;服务器解译引擎</td>
    <td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;服务器操作系统</td>
    <td>&nbsp;<%=Request.ServerVariables("OS")%></td>
  </tr>
</table>

<br>
<font class=fonts>组件支持情况</font>
<%
Dim strClass
    strClass = Trim(Request.Form("classname"))
    If "" <> strClass Then
    Response.Write "<br>您指定的组件的检查结果："
    ObjTest (strClass)
      If Not IsObj Then
        Response.Write "<br><font color=red>很遗憾，该服务器不支持 " & strClass & " 组件！</font>"
      Else
        Response.Write "<br><font class=fonts>恭喜！该服务器支持 " & strClass & " 组件。该组件版本是：" & VerObj & "</font>"
      End If
      Response.Write "<br>"
    End If
    %>


<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>IIS自带的ASP组件</strong></td>
  </tr>
    <%For i=0 to 10%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>×</b></font>"
        Else
            Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>常见的文件上传和管理组件</strong></td>
  </tr>
    <%For i=11 to 15%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>×</b></font>"
        Else
            Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>常见的收发邮件组件</strong></td>
  </tr>
    <%For i=16 to 23%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>×</b></font>"
        Else
            Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>
<br>
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td colspan='2' height=22 class="topbg"><strong>图像处理组件</strong></td>
  </tr>
    <%For i=24 to 25%>
    <tr class=tdbg>
        <td width='350' align=left>&nbsp;<%=ObjTotest(i,0) & "&nbsp;" & ObjTotest(i,1)%></font></td>
        <td align=left>&nbsp;<%
        If Not ObjTotest(i, 2) Then
            Response.Write "<font color=red><b>×</b></font>"
        Else
            Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
        End If%></td>
    </tr>
    <%next%>
</table>

<br>
<font class=fonts>其他组件支持情况检测</font><br>
在下面的输入框中输入你要检测的组件的ProgId或ClassId?
<table width=100% border="0" cellpadding="0" cellspacing="0" class="border" style="border-collapse: collapse">
<FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
    <tr height="18" class=tdbg>
        
      <td height=30 align="center">&nbsp;
        <input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value=" 确 定 " class=backc id=submit1 name=submit1>
<INPUT type=reset value=" 重 填 " class=backc id=reset1 name=reset1>
</td>
    </tr>
</FORM>
</table>
<br>
<font class=fonts>ASP脚本解释和运算速度测试</font><br>
我们让服务器执行50万次“1＋1”的计算，记录其所使用的时间。
<table width=100% border=0 cellpadding=3 cellspacing=1 class="border">
  <tr align="center">
    <td height=22 class="topbg"><strong>服&nbsp;&nbsp;&nbsp;务&nbsp;&nbsp;&nbsp;器</strong></td>
    <td height=22 class="topbg"><strong>完成时间</strong></td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;中国频道虚拟主机（2002-08-06 9:29）</td><td>&nbsp;610.9 毫秒</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;西部数码west263主机（2002-08-06 9:29）</td><td>&nbsp;357.8 毫秒</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;商务中国虚拟主机（2002-08-06 9:29）</td><td>&nbsp;353.1 毫秒</td>
  </tr>
  <tr class="tdbg">
    <td align=left>&nbsp;顶尖科技tonydns主机（2002-10-13 14:19）</td><td>&nbsp;303.2 毫秒</td>
  </tr>
  <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
<%

    '感谢网际同学录 http://www.5719.net 推荐使用timer函数
    '因为只进行50万次计算，所以去掉了是否检测的选项而直接检测
    
    Dim t1, t2, lsabc, thetime
    t1 = Timer
    For i = 1 To 500000
        lsabc = 1 + 1
    Next
    t2 = Timer

    thetime = CStr(Int(((t2 - t1) * 10000) + 0.5) / 10)
%>
  <tr class="tdbg">
    <td align=left>&nbsp;<font color=red>您正在使用的这台服务器</font>&nbsp;</td><td>&nbsp;<font color=red><%=thetime%> 毫秒</font></td>
  </tr>
  </form>
</table>
<br>
<div align="center"><a href="Admin_Index_Main.asp">【返回管理首页】</a></div>
</BODY>
</HTML>