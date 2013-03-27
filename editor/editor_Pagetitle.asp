<%@language=vbscript codepage=936 %>
<%
option explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> 插入分页标题 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script event=onclick for=Ok language=JavaScript>
if (pagetitle.value!=null)
{
 var str;
 str="[NextPage" + pagetitle.value + "]";
 window.returnValue="<br>"+str+"<br><br>";
}
 window.close();
</SCRIPT>
</HEAD>
<BODY bgColor=#D4D0C8 topmargin=15 leftmargin=15 >
<FIELDSET align=left>
<LEGEND align=left><strong> 插入分页标题 </strong></LEGEND>
<TABLE>
<TR>
<TD>
  <Input TYPE="text" id="pagetitle" NAME="pagetitle" size=50></TD>
<TD>
  <Input TYPE="submit" value=" 确 定 " id="Ok">
  <Input TYPE="reset" value=" 取 消 " Onclick="window.close();">
</TD>
</TR>
</table>
</fieldset>
</BODY>
</HTML>