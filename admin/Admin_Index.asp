<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Call CloseConn
Dim AdminName,AdminPassword,RndPassword,AdminLoginCode
AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
AdminLoginCode = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminLoginCode")))
If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Or (EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode) Then
    Response.redirect "Admin_login.asp"
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=SiteName & "--后台管理"%></title>
</head>
<frameset rows="*" cols="200,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
  <frame name="left" scrolling="yes" marginwidth="0" marginheight="0" src="Admin_Index_Left.asp">
  <frameset rows="53,*" cols="*" framespacing="0" border="false" rows="35,*" frameborder="0" scrolling="yes">
    <frame name="top" scrolling="no" src="Admin_Index_Top.asp">
    <frame name="main" scrolling="auto" src="Admin_Index_Main.asp">
  </frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
  </body>
</noframes>
</html>