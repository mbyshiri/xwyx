<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Edition.asp"-->
<!--#include file="../Include/PowerEasy.Version.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 0      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限    
Call CloseConn

Call CloseConn

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>顶部管理导航菜单</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<style type='text/css'>" & vbCrLf
Response.Write "a:link { color:#ffffff;text-decoration:none}" & vbCrLf
Response.Write "a:hover {color:#ffffff;}" & vbCrLf
Response.Write "a:visited {color:#f0f0f0;text-decoration:none}" & vbCrLf
Response.Write ".spa {FONT-SIZE: 9pt; FILTER: Glow(Color=#0F42A6, Strength=2) dropshadow(Color=#0F42A6, OffX=2, OffY=1,); COLOR: #8AADE9; FONT-FAMILY: '宋体'}" & vbCrLf
Response.Write "img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}" & vbCrLf
Response.Write "</style>" & vbCrLf
Response.Write "<base target='main'>" & vbCrLf
Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
Response.Write "function preloadImg(src) {" & vbCrLf
Response.Write "  var img=new Image();" & vbCrLf
Response.Write "  img.src=src" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "preloadImg('Images/admin_top_open.gif');" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "var displayBar=true;" & vbCrLf
Response.Write "function switchBar(obj) {" & vbCrLf
Response.Write "  if (displayBar) {" & vbCrLf
Response.Write "    parent.frame.cols='0,*';" & vbCrLf
Response.Write "    displayBar=false;" & vbCrLf
Response.Write "    obj.src='Images/admin_top_open.gif';" & vbCrLf
Response.Write "    obj.title='打开左边管理导航菜单';" & vbCrLf
Response.Write "  } else {" & vbCrLf
Response.Write "    parent.frame.cols='200,*';" & vbCrLf
Response.Write "    displayBar=true;" & vbCrLf
Response.Write "    obj.src='Images/admin_top_close.gif';" & vbCrLf
Response.Write "    obj.title='关闭左边管理导航菜单';" & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body background='Images/admin_top_bg.gif' leftmargin='0' topmargin='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
Response.Write "  <tr valign='middle'>" & vbCrLf
Response.Write "    <td width=60><img onclick='switchBar(this)' src='Images/admin_top_close.gif' title='关闭左边管理导航菜单' style='cursor:hand'></td>" & vbCrLf
If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "ModifyPwd") = True Then
    Response.Write "    <td width=92><a href='Admin_ModifyPwd.asp'><img src='Images/top_an_1.gif' border='0'></a></td>" & vbCrLf
End If
If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "MailList") = True Then
    Response.Write "    <td width=92><a href='Admin_MailList.asp'><img src='Images/top_an_2.gif' border='0'></a></td>" & vbCrLf
End If
Response.Write "    <td width=104><a href='../User/User_Message.asp' target='_blank'><img src='Images/top_an_4.gif' border='0'></a></td>" & vbCrLf
If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Cache") = True Then
    Response.Write "    <td width=92><a href='Admin_Cache.asp'><img src='Images/top_an_5.gif' border='0'></a></td>" & vbCrLf
End If
Response.Write "    <td width=92><a href='http://help.powereasy.net'  target='_blank'><img src='Images/top_an_6.gif' border='0'></a></td>" & vbCrLf
Response.Write "    <td align='right' class='spa'>PowerEasy&reg; SiteWeaver&#8482; " & SystemEdition & " " & SystemVersion & " "&SystemDatabaseType &" Build " & SystemBuildDate & "</td>"
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "<script src='http://update.powereasy.net/Version.asp?Trade=SiteWeaver&SystemVersion=" & SystemVersion & "&SystemBuildDate=" & SystemBuildDate & "'></script>"
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf
%>