<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim tUploadDir
tUploadDir = Trim(Request("UploadDir"))
If ChannelID > 0 Then

Else
    If tUploadDir = "UploadAdPic" Then
        ChannelName = "网站广告"
    End If
End If
Call CloseConn

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>顶部管理导航菜单</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body background='Images/admin_top_bg.gif' leftmargin='0' topmargin='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'> " & vbCrLf
Response.Write "    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>" & ChannelName & "管理----上传文件管理" & "</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10011' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td height='30'><strong>管理说明：</strong>&nbsp;请打开左边上传目录导航，以便您更快捷的管理上传目录中的文件。</td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf
%>
