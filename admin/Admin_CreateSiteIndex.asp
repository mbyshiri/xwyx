<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Class.asp"-->
<!--#include file="../Include/PowerEasy.Special.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<!--#include file="../Include/PowerEasy.Photo.asp"-->
<!--#include file="../Include/PowerEasy.Product.asp"-->
<!--#include file="../Include/PowerEasy.SiteIndex.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
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

ChannelID = 0
Response.Write "<html><head><title>生成网站首页</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf

If FileName_SiteIndex = "Index.asp" Then
    Response.Write "因为网站配置中未启用网站首页生成HTML功能，所以不用生成首页。"
    Response.End
End If
Response.Write "正在生成网站首页（" & InstallDir & FileName_SiteIndex & "）……"

Call GetHTML_SiteIndex

Call WriteToFile(InstallDir & FileName_SiteIndex, strHTML)
Response.Write "……………………生成网站首页成功！" & "</b>" & vbCrLf
Response.Write "</body></html>"
Call CloseConn
%>
