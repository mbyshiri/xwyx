<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<!--#include file="Include/PowerEasy.Class.asp"-->
<!--#include file="Include/PowerEasy.Special.asp"-->
<!--#include file="Include/PowerEasy.Article.asp"-->
<!--#include file="Include/PowerEasy.Soft.asp"-->
<!--#include file="Include/PowerEasy.Photo.asp"-->
<!--#include file="Include/PowerEasy.Product.asp"-->
<!--#include file="Include/PowerEasy.SiteIndex.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

ChannelID = 0
If fso.FileExists(Server.mappath("NotInsalled.txt")) Then
    Response.Write "<li>�������� <a href='Install.asp'>Install.asp</a> �Խ���ϵͳ��װ���̣�</li><br/><br/>"
    Response.Write "<li>������Ѿ��������д˳��򣬵���Ȼ���ִ���ʾ����ʹ��FTP�����ֶ�ɾ�� NotInstalled.txt �ļ���</li>"
    Response.End
End If

If FileName_SiteIndex <> "Index.asp" Then
    Call CloseConn
    Response.Redirect FileName_SiteIndex
Else
    If CurrentPage > 1 Or PE_Cache.CacheIsEmpty("Site_Index") Then
        Call GetHTML_SiteIndex
        If CurrentPage = 1 Then PE_Cache.SetValue "Site_Index", strHtml
    Else
        strHtml = PE_Cache.GetValue("Site_Index")
    End If
    Response.Write strHtml
End If
Call CloseConn
%>
