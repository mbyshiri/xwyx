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

ChannelID = PE_CLng(Trim(Request("ChannelID")))
Dim PrevSearchTime
PrevSearchTime = Trim(Session("Search_Time"))
If PrevSearchTime <> "" Then
    PrevSearchTime = CDate(PrevSearchTime)
    If DateDiff("s", PrevSearchTime, Now) < SearchInterval Then
        Response.Write "<br><br><br><p align='center'>Ϊ�˱���������������Ĵ���ϵͳ��Դ�������� " & SearchInterval & " �����ˢ�±�ҳ��</p>"
        Response.End
    End If
End If
Dim sModuleName
sModuleName = LCase(Trim(Request("ModuleName")))
ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
SkinID = DefaultSkinID
PageTitle = "�������"
strPageTitle = SiteName & "----" & PageTitle
strFileName = "Search.asp?ModuleName=" & sModuleName & "&ChannelID=" & ChannelID & "&Field=" & strField & "&Keyword=" & Keyword & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID
MaxPerPage = MaxPerPage_SearchResult

Dim PE_Search
Select Case sModuleName
Case "article"
    Set PE_Search = New Article
Case "soft"
    Set PE_Search = New Soft
Case "photo"
    Set PE_Search = New Photo
Case "shop"
    Set PE_Search = New Product
Case Else
    Set PE_Search = New Article
End Select
Call PE_Search.Init
Call PE_Search.GetHtml_Search
Response.Write strHtml
Set PE_Search = Nothing
Call CloseConn
Session("Search_Time") = Now
%>
