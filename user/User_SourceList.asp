<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.SourceList.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ChannelID, TypeSelect, Group, strTypeName, AllKeyList, AllUserList

MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
If MaxPerPage <= 0 Then MaxPerPage = 40
TypeSelect = ReplaceBadChar(Trim(Request("TypeSelect")))
Group = ReplaceBadChar(Trim(Request("Group")))
ChannelID = PE_CLng(Trim(Request("ChannelID")))
FileName = "User_SourceList.asp"
strFileName = "User_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=" & Group & "&KeyWord=" & Keyword

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>选择对话框</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<base target='_self'>"
Response.Write "<link href='../Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<form method='post' name='myform' action=''>" & vbCrLf
Select Case TypeSelect
Case "UserList"
    strTypeName = "会员列表"
    Call UserList
Case "KeyList"
    strTypeName = "关键字"
    Call Key
Case "AuthorList"
    strTypeName = "作者"
    Call Author
Case "CopyFromList"
    strTypeName = "来源"
    Call CopyFrom
Case "AgentList"
    strTypeName = "代理商"
    Call AgentList
Case Else
    Response.Write "参数丢失"
End Select
Response.Write "</form></body></html>"
Call CloseConn

%>
