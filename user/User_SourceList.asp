<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.SourceList.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
Response.Write "<title>ѡ��Ի���</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<base target='_self'>"
Response.Write "<link href='../Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<form method='post' name='myform' action=''>" & vbCrLf
Select Case TypeSelect
Case "UserList"
    strTypeName = "��Ա�б�"
    Call UserList
Case "KeyList"
    strTypeName = "�ؼ���"
    Call Key
Case "AuthorList"
    strTypeName = "����"
    Call Author
Case "CopyFromList"
    strTypeName = "��Դ"
    Call CopyFrom
Case "AgentList"
    strTypeName = "������"
    Call AgentList
Case Else
    Response.Write "������ʧ"
End Select
Response.Write "</form></body></html>"
Call CloseConn

%>
