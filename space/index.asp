<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Set xmlconfig = Nothing
If Action <> "xml" Then strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""index.xsl"" version=""1.0""?>"

Dim i, nodeCount, BlogSql, PE_Hits, nodeLis
Set xmlconfig = Server.CreateObject("Microsoft.XMLDOM")
xmlconfig.async = False
xmlconfig.Load (Server.MapPath("config.xml"))
Set bootnode = xmlconfig.getElementsByTagName("item")
nodeCount = bootnode.length
If nodeCount > 0 Then
	PE_Hits = XmlText("ShowSource", "Space/HitsOfHot", "100")
	On Error Resume Next
	For i = 1 To nodeCount
		Set SubNode = bootnode.nextNode()
		BlogSql = SubNode.selectSingleNode("sql").Text
		If Trim(SubNode.selectSingleNode("sqlwhere").Text & "") <> "" Then
			BlogSql = BlogSql & " where " & SubNode.selectSingleNode("sqlwhere").Text
			If TypeID > 0 Then
				BlogSql = BlogSql & " and A.ClassID=" & TypeID
			End If
		Else
			If TypeID > 0 Then
				BlogSql = BlogSql & " where A.ClassID=" & TypeID
			End If
		End If
		If Trim(SubNode.selectSingleNode("sqlorder").Text & "") <> "" Then
			BlogSql = BlogSql & " order by " & SubNode.selectSingleNode("sqlorder").Text
		End If
		BlogSql = Replace(Replace(Replace(Replace(Replace(Replace(BlogSql, "PE_True", PE_True), "PE_False", PE_False), "PE_OrderType", PE_OrderType), "PE_Now", PE_Now), "PE_DatePart_D", PE_DatePart_D), "PE_Hits", PE_Hits)
		Call GetBlogItem(SubNode.selectSingleNode("name").Text, BlogSql)
	Next
End If
Set xmlconfig = Nothing

'����ۺ��������б�
Call GetBlogClassList

'��������б�
Call GetAnnounceList

'���Ƶ���б�
Call GetChannelList

strtmp = strtmp & XMLDOM.documentElement.xml

Set Node = Nothing
Set SubNode = Nothing
Set XMLDOM = Nothing

Response.Write strtmp
Call CloseConn
%>