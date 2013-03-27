<!-- #include File="Start.asp" -->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Common.Rss.asp"-->
<!--#include file="Include/PowerEasy.Common.Content.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

If RssCodeType = True Then
    Response.ContentType = "text/xml; charset=gb2312"
Else
    Response.ContentType = "text/xml; charset=utf-8"
End If

Dim ClassID, SpecialID, BlogID, Hot, Elite, AuthorName, OutNum, PriClassID, ClassField(5)
Dim sqlChannel, rsChannel
Dim SubNode, SiteLogoUrl
Dim strNoSee, strDefAuthor, strtmp

If Right(SiteUrl, 1) <> "/" Then SiteUrl = SiteUrl & "/"
SiteLogoUrl = SiteUrl & LogoUrl

ChannelID = PE_CLng(Trim(Request("ChannelID")))
ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
BlogID = PE_CLng(Trim(Request("BlogID")))

Hot = Trim(Request("Hot"))
If Hot = "" Then
    Hot = 0
Else
    Hot = PE_CLng(Hot)
End If
Elite = Trim(Request("Elite"))
If Elite = "" Then
    Elite = 0
Else
    Elite = PE_CLng(Elite)
End If
AuthorName = Trim(Request("AuthorName"))
If AuthorName = "" Then
    AuthorName = "none"
Else
    AuthorName = ReplaceBadChar(AuthorName)
End If

strNoSee = XmlText("Rss", "NoSee", "内容不可预览，请登录网站后查看。")
strDefAuthor = XmlText("BaseText", "DefAuthor", "佚名")
    
'输出RSS数据
Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")


If EnableRss = False Then
    If RssCodeType = True Then
        XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""gb2312"""))
    Else
        XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8"""))
    End If
    XMLDOM.appendChild (XMLDOM.createElement("rss"))
    XMLDOM.documentElement.Attributes.setNamedItem(XMLDOM.createNode(2, "version", "")).text = "2.0"

    Set Node = XMLDOM.createNode(1, "channel", "")
    XMLDOM.documentElement.appendChild (Node)
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("title"))
    SubNode.text = XmlText("Rss", "CloseEd", "本站已关闭RSS功能！")
Else
    If action = "diary" Then
        Call ShowOtherRss("diary")
    ElseIf action = "music" Then
        Call ShowOtherRss("music")
    ElseIf action = "book" Then
        Call ShowOtherRss("book")
    ElseIf action = "photo" Then
        Call ShowOtherRss("photo")
    ElseIf action = "link" Then
        Call ShowOtherRss("link")
    Else
        If ChannelID > 0 Then
            sqlChannel = "select ChannelID,OrderID,ChannelName,ChannelDir,ModuleType,UseCreateHTML,StructureType,FileNameType,FileExt_Item,Disabled,HitsOfHot from PE_Channel where ChannelID=" & ChannelID & " and Disabled = " & PE_False & " and ChannelType<2 order by OrderID"
            Set rsChannel = Conn.Execute(sqlChannel)
            If rsChannel.BOF And rsChannel.EOF Then
                If RssCodeType = True Then
                    XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""gb2312"""))
                Else
                    XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8"""))
                End If
                XMLDOM.appendChild (XMLDOM.createElement("rss"))
                XMLDOM.documentElement.Attributes.setNamedItem(XMLDOM.createNode(2, "version", "")).text = "2.0"
        
                Set Node = XMLDOM.createNode(1, "channel", "")
                XMLDOM.documentElement.appendChild (Node)
            
                Set SubNode = Node.appendChild(XMLDOM.createElement("title"))
                SubNode.text = XmlText("BaseText", "ChannelErr", "找不到指定的频道，或频道已被禁用！")
            Else
                ChannelName = rsChannel("ChannelName")
                UseCreateHTML = rsChannel("UseCreateHTML")

                Select Case rsChannel("ModuleType")
                Case 1
                    Call ShowArtcileRss(Hot, Elite, AuthorName, rsChannel("HitsOfHot"))
                Case 2
                    Call ShowSoftRss(Hot, Elite, AuthorName, rsChannel("HitsOfHot"))
                Case 3
                    Call ShowPhotoRss(Hot, Elite, AuthorName, rsChannel("HitsOfHot"))
                Case 4
                    Call ShowGuestRss
                Case 5
                    Call ShowProductRss(Hot, Elite, AuthorName)
                End Select
            End If
            rsChannel.Close
            Set rsChannel = Nothing
        Else
            If FileExt_SiteIndex < 4 And fso.FileExists(Server.MapPath(InstallDir & "xml/Rss.xml")) Then
                Call CloseConn
                Response.Redirect InstallDir & "xml/Rss.xml"
            Else
                Call ShowIndexRss(0)
            End If
        End If
    End If
End If

If Not IsNull(XMLDOM.documentElement.xml) Then
    If RssCodeType = True Then
        strtmp = "<?xml version=""1.0"" encoding=""gb2312""?>" & vbCrLf
        strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""rss.xsl"" version=""1.0""?>" & vbCrLf
        strtmp = strtmp & XMLDOM.documentElement.xml
    Else
        strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
        strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""rss.xsl"" version=""1.0""?>" & vbCrLf
        strtmp = strtmp & unicode(XMLDOM.documentElement.xml)
    End If
End If

Set Node = Nothing
Set SubNode = Nothing
Set XMLDOM = Nothing
Response.Write strtmp
Call CloseConn
%>
