<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Response.Expires = -1

'���RSS����
Dim RssDOM, RssNode, RssSubNode, RssTempNode
Set RssDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
RssDOM.appendChild (RssDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""gb2312"""))
RssDOM.appendChild (RssDOM.createElement("rss"))
RssDOM.documentElement.Attributes.setNamedItem(RssDOM.createNode(2, "version", "")).text = "2.0"
Set RssNode = RssDOM.createNode(1, "channel", "")
RssDOM.documentElement.appendChild (RssNode)
Set RssSubNode = RssNode.appendChild(RssDOM.createElement("title"))
RssSubNode.text = "RssRebuder"
Set RssSubNode = RssNode.appendChild(RssDOM.createElement("language"))
RssSubNode.text = "zh-cn"
Set RssTempNode = RssNode

Dim urlReceived, rootNode, ListNum, Tlength, FeedUrl, RSSVersion
Set urlReceived = CreateObject("Microsoft.XMLDOM")
urlReceived.async = False
urlReceived.Load Request

Set rootNode = urlReceived.getElementsByTagName("root")
If rootNode.length < 1 Then
    Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
    RssNode.text = "�������ݴ�!"
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
Else
    ListNum = rootNode(0).selectSingleNode("listnum").text
    Tlength = rootNode(0).selectSingleNode("titlelength").text
    FeedUrl = rootNode(0).selectSingleNode("feedurl").text
    If ListNum = "" Or ListNum < 1 Then
        ListNum = 10
    Else
        ListNum = CLng(ListNum)
    End If
    If Tlength = "" Then
        Tlength = 35
    Else
        Tlength = CLng(Tlength)
    End If
End If
Set urlReceived = Nothing

If FeedUrl = "" Then
    Set RssNode = RssTempNode
    Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
    RssNode.text = "RssԴ��ַΪ��..."
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
Else
    Dim XmlRss, XMLDOM, i, j, oItem
    On Error Resume Next
    Set XmlRss = Server.CreateObject("MSXML2.ServerXMLHTTP")
    XmlRss.SetTimeouts 5000, 5000, 120000, 60000
    XmlRss.Open "GET", FeedUrl, False
    XmlRss.Send
    If Err.Number <> 0 Then
        Set RssNode = RssTempNode
        Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
        RssNode.text = "��������ʱ"
        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
        Err.Clear
    Else
        If XmlRss.ReadyState <> 4 Or Trim(XmlRss.responseText & "") = "" Then
            Set RssNode = RssTempNode
            Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
            RssNode.text = "����������Ӧ"
            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
        Else
            Set XMLDOM = Server.CreateObject("microsoft.XMLDOM")
            XMLDOM.async = False
            XMLDOM.Load (XmlRss.responseXML)
            If XMLDOM.ReadyState = 4 Then
                Set rootNode = XMLDOM.documentElement
                Select Case rootNode.nodename
                Case "rss"
                    RSSVersion = rootNode.getattribute("version")
                    If RSSVersion = "2.0" Then
                        Set oItem = XMLDOM.getElementsByTagName("item")
                        If oItem.length < 1 Then
                            Set RssNode = RssTempNode
                            Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                            RssNode.text = "��δ��������"
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                        Else
                            If oItem.length > ListNum Then
                                j = ListNum - 1
                            Else
                                j = oItem.length - 1
                            End If
                            For i = 0 To j
                                Set RssNode = RssTempNode
                                Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                                RssNode.text = GetSubStr(oItem(i).selectSingleNode("title").text, Tlength, True)
                                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                                RssNode.text = oItem(i).selectSingleNode("link").text
                                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                                RssNode.text = oItem(i).selectSingleNode("description").text
                                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                                RssNode.text = oItem(i).selectSingleNode("author").text
                                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                                RssNode.text = oItem(i).selectSingleNode("category").text
                                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                                RssNode.text = oItem(i).selectSingleNode("pubDate").text
                            Next
                        End If
                    Else
                        Response.Write "<item><title>RSS���ݰ汾����!</title><link /><description /><author /><category /><pubDate /></item>"
                    End If
                Case "rdf:RDF"
                    RSSVersion = "1.0"
                    Set oItem = XMLDOM.getElementsByTagName("item")
                    If oItem.length < 1 Then
                        Set RssNode = RssTempNode
                        Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                        RssNode.text = "������"
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                    Else
                        If oItem.length > ListNum Then
                            j = ListNum - 1
                        Else
                            j = oItem.length - 1
                        End If
                        For i = 0 To j
                            Set RssNode = RssTempNode
                            Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                            RssNode.text = GetSubStr(oItem(i).selectSingleNode("title").text, Tlength, True)
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                            RssNode.text = oItem(i).selectSingleNode("link").text
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                            RssNode.text = oItem(i).selectSingleNode("description").text
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                        Next
                    End If
                Case "feed"
                    RSSVersion = "atom"
                    Set oItem = XMLDOM.getElementsByTagName("entry")
                    If oItem.length < 1 Then
                        Set RssNode = RssTempNode
                        Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                        RssNode.text = "������"
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                    Else
                        If oItem.length > ListNum Then
                            j = ListNum - 1
                        Else
                            j = oItem.length - 1
                        End If
                        For i = 0 To j
                            Set RssNode = RssTempNode
                            Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                            RssNode.text = GetSubStr(oItem(i).selectSingleNode("title").text, Tlength, True)
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                            RssNode.text = oItem(i).selectSingleNode("link").getattribute("href")
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                            If oItem(i).selectSingleNode("summary").text <> "" Then
                                RssNode.text = oItem(i).selectSingleNode("summary").text
                            ElseIf oItem(i).selectSingleNode("content").text <> "" Then
                                RssNode.text = oItem(i).selectSingleNode("content").text
                            End If
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                            Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                        Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                        Next
                    End If
                Case Else
                    Set RssNode = RssTempNode
                    Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                    RssNode.text = "δ֪������Դ��ʽ!"
                    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                    Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
                End Select
            Else
                Set RssNode = RssTempNode
                Set RssSubNode = RssNode.appendChild(RssDOM.createElement("item"))
                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("title"))
                RssNode.text = "����Դ��ȡ��!"
                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("link"))
                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("description"))
                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("author"))
                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("category"))
                Set RssNode = RssSubNode.appendChild(RssDOM.createElement("pubDate"))
            End If
            Set XMLDOM = Nothing
        End If
        Set rootNode = Nothing
        Set XmlRss = Nothing
    End If
End If
Response.ContentType = "text/xml; charset=gb2312"
Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>" & vbCrLf
Response.Write RssDOM.documentElement.xml

'**************************************************
'��������GetSubStr
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'        strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'**************************************************
Function GetSubStr(ByVal str, ByVal strlen, bShowPoint)
    If str = "" Then
        GetSubStr = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(str)
    t = 0
    strTemp = str
    If strlen = "" Then
        strlen = 0
    Else
        strlen = CLng(strlen)
    End If
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next
    If strTemp <> str And bShowPoint = True Then
        strTemp = strTemp & "��"
    End If
    GetSubStr = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function
%>
