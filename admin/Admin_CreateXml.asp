<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim strHTML

Dim tempNode, CreatChannel, CreatGuest, CreatAuthor, CreatUser, CreatSite, CreatAnnounce

'检查管理员操作权限
If AdminPurview > 1 And CheckPurview_Other(AdminPurview_Others, "XML_" & ChannelDir) = False Then
    Response.Write "你没有此项操作的权限！"
    Response.End
End If

Action = Trim(Request("Action"))

Response.Write "<html><head><title>更新XML数据</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf

Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""gb2312"""))
If Action = "" Then
    CreatChannel = Trim(Request("Channel"))
    CreatGuest = Trim(Request("Guest"))
    CreatAuthor = Trim(Request("Author"))
    CreatUser = Trim(Request("User"))
    CreatSite = Trim(Request("Site"))
    CreatAnnounce = Trim(Request("Announce"))

    Response.Write "<br><br><b>正在生成XML数据输出页面..........<br>" & vbCrLf
    
    XMLDOM.appendChild (XMLDOM.createElement("powereasy"))
    XMLDOM.documentElement.Attributes.setNamedItem(XMLDOM.createNode(2, "version", "")).text = "2006"

    Set Node = XMLDOM.createNode(1, "SiteName", "")
    Node.text = SiteName
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "SiteTitle", "")
    Node.text = SiteTitle
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "SiteUrl", "")
    Node.text = SiteUrl
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "InstallDir", "")
    Node.text = InstallDir
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "LogoUrl", "")
    Node.text = LogoUrl
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "BannerUrl", "")
    Node.text = BannerUrl
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "WebmasterName", "")
    Node.text = WebmasterName
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "WebmasterEmail", "")
    Node.text = WebmasterEmail
    XMLDOM.documentElement.appendChild (Node)
    Set Node = XMLDOM.createNode(1, "Copyright", "")
    Node.text = Copyright
    XMLDOM.documentElement.appendChild (Node)
    
    
    If CreatChannel = True Then
        If ChannelID = 0 Then
            sqlChannel = "select ChannelID,OrderID,ChannelName,ChannelType,LinkUrl,ChannelDir,ModuleType,ReadMe,ChannelPicUrl,Disabled,ItemCount,ItemChecked,CommentCount from PE_Channel where Disabled = " & PE_False & " order by OrderID"
        Else
            sqlChannel = "select ChannelID,OrderID,ChannelName,ChannelType,LinkUrl,ChannelDir,ModuleType,ReadMe,ChannelPicUrl,Disabled,ItemCount,ItemChecked,CommentCount from PE_Channel where ChannelID=" & ChannelID & " and Disabled = " & PE_False & " and ChannelType=0 order by OrderID"
        End If
        Set rsChannel = Conn.Execute(sqlChannel)
        If rsChannel.BOF And rsChannel.EOF Then
            Set Node = XMLDOM.createNode(1, "errmsg", "")
            Node.text = "找不到指定的频道，或频道已被禁用！"
        Else
            Do While Not rsChannel.EOF
                Set Node = XMLDOM.createNode(1, "Channel", "")
                XMLDOM.documentElement.appendChild (Node)
                
                Set SubNode = XMLDOM.createNode(2, "ChannelID", "")
                    SubNode.text = rsChannel("ChannelID")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ChannelName", "")
                    SubNode.text = rsChannel("ChannelName")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ChannelReadMe", "")
                    If Not (rsChannel("ReadMe") = "" Or IsNull(rsChannel("ReadMe"))) Then SubNode.text = rsChannel("ReadMe")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ChannelPicUrl", "")
                    If Not (rsChannel("ChannelPicUrl") = "" Or IsNull(rsChannel("ChannelPicUrl"))) Then SubNode.text = rsChannel("ChannelPicUrl")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ModuleType", "")
                    SubNode.text = rsChannel("ModuleType")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ChannelType", "")
                    SubNode.text = rsChannel("ChannelType")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ItemCount", "")
                    SubNode.text = rsChannel("ItemCount")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "ItemChecked", "")
                    SubNode.text = rsChannel("ItemChecked")
                Node.Attributes.setNamedItem (SubNode)
                Set SubNode = XMLDOM.createNode(2, "CommentCount", "")
                    SubNode.text = rsChannel("CommentCount")
                Node.Attributes.setNamedItem (SubNode)
                    Set SubNode = XMLDOM.createNode(2, "LinkUrl", "")
                    If rsChannel("ChannelType") = 0 Then
                        SubNode.text = SiteUrl & "/" & rsChannel("ChannelDir")
                    Else
                        If Not (rsChannel("LinkUrl") = "" Or IsNull(rsChannel("LinkUrl"))) Then SubNode.text = rsChannel("LinkUrl")
                    End If
                Node.Attributes.setNamedItem (SubNode)

                If rsChannel("ChannelType") = 0 Then
                    If rsChannel("ModuleType") = 4 Then
                        If CreatGuest = True Then Call ShowGuestClass
                    Else
                        Call ShowClass(rsChannel("ChannelID"), rsChannel("ModuleType"))
                    End If
                    rsChannel.MoveNext
                Else
                    rsChannel.MoveNext
                End If
            Loop
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    
    If ChannelID = 0 Then
        If CreatUser = True Then Call ShowUserXml
        If CreatAuthor = True Then Call ShowAuthorXml
        If CreatSite = True Then Call ShowFsClass
        If CreatAnnounce = True Then Call ShowAnnounce
    End If
    
    If ChannelID = 0 Then
        XMLDOM.save (Server.MapPath(InstallDir & "xml/xml.xml"))
    Else
        XMLDOM.save (Server.MapPath(InstallDir & "xml/xml_Channel_" & ChannelID & ".xml"))
    End If
    
    Set Node = Nothing
    Set SubNode = Nothing
    Set XMLDOM = Nothing
    
    If ChannelID = 0 Then
        Response.Write "生成页面（<a href='" & InstallDir & "xml/xml.xml'>" & InstallDir & "xml/xml.xml</a>）<font color=red>成功!</font></b>"
    Else
        Response.Write "生成页面（<a href='" & InstallDir & "xml/xml_Channel_" & ChannelID & ".xml'>" & InstallDir & "xml/xml_Channel_" & ChannelID & ".xml</a>）<font color=red>成功!</font></b>" & vbCrLf
    End If
ElseIf Action = "GreatNav" Then
    If ChannelID = 0 Then
        Response.Write "<Li>频道ID数据输入错误！</Li>" & vbCrLf
    Else
        Dim rsChannel
        Set rsChannel = Conn.Execute("select top 1 ChannelID,OrderID,ChannelName,ChannelType,LinkUrl,ChannelDir,ModuleType,Disabled,UseCreateHTML,ListFileType,FileExt_List from PE_Channel where ChannelID=" & ChannelID & " and Disabled = " & PE_False & " and ModuleType<6 order by OrderID")
        If rsChannel.BOF And rsChannel.EOF Then
            Response.Write "<Li>找不到指定的频道，或频道已被禁用！</Li>"
            rsChannel.Close
            Set rsChannel = Nothing
            Response.End
        Else
            ChannelDir = rsChannel("ChannelDir")
            UseCreateHTML = rsChannel("UseCreateHTML")
            ListFileType = rsChannel("ListFileType")
            FileExt_List = rsChannel("FileExt_List")
            strHTML = "<?xml version=""1.0"" encoding=""GB2312""?>" & vbCrLf
            strHTML = strHTML & ("<menu>" & vbCrLf)
            If rsChannel("ModuleType") = 4 Then
                Call ShowGuestNav
            Else
                Call ShowNav(rsChannel("ChannelID"))
            End If
            strHTML = strHTML & ("</menu>" & vbCrLf)
            
            rsChannel.Close
            Set rsChannel = Nothing
        End If
        Call WriteToFile(InstallDir & "xml/nav_" & ChannelID & ".xml", strHTML)
    End If
    Response.Write "<b>生成网站频道" & ChannelID & "XML栏目导航数据输出页面（" & InstallDir & "xml/nav_" & ChannelID & ".xml）成功!</b><br><br>" & vbCrLf
End If
Response.Write "</body></html>"
Call CloseConn

Sub ShowClass(ByVal iChannelID, ByVal iType)
    Dim rsClass, sqlClass, preDepth, i, ClassNode, ClassNodeTemp
    sqlClass = "select ClassID,ClassName,Depth,ParentID,NextID,LinkUrl,Child,Readme,ClassPicUrl,ClassType,ParentDir,ClassDir,OpenType,ItemCount from PE_Class where ChannelID=" & iChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If Not (rsClass.BOF And rsClass.EOF) Then
        Do While Not rsClass.EOF
            
            Set SubNode = Node.appendChild(XMLDOM.createElement("Class"))
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "ClassID", "")
            ClassNodeTemp.text = rsClass("ClassID")
            SubNode.Attributes.setNamedItem (ClassNodeTemp)
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "ClassName", "")
            ClassNodeTemp.text = xml_nohtml(rsClass("ClassName"))
            SubNode.Attributes.setNamedItem (ClassNodeTemp)
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "ClassReadme", "")
            If Not (rsClass("Readme") = "" Or IsNull(rsClass("Readme"))) Then ClassNodeTemp.text = xml_nohtml(rsClass("Readme"))
            SubNode.Attributes.setNamedItem (ClassNodeTemp)
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "ClassPic", "")
            If Not (rsClass("ClassPicUrl") = "" Or IsNull(rsClass("ClassPicUrl"))) Then ClassNodeTemp.text = rsClass("ClassPicUrl")
            SubNode.Attributes.setNamedItem (ClassNodeTemp)
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "ClassDir", "")
            ClassNodeTemp.text = rsClass("ParentDir") & rsClass("ClassDir")
            SubNode.Attributes.setNamedItem (ClassNodeTemp)
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "Depth", "")
            ClassNodeTemp.text = rsClass("Depth")
            SubNode.Attributes.setNamedItem (ClassNodeTemp)
            
            Set ClassNodeTemp = XMLDOM.createNode(2, "ItemCount", "")
            ClassNodeTemp.text = rsClass("ItemCount")
            SubNode.Attributes.setNamedItem (ClassNodeTemp)

            Select Case iType
            Case 1
                Call ShowArticle(rsClass("ClassID"))
            Case 2
                Call ShowSoft(rsClass("ClassID"))
            Case 3
                Call ShowPhoto(rsClass("ClassID"))
            Case 5
                Call ShowProduct(rsClass("ClassID"))
            End Select
            
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub ShowGuestClass()
    Dim rsGuest, sqlGuest, SubNode2
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("Class"))
    Set SubNode2 = XMLDOM.createNode(2, "ClassID", "")
    SubNode2.text = 0
    SubNode.Attributes.setNamedItem (SubNode2)
    Set SubNode2 = XMLDOM.createNode(2, "ClassName", "")
    SubNode2.text = "未分类栏目"
    SubNode.Attributes.setNamedItem (SubNode2)
    Set SubNode2 = XMLDOM.createNode(2, "ClassReadme", "")
    SubNode2.text = ""
    SubNode.Attributes.setNamedItem (SubNode2)
    Set SubNode2 = XMLDOM.createNode(2, "ClassOrder", "")
    SubNode2.text = 0
    SubNode.Attributes.setNamedItem (SubNode2)
            
    Call ShowGuest(0)

    sqlGuest = "select KindID,KindName,Readme,OrderID from PE_GuestKind order by OrderID"
    Set rsGuest = Conn.Execute(sqlGuest)
    If Not (rsGuest.BOF And rsGuest.EOF) Then
        Do While Not rsGuest.EOF
            Set SubNode = Node.appendChild(XMLDOM.createElement("Class"))
            Set SubNode2 = XMLDOM.createNode(2, "ClassID", "")
            SubNode2.text = rsGuest("KindID")
            SubNode.Attributes.setNamedItem (SubNode2)
            Set SubNode2 = XMLDOM.createNode(2, "ClassName", "")
            SubNode2.text = xml_nohtml(rsGuest("KindName"))
            SubNode.Attributes.setNamedItem (SubNode2)
            Set SubNode2 = XMLDOM.createNode(2, "ClassReadme", "")
            If Not (rsGuest("ReadMe") = "" Or IsNull(rsGuest("ReadMe"))) Then SubNode2.text = xml_nohtml(rsGuest("ReadMe"))
            SubNode.Attributes.setNamedItem (SubNode2)
            Set SubNode2 = XMLDOM.createNode(2, "ClassOrder", "")
            SubNode2.text = rsGuest("OrderID")
            SubNode.Attributes.setNamedItem (SubNode2)
    
            Call ShowGuest(rsGuest("KindID"))
            
            rsGuest.MoveNext
        Loop
    End If
    rsGuest.Close
    Set rsGuest = Nothing
End Sub

Sub ShowArticle(ByVal iClassID)
    Dim rsArticle, ItemNode, ItemNodeTemp
    Set rsArticle = Conn.Execute("select top 20 * from PE_Article Where ClassID=" & iClassID & " and Status=3 and Deleted = " & PE_False & " order by ArticleID")
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        Do While Not rsArticle.EOF
            Set ItemNode = SubNode.appendChild(XMLDOM.createElement("Article"))
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "ArticleID", "")
            ItemNodeTemp.text = rsArticle("ArticleID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Title", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("Title"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Author", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("Author"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "CopyFrom", "")
            If Not (rsArticle("CopyFrom") = "" Or IsNull(rsArticle("CopyFrom"))) Then ItemNodeTemp.text = xml_nohtml(rsArticle("CopyFrom"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Inputer", "")
            If Not (rsArticle("Inputer") = "" Or IsNull(rsArticle("Inputer"))) Then ItemNodeTemp.text = rsArticle("Inputer")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "LinkUrl", "")
            If Not (rsArticle("LinkUrl") = "" Or IsNull(rsArticle("LinkUrl"))) Then
                ItemNodeTemp.text = rsArticle("LinkUrl")
            Else
                ItemNodeTemp.text = SiteUrl & "/" & rsChannel("ChannelDir") & "/ShowArticle.asp?ArticleID=" & rsArticle("ArticleID")
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Hits", "")
            If rsArticle("Hits") = "" Or IsNull(rsArticle("Hits")) Then
                ItemNodeTemp.text = 0
            Else
                ItemNodeTemp.text = rsArticle("Hits")
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Time", "")
            ItemNodeTemp.text = rsArticle("UpdateTime")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            '2006字段不存在
            'Set ItemNodeTemp = XMLDOM.createNode(2, "Hot", "")
            'If rsArticle("Hot") = PE_True Then
            '    ItemNodeTemp.Text = 1
            'Else
            '    ItemNodeTemp.Text = 0
            'End If
            'ItemNode.Attributes.setNamedItem (ItemNodeTemp)
 
            Set ItemNodeTemp = XMLDOM.createNode(2, "OnTop", "")
            If rsArticle("OnTop") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Elite", "")
            If rsArticle("Elite") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            Set ItemNodeTemp = XMLDOM.createNode(2, "DefaultPicUrl", "")
            If Not (rsArticle("DefaultPicUrl") = "" Or IsNull(rsArticle("DefaultPicUrl"))) Then ItemNodeTemp.text = rsArticle("DefaultPicUrl")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            If Not (rsArticle("Intro") = "" Or IsNull(rsArticle("Intro"))) Then
                ItemNode.text = xml_nohtml(rsArticle("Intro"))
            Else
                ItemNode.text = Left(xml_nohtml(rsArticle("Content")), 200)
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowSoft(ByVal iClassID)
    Dim rsArticle, ItemNode, ItemNodeTemp
    Set rsArticle = Conn.Execute("select top 20 * from PE_Soft Where ClassID=" & iClassID & " and Status=3 and Deleted = " & PE_False & " order by SoftID")
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        Do While Not rsArticle.EOF

            Set ItemNode = SubNode.appendChild(XMLDOM.createElement("Soft"))
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "SoftID", "")
            ItemNodeTemp.text = rsArticle("SoftID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Title", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("SoftName"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Author", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("Author"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "CopyFrom", "")
            If Not (rsArticle("CopyFrom") = "" Or IsNull(rsArticle("CopyFrom"))) Then ItemNodeTemp.text = xml_nohtml(rsArticle("CopyFrom"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Inputer", "")
            If Not (rsArticle("Inputer") = "" Or IsNull(rsArticle("Inputer"))) Then ItemNodeTemp.text = rsArticle("Inputer")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "LinkUrl", "")
            ItemNodeTemp.text = SiteUrl & "/" & rsChannel("ChannelDir") & "/ShowSoft.asp?SoftID=" & rsArticle("SoftID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Hits", "")
            If rsArticle("Hits") = "" Or IsNull(rsArticle("Hits")) Then
                ItemNodeTemp.text = 0
            Else
                ItemNodeTemp.text = rsArticle("Hits")
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Time", "")
            ItemNodeTemp.text = rsArticle("UpdateTime")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
             
            Set ItemNodeTemp = XMLDOM.createNode(2, "OnTop", "")
            If rsArticle("OnTop") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Elite", "")
            If rsArticle("Elite") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            Set ItemNodeTemp = XMLDOM.createNode(2, "SoftPicUrl", "")
            If Not (rsArticle("SoftPicUrl") = "" Or IsNull(rsArticle("SoftPicUrl"))) Then ItemNodeTemp.text = rsArticle("SoftPicUrl")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            ItemNode.text = xml_nohtml(rsArticle("SoftIntro"))
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowPhoto(ByVal iClassID)
    Dim rsArticle, ItemNode, ItemNodeTemp
    Set rsArticle = Conn.Execute("select top 20 * from PE_Photo Where ClassID=" & iClassID & " and Status=3 and Deleted = " & PE_False & " order by PhotoID")
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        Do While Not rsArticle.EOF
        
            Set ItemNode = SubNode.appendChild(XMLDOM.createElement("Photo"))
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "PhotoID", "")
            ItemNodeTemp.text = rsArticle("PhotoID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Title", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("PhotoName"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Author", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("Author"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "CopyFrom", "")
            If Not (rsArticle("CopyFrom") = "" Or IsNull(rsArticle("CopyFrom"))) Then ItemNodeTemp.text = xml_nohtml(rsArticle("CopyFrom"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Inputer", "")
            If Not (rsArticle("Inputer") = "" Or IsNull(rsArticle("Inputer"))) Then ItemNodeTemp.text = rsArticle("Inputer")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "LinkUrl", "")
            ItemNodeTemp.text = SiteUrl & "/" & rsChannel("ChannelDir") & "/ShowPhoto.asp?PhotoID=" & rsArticle("PhotoID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Hits", "")
            If rsArticle("Hits") = "" Or IsNull(rsArticle("Hits")) Then
                ItemNodeTemp.text = 0
            Else
                ItemNodeTemp.text = rsArticle("Hits")
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Time", "")
            ItemNodeTemp.text = rsArticle("UpdateTime")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
             
            Set ItemNodeTemp = XMLDOM.createNode(2, "OnTop", "")
            If rsArticle("OnTop") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Elite", "")
            If rsArticle("Elite") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            Set ItemNodeTemp = XMLDOM.createNode(2, "PhotoThumb", "")
            If Not (rsArticle("PhotoThumb") = "" Or IsNull(rsArticle("PhotoThumb"))) Then ItemNodeTemp.text = rsArticle("PhotoThumb")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            ItemNode.text = xml_nohtml(rsArticle("PhotoIntro"))
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowProduct(ByVal iClassID)
    Dim rsArticle, ItemNode, ItemNodeTemp
    Set rsArticle = Conn.Execute("select top 20 * from PE_Product Where ClassID=" & iClassID & " and EnableSale = " & PE_True & " and Deleted = " & PE_False & " order by ProductID")
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        Do While Not rsArticle.EOF
        
            Set ItemNode = SubNode.appendChild(XMLDOM.createElement("Product"))
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "ProductID", "")
            ItemNodeTemp.text = rsArticle("ProductID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Title", "")
            ItemNodeTemp.text = xml_nohtml(rsArticle("ProductName"))
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Price", "")
            ItemNodeTemp.text = rsArticle("Price")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Price_Original", "")
            ItemNodeTemp.text = rsArticle("Price_Original")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Unit", "")
            If Not (rsArticle("Unit") = "" Or IsNull(rsArticle("Unit"))) Then ItemNodeTemp.text = rsArticle("Unit")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "LinkUrl", "")
            ItemNodeTemp.text = SiteUrl & "/" & rsChannel("ChannelDir") & "/ShowProduct.asp?ProductID=" & rsArticle("ProductID")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Hits", "")
            If rsArticle("Hits") = "" Or IsNull(rsArticle("Hits")) Then
                ItemNodeTemp.text = 0
            Else
                ItemNodeTemp.text = rsArticle("Hits")
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Time", "")
            ItemNodeTemp.text = rsArticle("UpdateTime")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Hot", "")
            If rsArticle("IsHot") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
             
            Set ItemNodeTemp = XMLDOM.createNode(2, "OnTop", "")
            If rsArticle("OnTop") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)
            
            Set ItemNodeTemp = XMLDOM.createNode(2, "Elite", "")
            If rsArticle("IsElite") = PE_True Then
                ItemNodeTemp.text = 1
            Else
                ItemNodeTemp.text = 0
            End If
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            Set ItemNodeTemp = XMLDOM.createNode(2, "ProductThumb", "")
            If Not (rsArticle("ProductThumb") = "" Or IsNull(rsArticle("ProductThumb"))) Then ItemNodeTemp.text = rsArticle("ProductThumb")
            ItemNode.Attributes.setNamedItem (ItemNodeTemp)

            ItemNode.text = xml_nohtml(rsArticle("ProductIntro"))
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowGuest(ByVal iClassID)
    Dim rsArticle, rsArticle2, GuestNode, GuestNode2, GuestNode3, GuestNodeTemp
    Set rsArticle = Conn.Execute("select top 20 * from PE_GuestBook Where KindID=" & iClassID & " and GuestID=TopicID and GuestIsPassed = " & PE_True & " order by GuestID")
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        Do While Not rsArticle.EOF
            Set GuestNode = SubNode.appendChild(XMLDOM.createElement("Guest"))
            
            Set GuestNodeTemp = XMLDOM.createNode(2, "GuestID", "")
            GuestNodeTemp.text = rsArticle("GuestID")
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            Set GuestNodeTemp = XMLDOM.createNode(2, "GuestName", "")
            GuestNodeTemp.text = xml_nohtml(rsArticle("GuestName"))
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            Set GuestNodeTemp = XMLDOM.createNode(2, "Title", "")
            GuestNodeTemp.text = xml_nohtml(rsArticle("GuestTitle"))
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            Set GuestNodeTemp = XMLDOM.createNode(2, "Time", "")
            GuestNodeTemp.text = rsArticle("GuestDatetime")
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            Set GuestNodeTemp = XMLDOM.createNode(2, "LinkUrl", "")
            GuestNodeTemp.text = SiteUrl & "/" & rsChannel("ChannelDir") & "/Guest_Reply.asp?TopicID=" & rsArticle("TopicID")
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            Set GuestNodeTemp = XMLDOM.createNode(2, "Hits", "")
            GuestNodeTemp.text = PE_CLng(rsArticle("Hits"))
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            Set GuestNodeTemp = XMLDOM.createNode(2, "ReplyNum", "")
            GuestNodeTemp.text = PE_CLng(rsArticle("ReplyNum"))
            GuestNode.Attributes.setNamedItem (GuestNodeTemp)
            
            GuestNode.text = xml_nohtml(rsArticle("GuestReply"))
            
            If Not (IsNull(rsArticle("GuestReplyAdmin")) Or rsArticle("GuestReplyAdmin") = "") Then
                Set GuestNode2 = GuestNode.appendChild(XMLDOM.createElement("GuestAdminReply"))
                Set GuestNodeTemp = XMLDOM.createNode(2, "AdminName", "")
                GuestNodeTemp.text = xml_nohtml(rsArticle("GuestReplyAdmin"))
                GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                Set GuestNodeTemp = XMLDOM.createNode(2, "Time", "")
                If Not (rsArticle("GuestReplyDatetime") = "" Or IsNull(rsArticle("GuestReplyDatetime"))) Then GuestNodeTemp.text = rsArticle("GuestReplyDatetime")
                GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                GuestNode2.text = xml_nohtml(rsArticle("GuestReply"))
            End If
            If rsArticle("ReplyNum") > 0 Then
                Set rsArticle2 = Conn.Execute("select top 20 * from PE_GuestBook Where GuestID<>TopicID and TopicID=" & rsArticle("GuestID") & " and GuestIsPassed = " & PE_True & " order by GuestID")
                Do While Not rsArticle2.EOF
                    Set GuestNode2 = GuestNode.appendChild(XMLDOM.createElement("GuestReply"))
                
                    Set GuestNodeTemp = XMLDOM.createNode(2, "GuestID", "")
                    GuestNodeTemp.text = rsArticle2("GuestID")
                    GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                    Set GuestNodeTemp = XMLDOM.createNode(2, "GuestName", "")
                    GuestNodeTemp.text = xml_nohtml(rsArticle2("GuestName"))
                    GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                    Set GuestNodeTemp = XMLDOM.createNode(2, "Title", "")
                    GuestNodeTemp.text = xml_nohtml(rsArticle2("GuestTitle"))
                    GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                    Set GuestNodeTemp = XMLDOM.createNode(2, "Time", "")
                    GuestNodeTemp.text = rsArticle2("GuestDatetime")
                    GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                    Set GuestNodeTemp = XMLDOM.createNode(2, "Hits", "")
                    GuestNodeTemp.text = PE_CLng(rsArticle2("Hits"))
                    GuestNode2.Attributes.setNamedItem (GuestNodeTemp)
                    GuestNode2.text = xml_nohtml(rsArticle2("GuestReply"))
                
                    If Not (IsNull(rsArticle2("GuestReplyAdmin")) Or rsArticle2("GuestReplyAdmin") = "") Then
                        Set GuestNode3 = GuestNode2.appendChild(XMLDOM.createElement("GuestAdminReply"))
                        Set GuestNodeTemp = XMLDOM.createNode(2, "AdminName", "")
                        GuestNodeTemp.text = xml_nohtml(rsArticle2("GuestReplyAdmin"))
                        GuestNode3.Attributes.setNamedItem (GuestNodeTemp)
                        Set GuestNodeTemp = XMLDOM.createNode(2, "Time", "")
                        If Not (rsArticle2("GuestReplyDatetime") = "" Or IsNull(rsArticle2("GuestReplyDatetime"))) Then GuestNodeTemp.text = rsArticle2("GuestReplyDatetime")
                        GuestNode3.Attributes.setNamedItem (GuestNodeTemp)
                        GuestNode3.text = xml_nohtml(rsArticle2("GuestReply"))
                    End If
                    rsArticle2.MoveNext
                Loop
                rsArticle2.Close
                Set rsArticle2 = Nothing
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    Set GuestNode = Nothing
    Set GuestNode2 = Nothing
    Set GuestNode3 = Nothing
    Set GuestNodeTemp = Nothing
End Sub


Sub ShowUserXml()
    Dim rsUser, UserNode, UserNodeTemp
    
    Set rsUser = Conn.Execute("select top 40 UserID,UserName,PostItems,PassedItems from PE_User where IsLocked = " & PE_False & " order by UserID")
    If rsUser.BOF And rsUser.EOF Then
        Set Node = XMLDOM.createNode(1, "UserList", "")
        XMLDOM.documentElement.appendChild (Node)
    Else
        Set Node = XMLDOM.createNode(1, "UserList", "")
        XMLDOM.documentElement.appendChild (Node)
        Do While Not rsUser.EOF
            Set UserNode = Node.appendChild(XMLDOM.createElement("User"))
        
            Set UserNodeTemp = XMLDOM.createNode(2, "GuestID", "")
            UserNodeTemp.text = rsUser("UserID")
            UserNode.Attributes.setNamedItem (UserNodeTemp)
            Set UserNodeTemp = XMLDOM.createNode(2, "UserName", "")
            UserNodeTemp.text = xml_nohtml(rsUser("UserName"))
            UserNode.Attributes.setNamedItem (UserNodeTemp)
            Set UserNodeTemp = XMLDOM.createNode(2, "PostItems", "")
            If rsUser("PostItems") = "" Or IsNull(rsUser("PostItems")) Then
                UserNodeTemp.text = 0
            Else
                UserNodeTemp.text = rsUser("PostItems")
            End If
            UserNode.Attributes.setNamedItem (UserNodeTemp)
            Set UserNodeTemp = XMLDOM.createNode(2, "PassedItems", "")
            If rsUser("PassedItems") = "" Or IsNull(rsUser("PassedItems")) Then
                UserNodeTemp.text = 0
            Else
                UserNodeTemp.text = rsUser("PassedItems")
            End If
            UserNode.Attributes.setNamedItem (UserNodeTemp)
            rsUser.MoveNext
        Loop
    End If
    rsUser.Close
    Set rsUser = Nothing
    Set UserNode = Nothing
    Set UserNodeTemp = Nothing
End Sub

Sub ShowAuthorXml()
    Dim rsAuthor, AuthorNode, AuthorNodeTemp
    Set rsAuthor = Conn.Execute("select top 40 ID,ChannelID,AuthorName,Photo,Intro from PE_Author where Passed = " & PE_True & " order by ID")
    If rsAuthor.BOF And rsAuthor.EOF Then
        Set Node = XMLDOM.createNode(1, "AuthorList", "")
        XMLDOM.documentElement.appendChild (Node)
    Else
        Set Node = XMLDOM.createNode(1, "AuthorList", "")
        XMLDOM.documentElement.appendChild (Node)
        Do While Not rsAuthor.EOF
            Set AuthorNode = Node.appendChild(XMLDOM.createElement("Author"))
        
            Set AuthorNodeTemp = XMLDOM.createNode(2, "AuthorID", "")
            AuthorNodeTemp.text = rsAuthor("ID")
            AuthorNode.Attributes.setNamedItem (AuthorNodeTemp)
            
            Set AuthorNodeTemp = XMLDOM.createNode(2, "AuthorName", "")
            AuthorNodeTemp.text = xml_nohtml(rsAuthor("AuthorName"))
            AuthorNode.Attributes.setNamedItem (AuthorNodeTemp)
            
            'Set AuthorNodeTemp = XMLDOM.createNode(2, "NickName", "")
            'AuthorNodeTemp.Text = xml_nohtml(rsAuthor("NiceName"))
            'AuthorNode.Attributes.setNamedItem (AuthorNodeTemp)
            
            Set AuthorNodeTemp = XMLDOM.createNode(2, "Photo", "")
            If Not (rsAuthor("Photo") = "" Or IsNull(rsAuthor("Photo"))) Then AuthorNodeTemp.text = rsAuthor("Photo")
            AuthorNode.Attributes.setNamedItem (AuthorNodeTemp)
            
            AuthorNode.text = rsAuthor("Intro")
            rsAuthor.MoveNext
        Loop
    End If
    rsAuthor.Close
    Set rsAuthor = Nothing
    Set AuthorNode = Nothing
    Set AuthorNodeTemp = Nothing
End Sub

Sub ShowFsClass()
    Dim rsSiteClass, rsSite, FsNode, FsNode2, FsNodeTemp
    Set Node = XMLDOM.createNode(1, "FriendSite", "")
    XMLDOM.documentElement.appendChild (Node)
        
    Set FsNode = Node.appendChild(XMLDOM.createElement("FriendSiteClass"))
    
    Set FsNodeTemp = XMLDOM.createNode(2, "ClassID", "")
    FsNodeTemp.text = 0
    FsNode.Attributes.setNamedItem (FsNodeTemp)
    
    Set FsNodeTemp = XMLDOM.createNode(2, "ClassName", "")
    FsNodeTemp.text = "未分类友情链接"
    FsNode.Attributes.setNamedItem (FsNodeTemp)

    Set FsNodeTemp = XMLDOM.createNode(2, "Readme", "")
    FsNodeTemp.text = ""
    FsNode.Attributes.setNamedItem (FsNodeTemp)
    
    Set FsNodeTemp = XMLDOM.createNode(2, "KindType", "")
    FsNodeTemp.text = 0
    FsNode.Attributes.setNamedItem (FsNodeTemp)
    
    Set rsSite = Conn.Execute("select top 20 * from PE_FriendSite Where KindID=0 and Passed = " & PE_True & " order by OrderID")
    Do While Not rsSite.EOF
        Set FsNode2 = FsNode.appendChild(XMLDOM.createElement("FriendSite"))
        
        Set FsNodeTemp = XMLDOM.createNode(2, "SiteID", "")
        FsNodeTemp.text = rsSite("ID")
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "SiteName", "")
        FsNodeTemp.text = xml_nohtml(rsSite("SiteName"))
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "SiteUrl", "")
        FsNodeTemp.text = rsSite("SiteUrl")
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "LogoUrl", "")
        If Not (rsSite("LogoUrl") = "" Or IsNull(rsSite("LogoUrl"))) Then FsNodeTemp.text = rsSite("LogoUrl")
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "SiteAdmin", "")
        If Not (rsSite("SiteAdmin") = "" Or IsNull(rsSite("SiteAdmin"))) Then FsNodeTemp.text = xml_nohtml(rsSite("SiteAdmin"))
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "SiteEmail", "")
        If Not (rsSite("SiteEmail") = "" Or IsNull(rsSite("SiteEmail"))) Then FsNodeTemp.text = rsSite("SiteEmail")
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "Hits", "")
        FsNodeTemp.text = PE_CLng(rsSite("Hits"))
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        Set FsNodeTemp = XMLDOM.createNode(2, "Time", "")
        If Not (rsSite("UpdateTime") = "" Or IsNull(rsSite("UpdateTime"))) Then FsNodeTemp.text = rsSite("UpdateTime")
        FsNode2.Attributes.setNamedItem (FsNodeTemp)
        
        FsNode2.text = xml_nohtml(rsSite("SiteIntro"))
        
        rsSite.MoveNext
    Loop
    
    Set rsSiteClass = Conn.Execute("select KindID,KindName,Readme,KindType from PE_FsKind order by KindID")
    If Not (rsSiteClass.BOF And rsSiteClass.EOF) Then
        Do While Not rsSiteClass.EOF
            Set FsNode = Node.appendChild(XMLDOM.createElement("FriendSiteClass"))
            
            Set FsNodeTemp = XMLDOM.createNode(2, "ClassID", "")
            FsNodeTemp.text = rsSiteClass("KindID")
            FsNode.Attributes.setNamedItem (FsNodeTemp)
            
            Set FsNodeTemp = XMLDOM.createNode(2, "ClassName", "")
            FsNodeTemp.text = xml_nohtml(rsSiteClass("KindName"))
            FsNode.Attributes.setNamedItem (FsNodeTemp)
            
            Set FsNodeTemp = XMLDOM.createNode(2, "Readme", "")
            If Not (rsSiteClass("Readme") = "" Or IsNull(rsSiteClass("Readme"))) Then FsNodeTemp.text = xml_nohtml(rsSiteClass("Readme"))
            FsNode.Attributes.setNamedItem (FsNodeTemp)
            
            Set FsNodeTemp = XMLDOM.createNode(2, "KindType", "")
            FsNodeTemp.text = rsSiteClass("KindType")
            FsNode.Attributes.setNamedItem (FsNodeTemp)
           
            Set rsSite = Conn.Execute("select top 20 * from PE_FriendSite Where KindID=" & rsSiteClass("KindID") & " and Passed = " & PE_True & " order by OrderID")
            Do While Not rsSite.EOF
                Set FsNode2 = FsNode.appendChild(XMLDOM.createElement("FriendSite"))
                
                Set FsNodeTemp = XMLDOM.createNode(2, "SiteID", "")
                FsNodeTemp.text = rsSite("ID")
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "SiteName", "")
                FsNodeTemp.text = xml_nohtml(rsSite("SiteName"))
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "SiteUrl", "")
                FsNodeTemp.text = rsSite("SiteUrl")
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "LogoUrl", "")
                If Not (rsSite("LogoUrl") = "" Or IsNull(rsSite("LogoUrl"))) Then FsNodeTemp.text = rsSite("LogoUrl")
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "SiteAdmin", "")
                FsNodeTemp.text = xml_nohtml(rsSite("SiteAdmin"))
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "SiteEmail", "")
                If Not (rsSite("SiteEmail") = "" Or IsNull(rsSite("SiteEmail"))) Then FsNodeTemp.text = rsSite("SiteEmail")
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "Hits", "")
                FsNodeTemp.text = PE_CLng(rsSite("Hits"))
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                Set FsNodeTemp = XMLDOM.createNode(2, "Time", "")
                FsNodeTemp.text = rsSite("UpdateTime")
                FsNode2.Attributes.setNamedItem (FsNodeTemp)
                
                FsNode2.text = xml_nohtml(rsSite("SiteIntro"))
                
                rsSite.MoveNext
            Loop

            rsSiteClass.MoveNext
        Loop
    End If
    rsSiteClass.Close
    rsSite.Close
    Set rsSiteClass = Nothing
    Set rsSite = Nothing
    Set FsNode = Nothing
    Set FsNodeTemp = Nothing
End Sub


Sub ShowAnnounce()
    Dim rsAnnounce, AnnounceNode, AnnounceNodeTemp
    
    Set Node = XMLDOM.createNode(1, "AnnounceList", "")
    XMLDOM.documentElement.appendChild (Node)
    
    Set rsAnnounce = Conn.Execute("select * from PE_Announce Where IsSelected = " & PE_True & " order by ID")
    Do While Not rsAnnounce.EOF
        Set AnnounceNode = Node.appendChild(XMLDOM.createElement("Announce"))
        
        Set AnnounceNodeTemp = XMLDOM.createNode(2, "ID", "")
        AnnounceNodeTemp.text = rsAnnounce("ID")
        AnnounceNode.Attributes.setNamedItem (AnnounceNodeTemp)
        
        Set AnnounceNodeTemp = XMLDOM.createNode(2, "Title", "")
        AnnounceNodeTemp.text = xml_nohtml(rsAnnounce("Title"))
        AnnounceNode.Attributes.setNamedItem (AnnounceNodeTemp)
        
        Set AnnounceNodeTemp = XMLDOM.createNode(2, "Author", "")
        AnnounceNodeTemp.text = xml_nohtml(rsAnnounce("Author"))
        AnnounceNode.Attributes.setNamedItem (AnnounceNodeTemp)
        
        Set AnnounceNodeTemp = XMLDOM.createNode(2, "Time", "")
        AnnounceNodeTemp.text = rsAnnounce("DateAndTime")
        AnnounceNode.Attributes.setNamedItem (AnnounceNodeTemp)
        
        Set AnnounceNodeTemp = XMLDOM.createNode(2, "ChannelID", "")
        AnnounceNodeTemp.text = rsAnnounce("ChannelID")
        AnnounceNode.Attributes.setNamedItem (AnnounceNodeTemp)
        
        Set AnnounceNodeTemp = XMLDOM.createNode(2, "ShowType", "")
        AnnounceNodeTemp.text = rsAnnounce("ShowType")
        AnnounceNode.Attributes.setNamedItem (AnnounceNodeTemp)
        
        AnnounceNode.text = xml_nohtml(rsAnnounce("Content"))
        rsAnnounce.MoveNext
    Loop
    rsAnnounce.Close
    Set rsAnnounce = Nothing
    Set AnnounceNode = Nothing
    Set AnnounceNodeTemp = Nothing
End Sub


Sub ShowNav(ByVal iChannelID)
    Dim rsClass, sqlClass, preDepth, i, UrlTemp
    sqlClass = "select ClassID,ClassName,Depth,ParentID,NextID,LinkUrl,Child,ClassType,ParentDir,ClassDir,OpenType,ClassPurview from PE_Class where ChannelID=" & iChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If Not (rsClass.BOF And rsClass.EOF) Then
            preDepth = 0
        Do While Not rsClass.EOF
            If rsClass("ClassPurview") < 2 And (UseCreateHTML = 1 Or UseCreateHTML = 3) Then
                Select Case ListFileType
                Case 0
                    If FileExt_List = 0 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & rsClass("ParentDir") & rsClass("ClassDir") & "/index.html"
                    ElseIf FileExt_List = 1 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & rsClass("ParentDir") & rsClass("ClassDir") & "/index.htm"
                    ElseIf FileExt_List = 2 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & rsClass("ParentDir") & rsClass("ClassDir") & "/index.shtml"
                    ElseIf FileExt_List = 3 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & rsClass("ParentDir") & rsClass("ClassDir") & "/index.shtm"
                    ElseIf FileExt_List = 4 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & rsClass("ParentDir") & rsClass("ClassDir") & "/index.asp"
                    Else
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/ShowClass.asp?ClassID=" & rsClass("ClassID")
                    End If
                Case 1
                    If FileExt_List = 0 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List/List_" & rsClass("ClassID") & ".html"
                    ElseIf FileExt_List = 1 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List/List_" & rsClass("ClassID") & ".htm"
                    ElseIf FileExt_List = 2 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List/List_" & rsClass("ClassID") & ".shtml"
                    ElseIf FileExt_List = 3 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List/List_" & rsClass("ClassID") & ".shtm"
                    ElseIf FileExt_List = 4 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List/List_" & rsClass("ClassID") & ".asp"
                    Else
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/ShowClass.asp?ClassID=" & rsClass("ClassID")
                    End If
                Case 2
                    If FileExt_List = 0 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List_" & rsClass("ClassID") & ".html"
                    ElseIf FileExt_List = 1 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List_" & rsClass("ClassID") & ".htm"
                    ElseIf FileExt_List = 2 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List_" & rsClass("ClassID") & ".shtml"
                    ElseIf FileExt_List = 3 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List_" & rsClass("ClassID") & ".shtm"
                    ElseIf FileExt_List = 4 Then
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/List_" & rsClass("ClassID") & ".asp"
                    Else
                        UrlTemp = SiteUrl & "/" & ChannelDir & "/ShowClass.asp?ClassID=" & rsClass("ClassID")
                    End If
                Case Else
                    UrlTemp = SiteUrl & "/" & ChannelDir & "/ShowClass.asp?ClassID=" & rsClass("ClassID")
                End Select
            Else
                UrlTemp = SiteUrl & "/" & ChannelDir & "/ShowClass.asp?ClassID=" & rsClass("ClassID")
            End If
            If Not IsNull(rsClass("LinkUrl")) Then UrlTemp = rsClass("LinkUrl")
            If preDepth - rsClass("Depth") > 0 Then
                For i = 1 To (preDepth - rsClass("Depth"))
                    strHTML = strHTML & ("</item>" & vbCrLf)
                Next
            End If
            If rsClass("Child") = 0 Then
                strHTML = strHTML & ("<item label=""" & rsClass("ClassName") & """ url=""" & UrlTemp & """ target=""_Self"" />" & vbCrLf)
            Else
                strHTML = strHTML & ("<item label=""" & rsClass("ClassName") & """ url=""" & UrlTemp & """ target=""_Self"">" & vbCrLf)
            End If
            preDepth = rsClass("Depth")
            rsClass.MoveNext
        Loop
        If preDepth > 0 Then
            For i = 1 To preDepth
                strHTML = strHTML & ("</item>" & vbCrLf)
            Next
        End If
    End If
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub ShowGuestNav()
    Dim rsGuest, sqlGuest
    strHTML = strHTML & ("<item label=""留言板首页"" url=""" & SiteUrl & "/GuestBook/index.asp""  target=""_Self"" />" & vbCrLf)
    sqlGuest = "select KindID,KindName,OrderID from PE_GuestKind order by KindID"
    Set rsGuest = Conn.Execute(sqlGuest)
    If Not (rsGuest.BOF And rsGuest.EOF) Then
        Do While Not rsGuest.EOF
            strHTML = strHTML & ("<item label=""" & ReplaceBadChar(rsGuest("KindName")) & """ url=""" & SiteUrl & "/GuestBook/index.asp?KindID=" & rsGuest("KindID") & """  target=""_Self"" />" & vbCrLf)
            rsGuest.MoveNext
        Loop
    End If
    rsGuest.Close
    Set rsGuest = Nothing
End Sub

Function xml_nohtml(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        xml_nohtml = ""
        Exit Function
    End If
    Dim str
    str = Replace(fString, "&gt;", ">")
    str = Replace(str, "&lt;", "<")
    str = Replace(str, "&nbsp;", "")
    str = Replace(str, "&quot;", "")
    str = Replace(str, "&#39;", "")
    regEx.Pattern = "(\<.[^\<]*\>)"
    str = regEx.Replace(str, "")
    regEx.Pattern = "(\<\/[^\<]*\>)"
    str = regEx.Replace(str, "")
    
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    str = Replace(Replace(str, "<![CDATA[", ""), "]]>", "")
    xml_nohtml = str
End Function
%>
