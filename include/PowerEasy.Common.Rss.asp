<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub ShowIndexRss(iShowType)
    Dim rsChannel2, rsItem, sqlItem
    Dim ModuleType, tempNode
    
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
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    If iShowType = 0 Then
        SubNode.text = SiteUrl & "rss.asp"
    Else
        SubNode.text = SiteUrl & "xml/Rss.xml"
    End If
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = WebmasterName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    
    Set rsChannel2 = Conn.Execute("select ChannelID,ChannelName,ModuleType,ChannelDir,UseCreateHTML,StructureType,FileNameType,FileExt_Item,LinkUrl from PE_Channel where ModuleType<6 and ModuleType<>4 and Disabled = " & PE_False & " and ChannelType<2 order by OrderID")
    If Not (rsChannel2.BOF And rsChannel2.EOF) Then
        Do While Not rsChannel2.EOF
            ModuleType = rsChannel2("ModuleType")
            ChannelName = rsChannel2("ChannelName")
            ChannelDir = rsChannel2("ChannelDir")
            UseCreateHTML = rsChannel2("UseCreateHTML")
            StructureType = rsChannel2("StructureType")
            FileNameType = rsChannel2("FileNameType")
            FileExt_Item = arrFileExt(rsChannel2("FileExt_Item"))

            '只使用绝对地址时，才使用频道子域名
            If IsNull(rsChannel2("LinkUrl")) Or Trim(rsChannel2("LinkUrl")) = "" Or Left(strInstallDir, 7) <> "http://" Then
                ChannelUrl = SiteUrl & ChannelDir
            Else
                ChannelUrl = rsChannel2("LinkUrl")
            End If

            If Right(ChannelUrl, 1) = "/" Then
                ChannelUrl = Left(ChannelUrl, Len(ChannelUrl) - 1)
            End If

            If SystemDatabaseType = "SQL" Then
                ChannelUrl_ASPFile = ChannelUrl
            Else
                ChannelUrl_ASPFile = SiteUrl & ChannelDir
            End If

            OutNum = 0
            Select Case ModuleType
            Case 1
                sqlItem = "select top 100 ArticleID,ChannelID,ClassID,Title,Author,Hits,UpdateTime,Elite,Content,InfoPurview,InfoPoint,Status,Deleted from PE_Article Where ChannelID=" & rsChannel2("ChannelID")
                If BlogID > 0 Then sqlItem = sqlItem & " and BlogID=" & BlogID
                sqlItem = sqlItem & " and Status=3 and Deleted=" & PE_False & " order by UpdateTime Desc"
            Case 2
                sqlItem = "select top 100 SoftID,ChannelID,ClassID,SoftName,Author,Hits,UpdateTime,Elite,SoftIntro,InfoPurview,InfoPoint,Status,Deleted from PE_Soft Where ChannelID=" & rsChannel2("ChannelID")
                If BlogID > 0 Then sqlItem = sqlItem & " and BlogID=" & BlogID
                sqlItem = sqlItem & " and Status=3 and Deleted=" & PE_False & " order by UpdateTime Desc"
            Case 3
                sqlItem = "select top 100 PhotoID,ChannelID,ClassID,PhotoName,Author,Hits,UpdateTime,Elite,PhotoIntro,InfoPurview,InfoPoint,Status,Deleted from PE_Photo Where ChannelID=" & rsChannel2("ChannelID")
                If BlogID > 0 Then sqlItem = sqlItem & " and BlogID=" & BlogID
                sqlItem = sqlItem & " and Status=3 and Deleted=" & PE_False & " order by UpdateTime Desc"
            Case 5
                sqlItem = "select top 100 ProductID,ChannelID,ClassID,ProductName,ProducerName,Hits,UpdateTime,IsElite,ProductIntro,MinNumber,Stocks,EnableSale,Deleted from PE_Product Where ChannelID=" & rsChannel2("ChannelID")
                If BlogID > 0 Then sqlItem = sqlItem & " and BlogID=" & BlogID
                sqlItem = sqlItem & " and Deleted=" & PE_False & " and EnableSale=" & PE_True & " and Stocks>0 order by UpdateTime Desc"
            End Select
            Set rsItem = Server.CreateObject("adodb.recordset")
            rsItem.Open sqlItem, Conn, 1, 1
            Do While Not rsItem.EOF
                If GetClassFild(rsItem(2), 2) < 2 Or ModuleType = 5 Then
                    Set Node = tempNode
                    Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
                
                    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
                    Node.text = ChannelName & " - " & ReplaceText(xml_nohtml(rsItem(3)), 2)
    
                    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
                    If ModuleType = 5 Then
                        Node.text = GetProductUrl(GetClassFild(rsItem(2), 4), GetClassFild(rsItem(2), 3), rsItem(6), rsItem(0))
                    Else
                        Select Case ModuleType
                        Case 1
                            Node.text = GetArticleUrl(GetClassFild(rsItem(2), 4), GetClassFild(rsItem(2), 3), rsItem(6), rsItem(0), GetClassFild(rsItem(2), 2), rsItem(9), rsItem(10))
                        Case 2
                            Node.text = GetSoftUrl(GetClassFild(rsItem(2), 4), GetClassFild(rsItem(2), 3), rsItem(6), rsItem(0))
                        Case 3
                            Node.text = GetPhotoUrl(GetClassFild(rsItem(2), 4), GetClassFild(rsItem(2), 3), rsItem(6), rsItem(0), GetClassFild(rsItem(2), 2), rsItem(9), rsItem(10))
                        End Select
                    End If

                    Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
                    If ModuleType = 5 Or (rsItem(9) = 0 And rsItem(10) = 0 And GetClassFild(rsItem(2), 2) = 0) Then
                        Node.text = ReplaceText(Left(xml_nohtml(rsItem(8)), 200), 1)
                    Else
                        Node.text = strNoSee
                    End If
                    Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
                    If IsNull(rsItem(4)) Or rsItem(4) = "" Then
                        Node.text = strDefAuthor
                    Else
                        Node.text = xml_nohtml(rsItem(4))
                    End If
                    Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
                    Node.text = GetClassFild(rsItem(2), 1)
                    Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
                    Node.text = rsItem(6)
                    If OutNum > 19 Then
                        Exit Do
                    Else
                        OutNum = OutNum + 1
                    End If
                End If
                rsItem.MoveNext
            Loop

            rsItem.Close
            rsChannel2.MoveNext
        Loop
    End If
    Set rsItem = Nothing
    Set rsChannel2 = Nothing
End Sub

Sub ShowArtcileRss(ByVal iHot, ByVal iElite, ByVal iAuthorName, iHitsOfHot)
    Dim sqlArticle, rsArticle, tempNode, tempUrl
    Call GetChannel(ChannelID)

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

    If SpecialID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 专题更新")
    ElseIf ClassID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 栏目更新")
    Else
        SubNode.text = SiteName & " -- " & ChannelName
    End If
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    tempUrl = SiteUrl & "rss.asp?ChannelID=" & ChannelID
    If ClassID > 0 Then tempUrl = tempUrl & "&ClassID=" & ClassID
    If SpecialID > 0 Then tempUrl = tempUrl & "&SpecialID=" & SpecialID
    If BlogID > 0 Then tempUrl = tempUrl & "&BlogID=" & BlogID
    If iAuthorName <> "none" Then tempUrl = tempUrl & "&AuthorName=" & iAuthorName
    If iHot = 1 Then
        tempUrl = tempUrl & "&Hot=" & iHot
    ElseIf iElite = 1 Then
        tempUrl = tempUrl & "&Elite=" & iElite
    End If
    SubNode.text = tempUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    If SpecialID > 0 Then
        sqlArticle = "select top 100 A.ArticleID,A.ChannelID,A.ClassID,A.BlogID,A.Title,A.Author,A.Hits,A.UpdateTime,A.Elite,A.Content,A.InfoPurview,A.InfoPoint,A.Status,A.Deleted,A.Receive,I.SpecialID from PE_Article A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.ArticleID=I.ItemID Where I.SpecialID=" & SpecialID & " and A.ChannelID=" & ChannelID
        If ClassID > 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & ClassID
        If BlogID > 0 Then sqlArticle = sqlArticle & " and A.BlogID=" & BlogID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and A.Author='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and A.Status=3 and A.Deleted=" & PE_False
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and A.Hits>" & iHitsOfHot & " order by A.Hits " & PE_OrderType & ",A.UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and A.Elite=" & PE_True & " order by A.UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
        End If
    Else
        sqlArticle = "select top 100 ArticleID,ChannelID,ClassID,BlogID,Title,Author,Hits,UpdateTime,Elite,Content,InfoPurview,InfoPoint,Status,Deleted,Receive from PE_Article Where ChannelID=" & ChannelID
        If ClassID <> 0 Then sqlArticle = sqlArticle & " and ClassID=" & ClassID
        If BlogID > 0 Then sqlArticle = sqlArticle & " and BlogID=" & BlogID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and Author='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and Status=3 and Deleted=" & PE_False
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and Hits>" & iHitsOfHot & " order by Hits " & PE_OrderType & ",UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and Elite=" & PE_True & " order by UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by UpdateTime Desc"
        End If
    End If
    Set rsArticle = Conn.Execute(sqlArticle)
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        OutNum = 0
        Do While Not rsArticle.EOF
            If GetClassFild(rsArticle("ClassID"), 2) < 2 Then
                Set Node = tempNode
                Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
                Node.text = ReplaceText(xml_nohtml(rsArticle("Title")), 2)

                Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
                Node.text = GetArticleUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("ArticleID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint"))

                Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
                If rsArticle("InfoPurview") = 0 And rsArticle("InfoPoint") = 0 And GetClassFild(rsArticle("ClassID"), 2) = 0 And rsArticle("Receive") = False Then
                    Node.text = ReplaceText(Left(xml_nohtml(rsArticle("Content")), 200), 1) & "..."
                Else
                    Node.text = strNoSee
                End If
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
                If Trim(rsArticle("Author") & "") = "" Then
                    Node.text = strDefAuthor
                Else
                    Node.text = xml_nohtml(rsArticle("Author"))
                End If
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
                Node.text = GetClassFild(rsArticle("ClassID"), 1)
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
                Node.text = rsArticle("UpdateTime")
                If OutNum > 19 Then
                    Exit Do
                Else
                    OutNum = OutNum + 1
                End If
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowSoftRss(ByVal iHot, ByVal iElite, ByVal iAuthorName, iHitsOfHot)
    Dim sqlArticle, rsArticle, tempNode, tempUrl
    
    If IsNull(ChannelID) Or ChannelID = 0 Then
        Exit Sub
    End If

    Call GetChannel(ChannelID)

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
    If ClassID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 栏目更新")
    ElseIf SpecialID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 专题更新")
    Else
        SubNode.text = SiteName & " -- " & ChannelName
    End If
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    tempUrl = SiteUrl & "rss.asp?ChannelID=" & ChannelID
    If ClassID <> 0 Then tempUrl = tempUrl & "&ClassID=" & ClassID
    If SpecialID <> 0 Then tempUrl = tempUrl & "&SpecialID=" & SpecialID
    If BlogID > 0 Then tempUrl = tempUrl & "&BlogID=" & BlogID
    If iAuthorName <> "none" Then tempUrl = tempUrl & "&AuthorName=" & iAuthorName
    If iHot = 1 Then
        tempUrl = tempUrl & "&Hot=" & iHot
    ElseIf iElite = 1 Then
        tempUrl = tempUrl & "&Elite=" & iElite
    End If
    SubNode.text = tempUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    If SpecialID > 0 Then
        sqlArticle = "select top 100 A.SoftID,A.ChannelID,A.ClassID,A.BlogID,A.SoftName,A.SoftVersion,A.Author,A.Hits,A.UpdateTime,A.Elite,A.SoftIntro,A.InfoPoint,A.Status,A.Deleted,I.SpecialID from PE_Soft A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.SoftID=I.ItemID Where I.SpecialID=" & SpecialID & " and A.ChannelID=" & ChannelID
        If ClassID <> 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & ClassID
        If BlogID > 0 Then sqlArticle = sqlArticle & " and A.BlogID=" & BlogID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and A.Author='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and A.Status=3 and A.Deleted=" & PE_False
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and A.Hits>" & iHitsOfHot & " order by A.Hits " & PE_OrderType & ",A.UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and A.Elite=" & PE_True & " order by A.UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
        End If
    Else
        sqlArticle = "select top 100 SoftID,ChannelID,ClassID,BlogID,SoftName,SoftVersion,Author,Hits,UpdateTime,Elite,SoftIntro,InfoPoint,Status,Deleted from PE_Soft Where ChannelID=" & ChannelID
        If ClassID > 0 Then sqlArticle = sqlArticle & " and ClassID=" & ClassID
        If BlogID > 0 Then sqlArticle = sqlArticle & " and BlogID=" & BlogID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and Author='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and Status=3 and Deleted=" & PE_False
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and Hits>" & iHitsOfHot & " order by Hits " & PE_OrderType & ",UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and Elite=" & PE_True & " order by UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by UpdateTime Desc"
        End If
    End If
    Set rsArticle = Conn.Execute(sqlArticle)
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        OutNum = 0
        Do While Not rsArticle.EOF
            If GetClassFild(rsArticle("ClassID"), 2) < 2 Then
                Set Node = tempNode
                Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
                Node.text = xml_nohtml(rsArticle("SoftName") & rsArticle("SoftVersion"))

                Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
                Node.text = GetSoftUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("SoftID"))
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
                Node.text = Left(xml_nohtml(rsArticle("SoftIntro")), 200)
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
                If Trim(rsArticle("Author") & "") = "" Then
                    Node.text = strDefAuthor
                Else
                    Node.text = xml_nohtml(rsArticle("Author"))
                End If
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
                Node.text = GetClassFild(rsArticle("ClassID"), 1)
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
                Node.text = rsArticle("UpdateTime")
                If OutNum > 19 Then
                    Exit Do
                Else
                    OutNum = OutNum + 1
                End If
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowPhotoRss(ByVal iHot, ByVal iElite, ByVal iAuthorName, iHitsOfHot)
    Dim sqlArticle, rsArticle, tempNode, tempUrl
    
    If IsNull(ChannelID) Or ChannelID = 0 Then
        Exit Sub
    End If

    Call GetChannel(ChannelID)

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
    If ClassID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 栏目更新")
    ElseIf SpecialID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 专题更新")
    Else
        SubNode.text = SiteName & " -- " & ChannelName
    End If
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    tempUrl = SiteUrl & "rss.asp?ChannelID=" & ChannelID
    If ClassID <> 0 Then tempUrl = tempUrl & "&ClassID=" & ClassID
    If SpecialID <> 0 Then tempUrl = tempUrl & "&SpecialID=" & SpecialID
    If BlogID > 0 Then tempUrl = tempUrl & "&BlogID=" & BlogID
    If iAuthorName <> "none" Then tempUrl = tempUrl & "&AuthorName=" & iAuthorName
    If iHot = 1 Then
        tempUrl = tempUrl & "&Hot=" & iHot
    ElseIf iElite = 1 Then
        tempUrl = tempUrl & "&Elite=" & iElite
    End If
    SubNode.text = tempUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    If SpecialID > 0 Then
        sqlArticle = "select top 100 A.PhotoID,A.ChannelID,A.ClassID,A.BlogID,A.PhotoName,A.Author,A.Hits,A.UpdateTime,A.Elite,A.PhotoIntro,A.InfoPurview,A.InfoPoint,A.Status,A.Deleted,I.SpecialID from PE_Photo A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.PhotoID=I.ItemID Where I.SpecialID=" & SpecialID & " and A.ChannelID=" & ChannelID
        If ClassID > 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & ClassID
        If BlogID > 0 Then sqlArticle = sqlArticle & " and A.BlogID=" & BlogID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and A.Author='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and A.Status=3 and A.Deleted=" & PE_False
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and A.Hits>" & iHitsOfHot & " order by A.Hits " & PE_OrderType & ",A.UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and A.Elite=" & PE_True & " order by A.UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
        End If
    Else
        sqlArticle = "select top 100 PhotoID,ChannelID,ClassID,BlogID,PhotoName,Author,Hits,UpdateTime,Elite,PhotoIntro,InfoPurview,InfoPoint,Status,Deleted from PE_Photo Where ChannelID=" & ChannelID
        If ClassID > 0 Then sqlArticle = sqlArticle & " and ClassID=" & ClassID
        If BlogID > 0 Then sqlArticle = sqlArticle & " and BlogID=" & BlogID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and Author='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and Status=3 and Deleted=" & PE_False
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and Hits>" & iHitsOfHot & " order by Hits " & PE_OrderType & ",UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and Elite=" & PE_True & " order by UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by UpdateTime Desc"
        End If
    End If

    Set rsArticle = Conn.Execute(sqlArticle)
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        OutNum = 0
        Do While Not rsArticle.EOF
            If GetClassFild(rsArticle("ClassID"), 2) < 2 Then
                Set Node = tempNode
                Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
                Node.text = xml_nohtml(rsArticle("PhotoName"))

                Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
                Node.text = GetPhotoUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("PhotoID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint"))
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
                Node.text = Left(xml_nohtml(rsArticle("PhotoIntro")), 200)
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
                If Trim(rsArticle("Author") & "") = "" Then
                    Node.text = strDefAuthor
                Else
                    Node.text = xml_nohtml(rsArticle("Author"))
                End If
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
                Node.text = GetClassFild(rsArticle("ClassID"), 1)
            
                Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
                Node.text = rsArticle("UpdateTime")
                If OutNum > 19 Then
                    Exit Do
                Else
                    OutNum = OutNum + 1
                End If
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowProductRss(ByVal iHot, ByVal iElite, ByVal iAuthorName)
    Dim sqlArticle, rsArticle, tempNode, tempUrl
    
    If IsNull(ChannelID) Or ChannelID = 0 Then
        Exit Sub
    End If

    Call GetChannel(ChannelID)

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
    If ClassID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 栏目更新")
    ElseIf SpecialID > 0 Then
        SubNode.text = SiteName & " -- " & ChannelName & XmlText("Rss", "Gx1", " -- 专题更新")
    Else
        SubNode.text = SiteName & " -- " & ChannelName
    End If
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    tempUrl = SiteUrl & "rss.asp?ChannelID=" & ChannelID
    If ClassID <> 0 Then tempUrl = tempUrl & "&ClassID=" & ClassID
    If SpecialID <> 0 Then tempUrl = tempUrl & "&SpecialID=" & SpecialID
    If iAuthorName <> "none" Then tempUrl = tempUrl & "&AuthorName=" & iAuthorName
    If iHot = 1 Then
        tempUrl = tempUrl & "&Hot=" & iHot
    ElseIf iElite = 1 Then
        tempUrl = tempUrl & "&Elite=" & iElite
    End If
    SubNode.text = tempUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    If SpecialID <> 0 Then
        sqlArticle = "select top 100 A.ProductID,A.ChannelID,A.ClassID,A.ProductName,A.ProducerName,A.Hits,A.UpdateTime,A.IsHot,A.IsElite,A.ProductIntro,A.Stocks,A.EnableSale,A.Deleted,I.SpecialID from PE_Product A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.ProductID=I.ItemID Where I.SpecialID=" & SpecialID & " and A.ChannelID=" & ChannelID
        If ClassID <> 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & ClassID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and A.ProducerName='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and A.Deleted=" & PE_False & " and A.EnableSale=" & PE_True & " and A.Stocks>0"
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and A.IsHot=" & PE_True & "order by A.UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and A.IsElite=" & PE_True & "order by A.UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
        End If
    Else
        sqlArticle = "select top 100 ProductID,ChannelID,ClassID,ProductName,ProducerName,Hits,UpdateTime,IsHot,IsElite,ProductIntro,Stocks,EnableSale,Deleted from PE_Product Where ChannelID=" & ChannelID
        If ClassID <> 0 Then sqlArticle = sqlArticle & " and ClassID=" & ClassID
        If iAuthorName <> "none" Then sqlArticle = sqlArticle & " and ProducerName='" & iAuthorName & "'"
        sqlArticle = sqlArticle & " and Deleted=" & PE_False & " and EnableSale=" & PE_True & " and Stocks>0"
        If iHot = 1 Then
            sqlArticle = sqlArticle & " and IsHot=" & PE_True & " order by UpdateTime Desc"
        ElseIf iElite = 1 Then
            sqlArticle = sqlArticle & " and IsElite=" & PE_True & " order by UpdateTime Desc"
        Else
            sqlArticle = sqlArticle & " order by UpdateTime Desc"
        End If
    End If
    Set rsArticle = Conn.Execute(sqlArticle)
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        OutNum = 0
        Do While Not rsArticle.EOF
            Set Node = tempNode
            Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
            Node.text = xml_nohtml(rsArticle("ProductName"))

            Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
            Node.text = GetProductUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("ProductID"))
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
            Node.text = Left(xml_nohtml(rsArticle("ProductIntro")), 200)
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
            If Trim(rsArticle("ProducerName") & "") = "" Then
                Node.text = strDefAuthor
            Else
                Node.text = xml_nohtml(rsArticle("ProducerName"))
            End If
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
            Node.text = GetClassFild(rsArticle("ClassID"), 1)
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
            Node.text = rsArticle("UpdateTime")
            If OutNum > 19 Then
                Exit Do
            Else
                OutNum = OutNum + 1
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub ShowGuestRss()
    Dim rsArticle, tempNode, rsKind
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
    SubNode.text = SiteName & " -- " & ChannelName
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    SubNode.text = SiteUrl & "rss.asp?ChannelID=4"
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    Set rsArticle = Conn.Execute("select top 20 GuestID,KindID,TopicID,GuestTitle,GuestName,GuestContent,GuestDatetime,GuestIsPassed from PE_GuestBook Where GuestIsPassed=" & PE_True & " and GuestIsPrivate=" & PE_False & " order by GuestDatetime Desc")
    If Not (rsArticle.BOF And rsArticle.EOF) Then
        OutNum = 0
        Do While Not rsArticle.EOF
            Set Node = tempNode
            Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
            Node.text = xml_nohtml(rsArticle("GuestTitle"))

            Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
            Node.text = SiteUrl & "GuestBook/Guest_Reply.asp?TopicID=" & rsArticle("TopicID")
                        
            Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
            Node.text = Left(xml_nohtml(rsArticle("GuestContent")), 200)
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
            Node.text = xml_nohtml(rsArticle("GuestName"))
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
            If rsArticle("KindID") = 0 Then
                Node.text = "未分类"
            Else
                Set rsKind = Conn.Execute("select top 1 KindID,KindName from PE_GuestKind Where KindID=" & rsArticle("KindID"))
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Node.text = rsKind("KindName")
                Else
                    Node.text = "未分类"
                End If
                rsKind.Close
            End If
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
            Node.text = rsArticle("GuestDatetime")
            If OutNum > 19 Then
                Exit Do
            Else
                OutNum = OutNum + 1
            End If
            rsArticle.MoveNext
        Loop
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    Set rsKind = Nothing
End Sub


'*************************************************
'开始处理其他RSS输出
'*************************************************
Sub ShowOtherRss(iType)
    Dim rsRss, tempNode, sqlRss, temptxt
    Select Case iType
    Case "diary"
        temptxt = "日志"
    Case "music"
        temptxt = "音乐"
    Case "book"
        temptxt = "图书"
    Case "photo"
        temptxt = "图片"
    Case "link"
        temptxt = "连接"
    End Select
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
    SubNode.text = SiteName & " -- 个人" & temptxt
    Set SubNode = Node.appendChild(XMLDOM.createElement("description"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("link"))
    SubNode.text = SiteUrl
    Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
    Select Case iType
    Case "diary"
        If BlogID > 0 Then
            SubNode.text = SiteUrl & "rss.asp?Action=diary&BlogID=" & BlogID
        Else
            SubNode.text = SiteUrl & "rss.asp?Action=diary"
        End If
    Case "music"
        If BlogID > 0 Then
            SubNode.text = SiteUrl & "rss.asp?Action=music&BlogID=" & BlogID
        Else
            SubNode.text = SiteUrl & "rss.asp?Action=music"
        End If
    Case "book"
        If BlogID > 0 Then
            SubNode.text = SiteUrl & "rss.asp?Action=book&BlogID=" & BlogID
        Else
            SubNode.text = SiteUrl & "rss.asp?Action=book"
        End If
    Case "photo"
        If BlogID > 0 Then
            SubNode.text = SiteUrl & "rss.asp?Action=photo&BlogID=" & BlogID
        Else
            SubNode.text = SiteUrl & "rss.asp?Action=photo"
        End If
    Case "link"
        If BlogID > 0 Then
            SubNode.text = SiteUrl & "rss.asp?Action=link&BlogID=" & BlogID
        Else
            SubNode.text = SiteUrl & "rss.asp?Action=link"
        End If
    End Select
    Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
    SubNode.text = "zh-cn"
    Set SubNode = Node.appendChild(XMLDOM.createElement("docs"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("generator"))
    SubNode.text = SiteName
    Set SubNode = Node.appendChild(XMLDOM.createElement("webMaster"))
    SubNode.text = WebmasterName
    
    Set tempNode = Node
    
    Set SubNode = Node.appendChild(XMLDOM.createElement("image"))
    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
    Node.text = SiteName
    Set Node = SubNode.appendChild(XMLDOM.createElement("url"))
    Node.text = SiteLogoUrl
    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
    Node.text = SiteUrl
    Select Case iType
    Case "diary"
        sqlRss = "select top 30 A.ID,A.UserID,A.Title,A.Content,A.Datetime,C.UserName from PE_SpaceDiary A inner join PE_User C on A.UserID=C.UserID"
        If BlogID > 0 Then sqlRss = sqlRss & " Where A.BlogID=" & BlogID & " order by A.ID desc"
    Case "music"
        sqlRss = "select top 30 A.ID,A.UserID,A.Title,A.Content,A.Datetime,C.UserName from PE_SpaceMusic A inner join PE_User C on A.UserID=C.UserID"
        If BlogID > 0 Then sqlRss = sqlRss & " Where A.BlogID=" & BlogID & " order by A.ID desc"
    Case "book"
        sqlRss = "select top 30 A.ID,A.UserID,A.Title,A.Content,A.Datetime,C.UserName from PE_SpaceBook A inner join PE_User C on A.UserID=C.UserID"
        If BlogID > 0 Then sqlRss = sqlRss & " Where A.BlogID=" & BlogID & " order by A.ID desc"
    Case "photo"
        sqlRss = "select top 30 A.ID,A.UserID,A.Title,A.Content,A.Datetime,C.UserName from PE_SpacePhoto A inner join PE_User C on A.UserID=C.UserID"
        If BlogID > 0 Then sqlRss = sqlRss & " Where A.BlogID=" & BlogID & " order by A.ID desc"
    Case "link"
        sqlRss = "select top 30 A.ID,A.UserID,A.Title,A.Content,A.Datetime,C.UserName from PE_SpaceLink A inner join PE_User C on A.UserID=C.UserID"
        If BlogID > 0 Then sqlRss = sqlRss & " Where A.BlogID=" & BlogID & " order by A.ID desc"
    End Select
    Set rsRss = Conn.Execute(sqlRss)
    If Not (rsRss.BOF And rsRss.EOF) Then
        Do While Not rsRss.EOF
            Set Node = tempNode
            Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
            Node.text = xml_nohtml(rsRss("Title"))

            Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
            Node.text = SiteUrl & "Space/Show" & iType & ".asp?ID=" & rsRss("ID")
                        
            Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
            Node.text = Left(xml_nohtml(rsRss("Content")), 100)
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
            Node.text = rsRss("UserName")
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("category"))
            Node.text = temptxt
            
            Set Node = SubNode.appendChild(XMLDOM.createElement("pubDate"))
            Node.text = rsRss("Datetime")
            rsRss.MoveNext
        Loop
    End If
    rsRss.Close
    Set rsRss = Nothing
End Sub

'**************************************************
'函数名：ReplaceText
'作  用：过滤非法字符串
'参  数：iText-----输入字符串
'返回值：替换后字符串
'**************************************************
Function ReplaceText(iText, iType)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol
    If PE_Cache.GetValue("Site_ReplaceText") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=1 order by Priority"
        rsKey.Open sqlKey, Conn, 1, 1
        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_ReplaceText", rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceText = iText
            Exit Function
        End If
    End If
    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_ReplaceText"), "@@@")
    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        If Int(Keycol(3)) = 0 Or Int(Keycol(3)) = iType Then rText = PE_Replace(rText, Keycol(0), Keycol(1))
    Next
    ReplaceText = rText
End Function
%>
