<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Expires = -1
Response.ContentType = "text/xml; charset=gb2312"

'类私有变量
Private ModelShowType, BlogID, BlogDir, ClassID, TypeID, Hot, Elite, AuthorName, OutNum
Private ChannelUrl, ChannelName, ChannelDir, UseCreateHTML, StructureType, FileNameType, FileExt_Item, sqlChannel, rsChannel
Private SubNode
Private UBlogID, UBlogName, UBlogIntro, UBlogBirthDay, UBlogPhoto, UBlogHits, BlogAddress, UBlogTel, UBlogFax, UBlogCompany
Private UBlogAddress, UBlogDepartment, UBlogZipCode, UBlogHomePage, UBlogEmail, UBlogQQ, UBlogLastUseTime, UBlogShowList


Dim strtmp, SiteLogoUrl

If Right(SiteUrl, 1) <> "/" Then SiteUrl = SiteUrl & "/"
SiteLogoUrl = SiteUrl & LogoUrl

UserID = PE_CLng(Trim(Request("ID")))
ClassID = PE_CLng(Trim(Request("ClassID")))
TypeID = PE_CLng(Trim(Request("TypeID")))
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

XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

'输出RSS数据
Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")

strtmp = "<?xml version=""1.0"" encoding=""gb2312""?>"

XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""gb2312"""))

XMLDOM.appendChild (XMLDOM.createElement("body"))
XMLDOM.documentElement.Attributes.setNamedItem(XMLDOM.createNode(2, "version", "")).text = "PowerEasy Cms 2006"

Set Node = XMLDOM.createNode(1, "Site", "")
XMLDOM.documentElement.appendChild (Node)
    
Set SubNode = Node.appendChild(XMLDOM.createElement("SiteName"))
SubNode.text = SiteName
Set SubNode = Node.appendChild(XMLDOM.createElement("SiteTitle"))
SubNode.text = SiteTitle
Set SubNode = Node.appendChild(XMLDOM.createElement("SiteUrl"))
SubNode.text = SiteUrl
Set SubNode = Node.appendChild(XMLDOM.createElement("SiteLogo"))
SubNode.text = SiteLogoUrl
Set SubNode = Node.appendChild(XMLDOM.createElement("BannerUrl"))
SubNode.text = BannerUrl
Set SubNode = Node.appendChild(XMLDOM.createElement("Meta_Description"))
SubNode.text = Meta_Description
Set SubNode = Node.appendChild(XMLDOM.createElement("Meta_Keywords"))
SubNode.text = Meta_Keywords
Set SubNode = Node.appendChild(XMLDOM.createElement("Currentlink"))
SubNode.text = SiteUrl & "Blog"
Set SubNode = Node.appendChild(XMLDOM.createElement("language"))
SubNode.text = "zh-cn"
Set SubNode = Node.appendChild(XMLDOM.createElement("WebmasterName"))
SubNode.text = WebmasterName
Set SubNode = Node.appendChild(XMLDOM.createElement("WebmasterEmail"))
SubNode.text = WebmasterEmail
Set SubNode = Node.appendChild(XMLDOM.createElement("Copyright"))
SubNode.text = Copyright
Set SubNode = Node.appendChild(XMLDOM.createElement("EnableRss"))
If EnableRss = True Then SubNode.text = "enable"
Set SubNode = Node.appendChild(XMLDOM.createElement("EnableWap"))
If EnableWap = True Then SubNode.text = "enable"
Set SubNode = Node.appendChild(XMLDOM.createElement("ShowSiteChannel"))
If ShowSiteChannel = True Then SubNode.text = "enable"
Set SubNode = Node.appendChild(XMLDOM.createElement("ShowAdminLogin"))
If ShowAdminLogin = True Then SubNode.text = "enable"
Set SubNode = Node.appendChild(XMLDOM.createElement("AdminDir"))
If ShowAdminLogin = True Then SubNode.text = AdminDir

Dim xmlconfig, bootnode
Set xmlconfig = Server.CreateObject("Microsoft.XMLDOM")
xmlconfig.async = False
xmlconfig.Load (Server.MapPath("config.xml"))
Set bootnode = xmlconfig.getElementsByTagName("baseconfig")

Dim UqRs






Public Sub GetVisitorList(ibid)
    Dim rsBlog, TempNode
    Set Node = XMLDOM.createNode(1, "NewVisitor", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    Set rsBlog = Conn.Execute("select top 10 UserID,UserName,Datetime,num from PE_SpaceVisitor Where BlogID=" & ibid & " order by Datetime Desc")
    Do While Not rsBlog.EOF
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("visitor"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("userid"))
        Node.text = rsBlog("UserID")
        Set Node = SubNode.appendChild(XMLDOM.createElement("username"))
        Node.text = Replace(Replace(Replace(Replace(LCase(rsBlog("UserName")), "cdx", ""), "cer", ""), "asp", ""), "asa", "")
        Set Node = SubNode.appendChild(XMLDOM.createElement("time"))
        Node.text = rsBlog("Datetime")
        Set Node = SubNode.appendChild(XMLDOM.createElement("num"))
        Node.text = rsBlog("num")
        rsBlog.MoveNext
    Loop
    Set rsBlog = Nothing
End Sub

Public Sub GetChannelList()
    Dim rsBlog, TempNode
    Set Node = XMLDOM.createNode(1, "ChannelList", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    Set rsBlog = Conn.Execute("select ChannelName,LinkUrl,ChannelDir,ReadMe from PE_Channel where Disabled=" & PE_False & " and ShowName=" & PE_True & " order by OrderID")
    Do While Not rsBlog.EOF
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("Channelitem"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
        Node.text = rsBlog("ChannelName")
        Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
        If Trim(rsBlog("LinkUrl") & "") = "" Then
            Node.text = InstallDir & rsBlog("ChannelDir")
        Else
            Node.text = rsBlog("LinkUrl")
        End If
        Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
        If Trim(rsBlog("ReadMe") & "") <> "" Then Node.text = rsBlog("ReadMe")
        rsBlog.MoveNext
    Loop
    Set rsBlog = Nothing
End Sub

Public Sub GetAnnounceList()
    Dim rsBlog, TempNode
    Set Node = XMLDOM.createNode(1, "AnnounceList", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    Set rsBlog = Conn.Execute("select ID,Title,Content,DateAndTime from PE_Announce where IsSelected=" & PE_True & " and ChannelID=-1 and (ShowType=0 or ShowType=1) and (OutTime=0 or OutTime>DateDiff(" & PE_DatePart_D & ",DateAndTime, " & PE_Now & ")) order by ID Desc")
    Do While Not rsBlog.EOF
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("Announceitem"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
        Node.text = rsBlog("Title")
        Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
        Node.text = "Announce.asp?ID=" & rsBlog("ID")
        Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
        If Trim(rsBlog("Content") & "") <> "" Then Node.text = rsBlog("Content")
        Set Node = SubNode.appendChild(XMLDOM.createElement("DateAndTime"))
        Node.text = FormatDateTime(rsBlog("DateAndTime"), 1)
        rsBlog.MoveNext
    Loop
    Set rsBlog = Nothing
End Sub

Public Sub GetBlogClassList()
    Dim rsBlog, TempNode
    Set Node = XMLDOM.createNode(1, "BlogClassList", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    Set rsBlog = Conn.Execute("select KindName,KindId,ReadMe from PE_SpaceKind order by OrderID")
    Do While Not rsBlog.EOF
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("item"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
        Node.text = rsBlog("KindName")
        Set Node = SubNode.appendChild(XMLDOM.createElement("id"))
        Node.text = rsBlog("KindID")
        Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
        If Trim(rsBlog("ReadMe") & "") <> "" Then Node.text = rsBlog("ReadMe")
        rsBlog.MoveNext
    Loop
    Set rsBlog = Nothing
End Sub

Public Sub GetBlogItem(iNodeName, iSQL)
    Dim rsItem, TempNode, spacename
    Set Node = XMLDOM.createNode(1, iNodeName, "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    Set rsItem = Server.CreateObject("ADODB.Recordset")
    rsItem.Open iSQL, Conn, 1, 1
    Do While Not rsItem.EOF
        spacename = Replace(LCase(rsItem("UserName")), ".", "")
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("Blogitem"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
        Node.text = rsItem("Name")
        Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
        Node.text = spacename
        Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
        Node.text = InstallDir & "Space/" & spacename & rsItem("UserID") & "/"
        Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
        If Trim(rsItem("Intro") & "") <> "" Then Node.text = rsItem("Intro")
        Set Node = SubNode.appendChild(XMLDOM.createElement("BirthDay"))
        Node.text = rsItem("BirthDay")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Photo"))
        If Trim(rsItem("Photo") & "") = "" Then
            Node.text = InstallDir & "Space/default.gif"
        Else
            Node.text = rsItem("Photo")
        End If
        Set Node = SubNode.appendChild(XMLDOM.createElement("Top"))
        If rsItem("onTop") = True Then
            Node.text = 1
        Else
            Node.text = 0
        End If
        Set Node = SubNode.appendChild(XMLDOM.createElement("Elite"))
        If rsItem("IsElite") = True Then
            Node.text = 1
        Else
            Node.text = 0
        End If
        Set Node = SubNode.appendChild(XMLDOM.createElement("Hits"))
        Node.text = rsItem("Hits")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Address"))
        If Trim(rsItem("Address") & "") <> "" Then Node.text = rsItem("Address")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Tel"))
        If Trim(rsItem("Tel") & "") <> "" Then Node.text = rsItem("Tel")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Fax"))
        If Trim(rsItem("Fax") & "") <> "" Then Node.text = rsItem("Fax")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Company"))
        If Trim(rsItem("Company") & "") <> "" Then Node.text = rsItem("Company")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Department"))
        If Trim(rsItem("Department") & "") <> "" Then Node.text = rsItem("Department")
        Set Node = SubNode.appendChild(XMLDOM.createElement("ZipCode"))
        If Trim(rsItem("ZipCode") & "") <> "" Then Node.text = rsItem("ZipCode")
        Set Node = SubNode.appendChild(XMLDOM.createElement("HomePage"))
        If Trim(rsItem("HomePage") & "") <> "" Then Node.text = rsItem("HomePage")
        Set Node = SubNode.appendChild(XMLDOM.createElement("Email"))
        If Trim(rsItem("Email") & "") <> "" Then Node.text = rsItem("Email")
        Set Node = SubNode.appendChild(XMLDOM.createElement("QQ"))
        If Trim(rsItem("QQ") & "") <> "" Then Node.text = rsItem("QQ")
        Set Node = SubNode.appendChild(XMLDOM.createElement("LastUseTime"))
        Node.text = rsItem("LastUseTime")
        rsItem.MoveNext
    Loop
    Set rsItem = Nothing
End Sub


Public Sub ShowDiaryList() '输出日志列表
    Dim xmlconfig, i, bootnode, nodeCount, SubNode, BlogSql, PE_Hits, nodeLis

    '输出公告列表
    Call GetAnnounceList

    '输出频道列表
    Call GetChannelList

    Call CloseConn
End Sub

Public Sub ShowDiary(iModelType) '处理日志内容
    Dim sqlDiary, rsDiary, Node, SubNode, TempNode, Bid, datarange, totalpage, iCount, iType
    If Action = "savepl" Then
        Dim PlDom, SubNode2
        Set PlDom = CreateObject("Microsoft.XMLDOM")
        PlDom.async = False
        PlDom.Load Request
        Set Node = PlDom.getElementsByTagName("root")
        If Node.length < 1 Then
            Set SubNode = XMLDOM.createNode(1, "serverbackinfo", "")
            XMLDOM.documentElement.appendChild (SubNode)
            Set SubNode2 = SubNode.appendChild(XMLDOM.createElement("stat"))
            SubNode2.text = "err"
            Set SubNode2 = SubNode.appendChild(XMLDOM.createElement("infomation"))
            SubNode2.text = "输入数据错误!"
        Else
            Dim Dusername, Dnoname, Dpass, Dtitle, Dcontent, Dtype, Did, Ds, Dt
            Dnoname = PE_CLng(Node(0).selectSingleNode("noname").text)
            If Dnoname = 1 Then
                Dusername = "匿名用户"
            Else
                If Node(0).selectSingleNode("username").text <> "" Then
                    Dusername = ReplaceBadChar(Node(0).selectSingleNode("username").text)
                Else
                    Dt = True
                    Ds = "用户名不能为空!"
                End If
                If Node(0).selectSingleNode("password").text <> "" Then
                    Dpass = ReplaceBadChar(Node(0).selectSingleNode("password").text)
                Else
                    Dt = True
                    Ds = "密码不能为空!"
                End If
            End If
            If Node(0).selectSingleNode("title").text <> "" Then
                Dtitle = Node(0).selectSingleNode("title").text
            Else
                Dt = True
                Ds = Ds & "标题不能为空!"
            End If
            If Node(0).selectSingleNode("content").text <> "" Then
                Dcontent = Node(0).selectSingleNode("content").text
            Else
                Dt = True
                Ds = Ds & "内容不能为空!"
            End If
            If Node(0).selectSingleNode("type").text <> "" Then
                Dtype = PE_CLng(Node(0).selectSingleNode("type").text)
            Else
                Dtype = 3
            End If
            If Node(0).selectSingleNode("id").text <> "" Then
                Did = PE_CLng(Node(0).selectSingleNode("id").text)
            Else
                Dt = True
                Ds = Ds & "ID不能为空!"
            End If
            Bid = PE_CLng(Node(0).selectSingleNode("blogid").text)

            If Dt = False Then '评论存盘处理
                Dim CheckUser, RsPl, sqlpl
                If Dnoname = 1 Then
                    Set RsPl = Server.CreateObject("adodb.recordset")
                    sqlpl = "select top 1 * from PE_SpaceComment"
                    RsPl.Open sqlpl, Conn, 1, 3
                        RsPl.addnew
                        RsPl("ItemID") = Did
                        RsPl("BlogID") = Bid
                        RsPl("Type") = Dtype
                        RsPl("Uname") = Dusername
                        RsPl("Title") = Dtitle
                        RsPl("Content") = Dcontent
                        RsPl("Datetime") = Now()
                        RsPl.Update
                    RsPl.Close
                    Select Case iModelType
                    Case "diary"
                        Conn.Execute ("update PE_SpaceDiary Set PlNum=PlNum+1 where ID=" & Did)
                    Case "book"
                        Conn.Execute ("update PE_SpaceBook Set PlNum=PlNum+1 where ID=" & Did)
                    Case "music"
                        Conn.Execute ("update PE_SpaceMusic Set PlNum=PlNum+1 where ID=" & Did)
                    Case "photo"
                        Conn.Execute ("update PE_SpacePhoto Set PlNum=PlNum+1 where ID=" & Did)
                    Case "link"
                        Conn.Execute ("update PE_SpaceLink Set PlNum=PlNum+1 where ID=" & Did)
                    End Select
                Else
                    Set CheckUser = Conn.Execute("select Top 1 UserName from PE_User where UserName='" & Dusername & "' and UserPassword='" & MD5(Dpass, 16) & "'")
                    If CheckUser.BOF And CheckUser.EOF Then
                        Dt = True
                        Ds = Ds & "用户名或密码错!"
                    Else
                        Set RsPl = Server.CreateObject("adodb.recordset")
                        sqlpl = "select top 1 * from PE_SpaceComment"
                        RsPl.Open sqlpl, Conn, 1, 3
                            RsPl.addnew
                            RsPl("ItemID") = Did
                            RsPl("BlogID") = Bid
                            RsPl("Type") = Dtype
                            RsPl("Uname") = Dusername
                            RsPl("Title") = Dtitle
                            RsPl("Content") = Dcontent
                            RsPl("Datetime") = Now()
                        RsPl.Update
                        RsPl.Close
                        Select Case iModelType
                        Case "diary"
                            Conn.Execute ("update PE_SpaceDiary Set PlNum=PlNum+1 where ID=" & Did)
                        Case "book"
                            Conn.Execute ("update PE_SpaceBook Set PlNum=PlNum+1 where ID=" & Did)
                        Case "music"
                            Conn.Execute ("update PE_SpaceMusic Set PlNum=PlNum+1 where ID=" & Did)
                        Case "photo"
                            Conn.Execute ("update PE_SpacePhoto Set PlNum=PlNum+1 where ID=" & Did)
                        Case "link"
                            Conn.Execute ("update PE_SpaceLink Set PlNum=PlNum+1 where ID=" & Did)
                        End Select
                    End If
                    Set CheckUser = Nothing
                End If
                Set RsPl = Nothing
            End If
            Set SubNode = XMLDOM.createNode(1, "serverbackinfo", "")
            XMLDOM.documentElement.appendChild (SubNode)
            If Dt = True Then
                Set SubNode2 = SubNode.appendChild(XMLDOM.createElement("stat"))
                SubNode2.text = "err"
                Set SubNode2 = SubNode.appendChild(XMLDOM.createElement("infomation"))
                SubNode2.text = Ds
            Else
                Set SubNode2 = SubNode.appendChild(XMLDOM.createElement("stat"))
                SubNode2.text = "ok"
                Set SubNode2 = SubNode.appendChild(XMLDOM.createElement("infomation"))
                SubNode2.text = "评论保存成功"
            End If
        End If
        Set Node = Nothing
        Set PlDom = Nothing
    Else
        strField = Trim(Request("Field"))
        Keyword = Trim(Request("keyword"))
        CurrentPage = PE_CLng1(Trim(Request("page")))
        datarange = Trim(Request("data"))

        Select Case iModelType
        Case "diary"
            iType = 3
            sqlDiary = "select A.ID,A.BlogID,A.Title,A.Content,A.Datetime,A.Hits,A.PlNum,C.Name,C.Intro,C.BirthDay,C.Hits,C.LastUseTime,C.listnum from PE_SpaceDiary A inner join PE_Space C on A.BlogID=C.ID"
        Case "music"
            iType = 4
            sqlDiary = "select A.ID,A.BlogID,A.Title,A.Content,A.Datetime,A.Hits,A.PlNum,C.Name,C.Intro,C.BirthDay,C.Hits,C.LastUseTime,C.listnum from PE_SpaceMusic A inner join PE_Space C on A.BlogID=C.ID"
        Case "book"
            iType = 5
            sqlDiary = "select A.ID,A.BlogID,A.Title,A.Content,A.Datetime,A.Hits,A.PlNum,C.Name,C.Intro,C.BirthDay,C.Hits,C.LastUseTime,C.listnum from PE_SpaceBook A inner join PE_Space C on A.BlogID=C.ID"
        Case "photo"
            iType = 6
            sqlDiary = "select A.ID,A.BlogID,A.Title,A.Content,A.Datetime,A.Hits,A.PlNum,C.Name,C.Intro,C.BirthDay,C.Hits,C.LastUseTime,C.listnum from PE_SpacePhoto A inner join PE_Space C on A.BlogID=C.ID"
        Case "link"
            iType = 7
            sqlDiary = "select A.ID,A.BlogID,A.Title,A.Content,A.Datetime,A.Hits,A.PlNum,C.Name,C.Intro,C.BirthDay,C.Hits,C.LastUseTime,C.listnum from PE_SpaceLink A inner join PE_Space C on A.BlogID=C.ID"
        End Select
        If BlogID = 0 Then
            sqlDiary = sqlDiary & " Where A.ID=" & UserID
        Else
            sqlDiary = sqlDiary & " Where A.BlogID=" & BlogID
        End If
        If datarange <> "" Then
            If Not IsDate(datarange) Then
                datarange = Date
            End If
            sqlDiary = sqlDiary & " and A.Datetime=" & datarange
        End If
        sqlDiary = sqlDiary & " order by A.ID desc"
        Set rsDiary = Server.CreateObject("adodb.recordset")
        rsDiary.Open sqlDiary, Conn, 1, 3
        If PE_CLng(rsDiary("listnum")) < 1 Then
            MaxPerPage = 10
        Else
            MaxPerPage = rsDiary("listnum")
        End If
        totalPut = rsDiary.RecordCount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalpage = totalPut \ MaxPerPage
        Else
            totalpage = totalPut \ MaxPerPage + 1
        End If
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
            totalpage = totalpage
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                iMod = 0
                If CurrentPage > MaxPerPage Then
                    iMod = totalPut Mod MaxPerPage
                    If iMod <> 0 Then iMod = MaxPerPage - iMod
                End If
                rsDiary.Move (CurrentPage - 1) * MaxPerPage - iMod
            Else
                CurrentPage = 1
            End If
        End If
        If Not (rsDiary.BOF And rsDiary.EOF) Then
            Bid = rsDiary(1)
            Dim plrs, plnode
            Set Node = XMLDOM.createNode(1, "MyBlog", "")
            Set TempNode = Node
            XMLDOM.documentElement.appendChild (Node)
            Set SubNode = Node.appendChild(XMLDOM.createElement("BlogName"))
            SubNode.text = rsDiary("Name")
            Set SubNode = Node.appendChild(XMLDOM.createElement("BlogID"))
            SubNode.text = rsDiary("BlogID")
            Set SubNode = Node.appendChild(XMLDOM.createElement("BlogDir"))
            SubNode.text = BlogDir
            Set SubNode = Node.appendChild(XMLDOM.createElement("IsRoot"))
            If BlogID = 0 Then
                SubNode.text = 0
            Else
                SubNode.text = 1
            End If
            Set SubNode = Node.appendChild(XMLDOM.createElement("Hits"))
            SubNode.text = rsDiary(10)
            Set SubNode = Node.appendChild(XMLDOM.createElement("BlogIntro"))
            If Trim(rsDiary("Intro") & "") <> "" Then SubNode.text = rsDiary("Intro")
            Set SubNode = Node.appendChild(XMLDOM.createElement("BirthDay"))
            SubNode.text = rsDiary("BirthDay")
            Set SubNode = Node.appendChild(XMLDOM.createElement("LastUseTime"))
            SubNode.text = rsDiary("LastUseTime")
            Set SubNode = Node.appendChild(XMLDOM.createElement("totalPut"))
            SubNode.text = totalPut
            Set SubNode = Node.appendChild(XMLDOM.createElement("TotalPage"))
            SubNode.text = totalpage
            Set SubNode = Node.appendChild(XMLDOM.createElement("CurrentPage"))
            SubNode.text = CurrentPage
            If BlogID = 0 Then
                rsDiary(5) = rsDiary(5) + 1
            Else
                rsDiary(9) = rsDiary(10) + 1
            End If
            rsDiary.Update
            iCount = 0
            Do While Not rsDiary.EOF
                Set Node = TempNode
                Set SubNode = Node.appendChild(XMLDOM.createElement("Diary"))

                Set Node = SubNode.appendChild(XMLDOM.createElement("ID"))
                Node.text = rsDiary("ID")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Title"))
                Node.text = rsDiary("Title")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Content"))
                Node.text = rsDiary("Content")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Datetime"))
                Node.text = rsDiary("Datetime")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Hits"))
                Node.text = rsDiary(5)
                Set Node = SubNode.appendChild(XMLDOM.createElement("Comment"))
                Node.text = rsDiary("PlNum")
                Set plrs = Conn.Execute("select Top 10 Uname,Title,Content,Datetime from PE_SpaceComment Where Type=" & iType & " and ItemID=" & rsDiary("ID") & " order by Datetime desc")
                Do While Not plrs.EOF
                    Set Node = SubNode.appendChild(XMLDOM.createElement("CommentList"))
                    Set plnode = Node.appendChild(XMLDOM.createElement("name"))
                    plnode.text = plrs("Uname")
                    Set plnode = Node.appendChild(XMLDOM.createElement("title"))
                    plnode.text = plrs("Title")
                    Set plnode = Node.appendChild(XMLDOM.createElement("content"))
                    plnode.text = plrs("Content")
                    Set plnode = Node.appendChild(XMLDOM.createElement("datetime"))
                    plnode.text = plrs("Datetime")
                    plrs.MoveNext
                Loop
                rsDiary.MoveNext
                iCount = iCount + 1
                If iCount >= MaxPerPage Then Exit Do
            Loop
            Set plrs = Nothing
        End If
        Set rsDiary = Nothing

    '输出最新评论列表
    If BlogID = 0 And UserID > 0 Then
        Call NewCommentList("U", UserID, iType)
    Else
        Call NewCommentList("B", BlogID, iType)
    End If

    '输出最近访客列表
    Call GetVisitorList(Bid)

    '输出公告列表
    Call GetAnnounceList

    '输出频道列表
    Call GetChannelList
End If
Call CloseConn
End Sub

Public Sub NewCommentList(iList, iID, iType)
    Dim rsBlog, TempNode, tempsql
    Set Node = XMLDOM.createNode(1, "NewCommentList", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    tempsql = "select Top 10 ItemID,Uname,Title,Content,Datetime from PE_SpaceComment Where Type=" & iType
    If iList = "U" Then
        If iID > 0 Then tempsql = tempsql & " and ItemID=" & iID
    Else
        If iID > 0 Then tempsql = tempsql & " and BlogID=" & iID
    End If
    tempsql = tempsql & " order by Datetime desc"
    Set rsBlog = Conn.Execute(tempsql)
    Do While Not rsBlog.EOF
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("Commentitem"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("name"))
        Node.text = rsBlog("Uname")
        Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
        Node.text = rsBlog("Title")
        Set Node = SubNode.appendChild(XMLDOM.createElement("content"))
        Node.text = rsBlog("Content")
        Set Node = SubNode.appendChild(XMLDOM.createElement("datetime"))
        Node.text = rsBlog("Datetime")
        rsBlog.MoveNext
    Loop
    Set rsBlog = Nothing
End Sub

Private Sub addfang()
    Dim FangDom, FangNode, FangRs, FangSql, iuid, ibid
    Set FangDom = CreateObject("Microsoft.XMLDOM")
    FangDom.async = False
    FangDom.Load Request
    Set FangNode = FangDom.getElementsByTagName("root")
    If FangNode.length > 0 Then
        ibid = PE_CLng(FangNode(0).selectSingleNode("blogid").text)
        iuid = PE_CLng(FangNode(0).selectSingleNode("userid").text)
        If iuid > 0 And FangNode(0).selectSingleNode("username").text <> "" Then
            Set FangRs = Server.CreateObject("adodb.recordset")
            FangSql = "select top 1 BlogID,UserID,UserName,Datetime,num from PE_SpaceVisitor Where BlogID=" & ibid & " and UserID=" & iuid
            FangRs.Open FangSql, Conn, 1, 3
            If FangRs.BOF And FangRs.EOF Then
                FangRs.addnew
                FangRs("BlogID") = ibid
                FangRs("UserID") = iuid
                FangRs("UserName") = FangNode(0).selectSingleNode("username").text
                FangRs("Datetime") = Now()
            Else
                FangRs("Datetime") = Now()
                FangRs("num") = FangRs("num") + 1
            End If
            FangRs.Update
            FangRs.Close
            Set FangRs = Nothing
        End If
    End If
    Set FangNode = Nothing
    Set FangDom = Nothing
End Sub
%>
