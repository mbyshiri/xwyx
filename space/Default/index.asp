<!--#include file="../CommonCode.asp"-->
<%
If Action = "addfang" Then
	Call addfang
Else
	If bootnode.length = 0 Then
		UserID = 0
	Else
		On Error Resume Next
		UserID = PE_CLng(bootnode(0).selectSingleNode("userid").Text)
	End If
	Set xmlconfig = Nothing
	If UserID = 0 Then
		strtmp = ""
	Else
		If Action <> "xml" Then strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""index.xsl"" version=""1.0""?>"
		Call ShowUser(UserID, 0)
	End If
End If
strtmp = strtmp & XMLDOM.documentElement.xml

Set Node = Nothing
Set SubNode = Nothing
Set XMLDOM = Nothing

Response.Write strtmp

Call CloseConn

Public Sub ShowUser(iUserID, iClassID)
    If iUserID = 0 Then Exit Sub

    '输出用户BLOG列表
    Call GetUserInfo(iUserID)

    Dim rsBlog, TempNode
    Set Node = XMLDOM.createNode(1, "MyBlog", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    If UBlogID = 0 Then
        Set SubNode = Node.appendChild(XMLDOM.createElement("BlogName"))
        SubNode.Text = "审核中..."
    Else
        Set SubNode = Node.appendChild(XMLDOM.createElement("BlogName"))
        SubNode.Text = UBlogName
        Set SubNode = Node.appendChild(XMLDOM.createElement("BlogID"))
        SubNode.Text = UBlogID
        Set SubNode = Node.appendChild(XMLDOM.createElement("UserName"))
        SubNode.Text = Replace(Replace(Replace(Replace(LCase(UserName), "cdx", ""), "cer", ""), "asp", ""), "asa", "")
        Set SubNode = Node.appendChild(XMLDOM.createElement("BlogIntro"))
        If Trim(UBlogIntro & "") <> "" Then SubNode.Text = UBlogIntro
        Set SubNode = Node.appendChild(XMLDOM.createElement("BirthDay"))
        SubNode.Text = UBlogBirthDay
        Set SubNode = Node.appendChild(XMLDOM.createElement("Photo"))
        If Trim(UBlogPhoto & "") <> "" Then
            SubNode.Text = UBlogPhoto
        Else
            SubNode.Text = InstallDir & "Space/default.gif"
        End If
        Set SubNode = Node.appendChild(XMLDOM.createElement("Hits"))
        SubNode.Text = UBlogHits
        Set SubNode = Node.appendChild(XMLDOM.createElement("Address"))
        If Trim(BlogAddress & "") <> "" Then SubNode.Text = BlogAddress
        Set SubNode = Node.appendChild(XMLDOM.createElement("Tel"))
        If Trim(UBlogTel & "") <> "" Then SubNode.Text = UBlogTel
        Set SubNode = Node.appendChild(XMLDOM.createElement("Fax"))
        If Trim(UBlogFax & "") <> "" Then SubNode.Text = UBlogFax
        Set SubNode = Node.appendChild(XMLDOM.createElement("Company"))
        If Trim(UBlogCompany & "") <> "" Then SubNode.Text = UBlogCompany
        Set SubNode = Node.appendChild(XMLDOM.createElement("Department"))
        If Trim(UBlogDepartment & "") <> "" Then SubNode.Text = UBlogDepartment
        Set SubNode = Node.appendChild(XMLDOM.createElement("ZipCode"))
        If Trim(UBlogZipCode & "") <> "" Then SubNode.Text = UBlogZipCode
        Set SubNode = Node.appendChild(XMLDOM.createElement("HomePage"))
        If Trim(UBlogHomePage & "") <> "" Then SubNode.Text = UBlogHomePage
        Set SubNode = Node.appendChild(XMLDOM.createElement("Email"))
        If Trim(UBlogEmail & "") <> "" Then SubNode.Text = UBlogEmail
        Set SubNode = Node.appendChild(XMLDOM.createElement("QQ"))
        SubNode.Text = UBlogQQ
        Set SubNode = Node.appendChild(XMLDOM.createElement("LastUseTime"))
        SubNode.Text = UBlogLastUseTime

        If iClassID = 0 Then
            Dim rschannellist
            Set rschannellist = Conn.Execute("select ChannelName,ChannelID,ChannelPicUrl,ChannelDir,ModuleType from PE_Channel Where Disabled=" & PE_False & " and ModuleType>0 and ModuleType<4 order by OrderID")
            Do While Not rschannellist.EOF
                If FoundInArr(UBlogShowList, rschannellist("ChannelID"), ",") Then
                    Set Node = TempNode
                    Set SubNode = Node.appendChild(XMLDOM.createElement("Blogitem"))
                    Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
                    Node.Text = rschannellist("ChannelName")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
                    Node.Text = Replace(Replace(Replace(Replace(LCase(UserName), "cdx", ""), "cer", ""), "asp", ""), "asa", "")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("listnum"))
                    Node.Text = 10
                    Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
                    Node.Text = SiteUrl & "rss.asp?ChannelID=" & rschannellist("ChannelID") & "&BlogID=" & UBlogID
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = rschannellist("ModuleType")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("ClassID"))
                    Node.Text = 0
                    Set Node = SubNode.appendChild(XMLDOM.createElement("Photo"))
                    If Trim(rschannellist("ChannelPicUrl") & "") = "" Then
                        Node.Text = InstallDir & "Images/defaultface.gif"
                    Else
                        Node.Text = rschannellist("ChannelPicUrl")
                    End If
                    Set Node = TempNode
                End If
                rschannellist.MoveNext
            Loop
            Set rschannellist = Nothing

            If TypeID > 0 Then
                Set rsBlog = Conn.Execute("select * from PE_Space Where ClassID=" & TypeID & " and Passed=" & PE_True & " and UserID=" & iUserID & " and Type>1 order by OrderID,ID")
            Else
                Set rsBlog = Conn.Execute("select * from PE_Space Where Passed=" & PE_True & " and UserID=" & iUserID & " and Type>1 order by OrderID,ID")
            End If
        Else
            Set rsBlog = Conn.Execute("select top 1 * from PE_Space Where Passed=" & PE_True & " and UserID=" & iUserID & " and Type>1 and ID=" & iClassID)
        End If
        If Not (rsBlog.BOF And rsBlog.EOF) Then
            Do While Not rsBlog.EOF
                Set Node = TempNode
                Set SubNode = Node.appendChild(XMLDOM.createElement("Blogitem"))
                
                Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
                Node.Text = rsBlog("Name")
                Set Node = SubNode.appendChild(XMLDOM.createElement("author"))
                Node.Text = Replace(Replace(Replace(Replace(LCase(GetUserName(rsBlog("UserID"))), "cdx", ""), "cer", ""), "asp", ""), "asa", "")
                Set Node = SubNode.appendChild(XMLDOM.createElement("listnum"))
                Node.Text = rsBlog("listnum")
                Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
                Select Case rsBlog("Type")
                Case 2 '显示RSS
                    Node.Text = rsBlog("LinkUrl")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = "rss"
                Case 3 '显示日志
                    Node.Text = SiteUrl & "rss.asp?Action=diary&BlogID=" & rsBlog("ID")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = "diary"
                Case 4 '显示音乐
                    Node.Text = SiteUrl & "rss.asp?Action=music&BlogID=" & rsBlog("ID")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = "music"
                Case 5 '显示图书
                    Node.Text = SiteUrl & "rss.asp?Action=book&BlogID=" & rsBlog("ID")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = "book"
                Case 6 '显示我的图片
                    Node.Text = SiteUrl & "rss.asp?Action=photo&BlogID=" & rsBlog("ID")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = "photo"
                Case 7 '显示联合链接
                    Node.Text = SiteUrl & "rss.asp?Action=link&BlogID=" & rsBlog("ID")
                    Set Node = SubNode.appendChild(XMLDOM.createElement("type"))
                    Node.Text = "link"
                End Select
                Set Node = SubNode.appendChild(XMLDOM.createElement("description"))
                If Trim(rsBlog("Intro") & "") <> "" Then Node.Text = rsBlog("Intro")
                Set Node = SubNode.appendChild(XMLDOM.createElement("ClassID"))
                Node.Text = rsBlog("ID")
                Set Node = SubNode.appendChild(XMLDOM.createElement("BirthDay"))
                Node.Text = FormatDateTime(rsBlog("BirthDay"), 1)
                Set Node = SubNode.appendChild(XMLDOM.createElement("Photo"))
                If Trim(rsBlog("Photo") & "") = "" Then
                    Node.Text = InstallDir & "Space/default.gif"
                Else
                    Node.Text = rsBlog("Photo")
                End If
                Set Node = SubNode.appendChild(XMLDOM.createElement("Top"))
                If rsBlog("onTop") = True Then
                    Node.Text = 1
                Else
                    Node.Text = 0
                End If
                Set Node = SubNode.appendChild(XMLDOM.createElement("Elite"))
                If rsBlog("IsElite") = True Then
                    Node.Text = 1
                Else
                    Node.Text = 0
                End If
                Set Node = SubNode.appendChild(XMLDOM.createElement("Hits"))
                Node.Text = rsBlog("Hits")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Address"))
                If Trim(rsBlog("Address") & "") <> "" Then Node.Text = rsBlog("Address")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Tel"))
                If Trim(rsBlog("Tel") & "") <> "" Then Node.Text = rsBlog("Tel")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Fax"))
                If Trim(rsBlog("Fax") & "") <> "" Then Node.Text = rsBlog("Fax")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Company"))
                If Trim(rsBlog("Company") & "") <> "" Then Node.Text = rsBlog("Company")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Department"))
                If Trim(rsBlog("Department") & "") <> "" Then Node.Text = rsBlog("Department")
                Set Node = SubNode.appendChild(XMLDOM.createElement("ZipCode"))
                If Trim(rsBlog("ZipCode") & "") <> "" Then Node.Text = rsBlog("ZipCode")
                Set Node = SubNode.appendChild(XMLDOM.createElement("HomePage"))
                If Trim(rsBlog("HomePage") & "") <> "" Then Node.Text = rsBlog("HomePage")
                Set Node = SubNode.appendChild(XMLDOM.createElement("Email"))
                If Trim(rsBlog("Email") & "") <> "" Then Node.Text = rsBlog("Email")
                Set Node = SubNode.appendChild(XMLDOM.createElement("QQ"))
                If Trim(rsBlog("QQ") & "") <> "" Then Node.Text = rsBlog("QQ")
                Set Node = SubNode.appendChild(XMLDOM.createElement("LastUseTime"))
                Node.Text = rsBlog("LastUseTime")
                rsBlog.MoveNext
            Loop
        End If
        Set rsBlog = Nothing
    End If

    '输出我的连接
    Call GetMyLink(iUserID)

    '输出最近访客列表
    Call GetVisitorList(UBlogID)

    '输出公告列表
    Call GetAnnounceList

    '输出频道列表
    Call GetChannelList

    Call CloseConn
End Sub

Public Sub GetMyLink(iuid)
    Dim rsBlog, TempNode
    Set Node = XMLDOM.createNode(1, "MyLink", "")
    Set TempNode = Node
    XMLDOM.documentElement.appendChild (Node)
    Set rsBlog = Conn.Execute("select Title,Content from PE_SpaceLink Where UserID=" & iuid)
    Do While Not rsBlog.EOF
        Set Node = TempNode
        Set SubNode = Node.appendChild(XMLDOM.createElement("linkitem"))
        Set Node = SubNode.appendChild(XMLDOM.createElement("title"))
        Node.Text = rsBlog("Title")
        Set Node = SubNode.appendChild(XMLDOM.createElement("link"))
        Node.Text = rsBlog("Content")
        rsBlog.MoveNext
    Loop
    Set rsBlog = Nothing
End Sub

Private Sub GetUserInfo(iUserID)
    Dim rsUser
    If IsNull(iUserID) Then Exit Sub
    Set rsUser = Conn.Execute("select top 1 A.ID,A.Name,A.Intro,A.BirthDay,A.Photo,A.Hits,A.Address,A.Tel,A.Fax,A.Company,A.Department,A.ZipCode,A.HomePage,A.Email,A.QQ,A.LastUseTime,A.LinkUrl,C.UserName from PE_Space A inner join PE_User C on A.UserID=C.UserID Where A.Passed=" & PE_True & " and A.Type=1 and A.UserID=" & iUserID)
    If Not (rsUser.BOF And rsUser.EOF) Then
        UBlogID = rsUser("ID")
        UBlogName = rsUser("Name")
        UBlogIntro = rsUser("Intro")
        UBlogBirthDay = rsUser("BirthDay")
        UBlogPhoto = rsUser("Photo")
        UBlogHits = rsUser("Hits")
        UBlogAddress = rsUser("Address")
        UBlogTel = rsUser("Tel")
        UBlogFax = rsUser("Fax")
        UBlogCompany = rsUser("Company")
        UBlogDepartment = rsUser("Department")
        UBlogZipCode = rsUser("ZipCode")
        UBlogHomePage = rsUser("HomePage")
        UBlogEmail = rsUser("Email")
        UBlogQQ = rsUser("QQ")
        UBlogLastUseTime = rsUser("LastUseTime")
        UBlogShowList = rsUser("LinkUrl")
        UserName = rsUser("UserName")
    Else
        UserName = "该用户不存在"
    End If
    Set rsUser = Nothing
End Sub

Private Function GetUserName(iUserID)
    Dim rsUser
    If IsNull(iUserID) Then Exit Function
    Set rsUser = Conn.Execute("select top 1 UserName from PE_User Where UserID=" & iUserID)
    If Not (rsUser.BOF And rsUser.EOF) Then
        GetUserName = rsUser("UserName")
    End If
    Set rsUser = Nothing
End Function

%>