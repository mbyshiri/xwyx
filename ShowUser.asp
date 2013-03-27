<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<!--#include file="Include/PowerEasy.Common.Content.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 0
PageTitle = "会员信息"

Dim strTemp, arrTemp, UseWordsList
Dim rsUser, sqlUser
UseWordsList = False

UserID = PE_CLng(Trim(Request("UserID")))
UserName = ReplaceBadChar(Trim(Request("UserName")))


If UserID <= 0 And UserName = "" Then
    Call WriteErrMsg("请指定要查看的会员的ID或用户名", ComeUrl)
    Response.End
End If

Dim ListType
ListType = PE_CLng(Request("ListType"))

sqlUser = "select U.UserID,U.UserName,U.UserType,U.UserFace,U.Sign,U.Privacy,C.Sex, C.ZipCode,C.Fax,C.OfficePhone,C.HomePhone,C.Address,C.Department,C.Company,C.TrueName,C.QQ,C.Email,C.HomePage,C.Birthday from PE_User U left join PE_Contacter C on U.ContacterID=C.ContacterID"
If UserID > 0 Then
    sqlUser = sqlUser & " where U.UserID=" & UserID & ""
Else
    sqlUser = sqlUser & " where U.UserName='" & UserName & "'"
End If
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open sqlUser, Conn, 1, 1
If rsUser.BOF And rsUser.EOF Then
    rsUser.Close
    Set rsUser = Nothing
    Call WriteErrMsg("找不到指定的会员", ComeUrl)
    Response.End
End If

UserID = rsUser("UserID")
strFileName = "ShowUser.asp?UserID=" & UserID & "&ListType=" & ListType

strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

strHtml = GetTemplate(0, 8, 0)
strHtml = Replace(strHtml, "{$UserID}", rsUser("UserID"))
Call ReplaceCommonLabel
strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
strHtml = Replace(strHtml, "{$NickName}", "")
strHtml = Replace(strHtml, "{$ShowList}", "UserList.asp")
If rsUser("Sex") = 2 Then
    strHtml = Replace(strHtml, "{$Sex}", strGirl)
ElseIf rsUser("Sex") = 1 Then
    strHtml = Replace(strHtml, "{$Sex}", strMan)
Else
    strHtml = Replace(strHtml, "{$Sex}", Secrit)
End If
If rsUser("Privacy") > 1 Then
    strHtml = Replace(strHtml, "{$UserFace}", "<img id='UserFace' src='" & InstallDir & "Images/defaultface.gif' width='150' height='172'>")
    strHtml = Replace(strHtml, "{$UserName}", Secrit)
    strHtml = Replace(strHtml, "{$TrueName}", Secrit)
    strHtml = Replace(strHtml, "{$BirthDay}", Secrit)
    strHtml = Replace(strHtml, "{$Company}", Secrit)
    strHtml = Replace(strHtml, "{$Department}", Secrit)
    strHtml = Replace(strHtml, "{$Address}", Secrit)
    strHtml = Replace(strHtml, "{$HomePhone}", Secrit)
    strHtml = Replace(strHtml, "{$OfficePhone}", Secrit)
    strHtml = Replace(strHtml, "{$Fax}", Secrit)
    strHtml = Replace(strHtml, "{$ZipCode}", Secrit)
    strHtml = Replace(strHtml, "{$HomePage}", Secrit)
    strHtml = Replace(strHtml, "{$Email}", Secrit)
    strHtml = Replace(strHtml, "{$QQ}", Secrit)
    strHtml = Replace(strHtml, "{$UserType}", Secrit)
Else
    If Trim(rsUser("UserFace") & "") = "" Then
        strHtml = Replace(strHtml, "{$UserFace}", "<img id='UserFace' src='" & InstallDir & "Images/defaultface.gif' width='150' height='172'>")
    Else
        strHtml = Replace(strHtml, "{$UserFace}", "<img id='UserFace' src='" & rsUser("UserFace") & "' width='150' height='172'>")
    End If
    strHtml = Replace(strHtml, "{$UserName}", rsUser("UserName"))
    strHtml = Replace(strHtml, "{$BirthDay}", ReplaceSpace(rsUser("BirthDay")))
    strHtml = Replace(strHtml, "{$HomePage}", ReplaceSpace(rsUser("HomePage")))
    strHtml = Replace(strHtml, "{$Email}", ReplaceSpace(rsUser("Email")))
    strHtml = Replace(strHtml, "{$QQ}", ReplaceSpace(rsUser("QQ")))

    If rsUser("Privacy") = 1 Then
        strHtml = Replace(strHtml, "{$TrueName}", Secrit)
        strHtml = Replace(strHtml, "{$Company}", Secrit)
        strHtml = Replace(strHtml, "{$Department}", Secrit)
        strHtml = Replace(strHtml, "{$Address}", Secrit)
        strHtml = Replace(strHtml, "{$HomePhone}", Secrit)
        strHtml = Replace(strHtml, "{$OfficePhone}", Secrit)
        strHtml = Replace(strHtml, "{$Fax}", Secrit)
        strHtml = Replace(strHtml, "{$ZipCode}", Secrit)
        strHtml = Replace(strHtml, "{$UserType}", Secrit)
    Else
        strHtml = Replace(strHtml, "{$TrueName}", ReplaceSpace(rsUser("TrueName")))
        strHtml = Replace(strHtml, "{$Company}", ReplaceSpace(rsUser("Company")))
        strHtml = Replace(strHtml, "{$Department}", ReplaceSpace(rsUser("Department")))
        strHtml = Replace(strHtml, "{$Address}", ReplaceSpace(rsUser("Address")))
        strHtml = Replace(strHtml, "{$HomePhone}", ReplaceSpace(rsUser("HomePhone")))
        strHtml = Replace(strHtml, "{$OfficePhone}", ReplaceSpace(rsUser("OfficePhone")))
        strHtml = Replace(strHtml, "{$Fax}", ReplaceSpace(rsUser("Fax")))
        strHtml = Replace(strHtml, "{$ZipCode}", ReplaceSpace(rsUser("ZipCode")))
        If rsUser("UserType") = 1 Then
            strHtml = Replace(strHtml, "{$UserType}", XmlText("ShowSource", "ShowUser/Persen", "个人会员"))
        Else
            strHtml = Replace(strHtml, "{$UserType}", XmlText("ShowSource", "ShowUser/Group", "单位会员"))
        End If
    End If
End If
strHtml = PE_Replace(strHtml, "{$Sign}", ReplaceSpace(rsUser("Sign")))

regEx.Pattern = "\{\$ItemList\((.*?)\)\}"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    arrTemp = Split(Match.SubMatches(0), ",")
    If UBound(arrTemp) <> 1 Then
        strTemp = "函数式标签：{$AuthorList()}的参数个数不对。请检查模板中的此标签。"
    Else
        If arrTemp(0) < 2 Then
            strTemp = GetItemList(rsUser("UserID"), arrTemp(0), arrTemp(1))
        Else
            strTemp = GetItemList(rsUser("UserName"), arrTemp(0), arrTemp(1))
        End If
        If arrTemp(0) > 1 Then UseWordsList = True
    End If
    strHtml = Replace(strHtml, Match.value, strTemp)
Next
If UseWordsList = True Then
    strTemp = ""
    Dim rs
    Set rs = Conn.Execute("select ModuleType from PE_Channel where ModuleType > 0 and ModuleType < 4 and Disabled=" & PE_False & " GROUP BY ModuleType")
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            Select Case rs("ModuleType")
            Case 1
                strTemp = strTemp & "&nbsp;<a class=""workslist"" href=""ShowUser.asp?UserID=" & rsUser("UserID") & "&ListType=1"">查看文章集</a>"
            Case 2
                strTemp = strTemp & "&nbsp;<a class=""workslist"" href=""ShowUser.asp?UserID=" & rsUser("UserID") & "&ListType=2"">查看软件集</a>"
            Case 3
                strTemp = strTemp & "&nbsp;<a class=""workslist"" href=""ShowUser.asp?UserID=" & rsUser("UserID") & "&ListType=3"">查看图片集</a>"
            End Select
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    strHtml = Replace(strHtml, "{$WorksList}", strTemp)
Else
    strHtml = Replace(strHtml, "{$WorksList}", "")
End If
rsUser.Close
Set rsUser = Nothing
Response.Write strHtml
Call CloseConn

Function GetItemList(ByVal iUsername, ByVal TitleLen, ByVal iorder)
    Dim strtmp, i, HotNum
    Dim rsAuthor, sqlAuthor
    Dim Character_Class

    If iUsername = "" Then
        GetItemList = "用户名丢失"
        Exit Function
    Else
        iUsername = ReplaceBadChar(iUsername)		
    End If
    strtmp = "<table class=""user_item_list"" width='100%'>"
    Dim iTitleLen, TitleStr, strLink
    Dim iTop, iElite, iCommon, iHot, iNew, ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4

    Select Case ListType
    Case 0, 1   '显示该用户文集
        iTop = XmlText("Article", "ArticleList/t4", "固顶")
        iElite = XmlText("Article", "ArticleList/t3", "推荐")
        iCommon = XmlText("Article", "ArticleList/t5", "普通")
        iHot = XmlText("Article", "ArticleList/t7", "热点")
        iNew = XmlText("Article", "ArticleList/t6", "最新")
        ArticlePro1 = XmlText("Article", "ArticlePro1", "[图文]")
        ArticlePro2 = XmlText("Article", "ArticlePro2", "[组图]")
        ArticlePro3 = XmlText("Article", "ArticlePro3", "[推荐]")
        ArticlePro4 = XmlText("Article", "ArticlePro4", "[注意]")
        Character_Class = XmlText("Article", "Include/ClassChar", "[{$Text}]")

        sqlAuthor = "select A.ChannelID,A.ArticleID,A.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,A.Title,A.Author,A.Inputer,A.UpdateTime,A.TitleFontColor,A.TitleFontType,A.ShowCommentLink,A.Hits,A.OnTop,A.Elite,A.IncludePic,A.InfoPurview,A.InfoPoint from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where Inputer='" & iUsername & "' and A.Deleted=" & PE_False & " and A.Status=3 and A.ReceiveType=0"
        Select Case iorder
        Case 0
            sqlAuthor = sqlAuthor & " order by A.ArticleID Desc"
        Case 1
            sqlAuthor = sqlAuthor & " order by A.ArticleID"
        Case 2
            sqlAuthor = sqlAuthor & " order by A.Hits Desc,A.ArticleID Desc"
        Case 3
            sqlAuthor = sqlAuthor & " order by A.Hits,A.ArticleID Desc"
        End Select

        Set rsAuthor = Server.CreateObject("ADODB.Recordset")
        rsAuthor.Open sqlAuthor, Conn, 1, 1
        If rsAuthor.BOF And rsAuthor.EOF Then
            totalPut = 0
            strtmp = strtmp & "<tr><td>用户" & iUsername & "尚未发表文章！</td></tr>"
        Else
            totalPut = rsAuthor.RecordCount
            If CurrentPage > 1 Then
                If (CurrentPage - 1) * MaxPerPage < totalPut Then
                    rsAuthor.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
            End If
            i = 0

            strtmp = strtmp & "<tr class='Channel_title'><td align='center'>作品名称</td><td align='center' width='60'>作者</td><td align='center' width='80'>发表时间</td></tr>"
            Do While Not rsAuthor.EOF
                If rsAuthor("ChannelID") <> PrevChannelID Then
                    Call GetChannel(rsAuthor("ChannelID"))
                    PrevChannelID = rsAuthor("ChannelID")
                End If
                If TitleLen > 0 Then
                    If rsAuthor("IncludePic") > 0 Then
                        iTitleLen = TitleLen - 6
                    Else
                        iTitleLen = TitleLen
                    End If
                    TitleStr = GetSubStr(ReplaceText(rsAuthor("Title"), 2), iTitleLen, True)
                Else
                    TitleStr = rsAuthor("Title")
                End If
                Select Case rsAuthor("TitleFontType")
                Case 1
                    TitleStr = "<b>" & TitleStr & "</b>"
                Case 2
                    TitleStr = "<em>" & TitleStr & "</em>"
                Case 3
                    TitleStr = "<b><em>" & TitleStr & "</em></b>"
                End Select
                If rsAuthor("TitleFontColor") <> "" Then
                    TitleStr = "<font color='" & rsAuthor("TitleFontColor") & "'>" & TitleStr & "</font>"
                End If

                If rsAuthor("OnTop") = True Then
                    strLink = "<img src='" & ChannelUrl & "/images/article_ontop.gif' alt='" & iTop & ChannelShortName & "'>"
                ElseIf rsAuthor("Elite") = True Then
                    strLink = "<img src='" & ChannelUrl & "/images/article_elite.gif' alt='" & iElite & ChannelShortName & "'>"
                Else
                    strLink = "<img src='" & ChannelUrl & "/images/article_common.gif' alt='" & iCommon & ChannelShortName & "'>"
                End If

                If rsAuthor("ClassID") <> -1 Then
                    strLink = strLink & Replace(Character_Class, "{$Text}", "<a href='" & GetClassUrl(rsAuthor("ParentDir"), rsAuthor("ClassDir"), rsAuthor("ClassID"), rsAuthor("ClassPurview")) & "'>" & rsAuthor("ClassName") & "</a>")
                End If

                Select Case rsAuthor("IncludePic")
                Case 1
                    strLink = strLink & "<span class='S_headline1'>" & ArticlePro1 & "</span>"
                Case 2
                    strLink = strLink & "<span class='S_headline2'>" & ArticlePro2 & "</span>"
                Case 3
                    strLink = strLink & "<span class='S_headline3'>" & ArticlePro3 & "</span>"
                Case 4
                    strLink = strLink & "<span class='S_headline4'>" & ArticlePro4 & "</span>"
                End Select

                If Left(ChannelUrl, 1) <> "/" Then
                    strLink = strLink & "<a href='"
                Else
                    strLink = strLink & "<a href='http://" & Trim(Request.ServerVariables("HTTP_HOST"))
                End If
                strLink = strLink & GetArticleUrl(rsAuthor("ParentDir"), rsAuthor("ClassDir"), rsAuthor("UpdateTime"), rsAuthor("ArticleID"), rsAuthor("ClassPurview"), rsAuthor("InfoPurview"), rsAuthor("InfoPoint")) & "'>"
                strtmp = strtmp & ("<tr class='article_list_body'><td>" & strLink & TitleStr & "</a></td><td>" & GetAuthorInfo(rsAuthor("Author"), rsAuthor("ChannelID")) & "<td align='center' width='80'>" & FormatDateTime(rsAuthor("UpdateTime"), 2) & "</td></tr>")
                rsAuthor.MoveNext
                i = i + 1
                If i >= MaxPerPage Then Exit Do
            Loop
        End If
        rsAuthor.Close
        Set rsAuthor = Nothing
        strtmp = strtmp & "</table>"
        strtmp = strtmp & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "篇文章", False)
    Case 2   '显示该用户软件集
       iTop = XmlText("Soft", "SoftList/t4", "固顶")
        iElite = XmlText("Soft", "SoftList/t3", "推荐")
        iCommon = XmlText("Soft", "SoftList/t5", "普通")
        iHot = XmlText("Soft", "SoftList/t7", "热点")
        iNew = XmlText("Soft", "SoftList/t6", "最新")
        Character_Class = XmlText("Soft", "Include/ClassChar", "[{$Text}]")
        sqlAuthor = "select S.ChannelID,S.SoftID,S.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,S.SoftName,S.SoftVersion,S.Author,S.Inputer,S.keyword,S.UpdateTime,S.Hits,S.DayHits,S.WeekHits,S.MonthHits,S.OnTop,S.Elite from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where S.Inputer='" & iUsername & "' and S.Deleted=" & PE_False & " and S.Status=3"
        Select Case iorder
        Case 0
            sqlAuthor = sqlAuthor & " order by S.SoftID Desc"
        Case 1
            sqlAuthor = sqlAuthor & " order by S.SoftID"
        Case 2
            sqlAuthor = sqlAuthor & " order by S.Hits Desc,S.SoftID Desc"
        Case 3
            sqlAuthor = sqlAuthor & " order by S.Hits,S.SoftID Desc"
        End Select
    
        Set rsAuthor = Server.CreateObject("ADODB.Recordset")
        rsAuthor.Open sqlAuthor, Conn, 1, 1
        If rsAuthor.BOF And rsAuthor.EOF Then
            totalPut = 0
            strtmp = strtmp & "<tr><td>用户" & iUsername & "尚未发表软件！</td></tr>"
        Else
            totalPut = rsAuthor.RecordCount
            If CurrentPage > 1 Then
                If (CurrentPage - 1) * MaxPerPage < totalPut Then
                    rsAuthor.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
            End If
            i = 0

            strtmp = strtmp & "<tr class='Channel_title'><td align='center'>作品名称</td><td align='center' width='60'>作者</td><td align='center' width='80'>发表时间</td></tr>"
            Do While Not rsAuthor.EOF
                If rsAuthor("ChannelID") <> PrevChannelID Then
                    Call GetChannel(rsAuthor("ChannelID"))
                    PrevChannelID = rsAuthor("ChannelID")
                End If
                If TitleLen > 0 Then
                    TitleStr = GetSubStr(ReplaceText(rsAuthor("SoftName"), 2), TitleLen, True)
                Else
                    TitleStr = rsAuthor("SoftName")
                End If

                If rsAuthor("OnTop") = True Then
                    strLink = "<img src='" & ChannelUrl & "/images/Soft_ontop.gif' alt='" & iTop & ChannelShortName & "'>"
                ElseIf rsAuthor("Elite") = True Then
                    strLink = "<img src='" & ChannelUrl & "/images/Soft_elite.gif' alt='" & iElite & ChannelShortName & "'>"
                Else
                    strLink = "<img src='" & ChannelUrl & "/images/Soft_common.gif' alt='" & iCommon & ChannelShortName & "'>"
                End If

                If rsAuthor("ClassID") <> -1 Then
                    strLink = strLink & Replace(Character_Class, "{$Text}", "<a href='" & GetClassUrl(rsAuthor("ParentDir"), rsAuthor("ClassDir"), rsAuthor("ClassID"), rsAuthor("ClassPurview")) & "'>" & rsAuthor("ClassName") & "</a>")
                End If

                If Left(ChannelUrl, 1) <> "/" Then
                    strLink = strLink & "<a href='"
                Else
                    strLink = strLink & "<a href='http://" & Trim(Request.ServerVariables("HTTP_HOST"))
                End If
                strLink = strLink & GetSoftUrl(rsAuthor("ParentDir"), rsAuthor("ClassDir"), rsAuthor("UpdateTime"), rsAuthor("SoftID")) & "'>"
                strtmp = strtmp & ("<tr class='article_list_body'><td>" & strLink & TitleStr & "</a></td><td>" & GetAuthorInfo(rsAuthor("Author"), rsAuthor("ChannelID")) & "<td align='center' width='80'>" & FormatDateTime(rsAuthor("UpdateTime"), 2) & "</td></tr>")
                rsAuthor.MoveNext
                i = i + 1
                If i >= MaxPerPage Then Exit Do
            Loop
        End If
        rsAuthor.Close
        Set rsAuthor = Nothing
        strtmp = strtmp & "</table>"
        strtmp = strtmp & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个软件", False)
    Case 3   '显示该用户图片集
        iTop = XmlText("Photo", "PhotoList/t4", "固顶")
        iElite = XmlText("Photo", "PhotoList/t3", "推荐")
        iCommon = XmlText("Photo", "PhotoList/t5", "普通")
        iHot = XmlText("Photo", "PhotoList/t7", "热点")
        iNew = XmlText("Photo", "PhotoList/t6", "最新")
        Character_Class = XmlText("Photo", "Include/ClassChar", "[{$Text}]")

        sqlAuthor = "select P.ChannelID,P.PhotoID,P.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,P.PhotoName,P.Author,P.Inputer,P.UpdateTime,P.Hits,P.DayHits,P.WeekHits,P.MonthHits,P.OnTop,P.Elite,P.InfoPurview,P.InfoPoint from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID where P.Inputer='" & iUsername & "' and P.Deleted=" & PE_False & " and P.Status=3"
        Select Case iorder
        Case 0
            sqlAuthor = sqlAuthor & " order by P.PhotoID Desc"
        Case 1
            sqlAuthor = sqlAuthor & " order by P.PhotoID"
        Case 2
            sqlAuthor = sqlAuthor & " order by P.Hits Desc,P.PhotoID Desc"
        Case 3
            sqlAuthor = sqlAuthor & " order by P.Hits,P.PhotoID Desc"
        End Select
    
        Set rsAuthor = Server.CreateObject("ADODB.Recordset")
        rsAuthor.Open sqlAuthor, Conn, 1, 1
        If rsAuthor.BOF And rsAuthor.EOF Then
            totalPut = 0
            strtmp = strtmp & "<tr><td>用户" & iUsername & "尚未发表图片！</td></tr>"
        Else
            totalPut = rsAuthor.RecordCount
            If CurrentPage > 1 Then
                If (CurrentPage - 1) * MaxPerPage < totalPut Then
                    rsAuthor.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
            End If
            i = 0

            strtmp = strtmp & "<tr class='Channel_title'><td align='center'>作品名称</td><td align='center' width='60'>作者</td><td align='center' width='80'>发表时间</td></tr>"
            Do While Not rsAuthor.EOF
                If rsAuthor("ChannelID") <> PrevChannelID Then
                    Call GetChannel(rsAuthor("ChannelID"))
                    PrevChannelID = rsAuthor("ChannelID")
                End If
                If TitleLen > 0 Then
                    TitleStr = GetSubStr(ReplaceText(rsAuthor("PhotoName"), 2), TitleLen, True)
                Else
                    TitleStr = rsAuthor("Title")
                End If

                If rsAuthor("OnTop") = True Then
                    strLink = "<img src='" & ChannelUrl & "/images/Photo_ontop.gif' alt='" & iTop & ChannelShortName & "'>"
                ElseIf rsAuthor("Elite") = True Then
                    strLink = "<img src='" & ChannelUrl & "/images/Photo_elite.gif' alt='" & iElite & ChannelShortName & "'>"
                Else
                    strLink = "<img src='" & ChannelUrl & "/images/Photo_common.gif' alt='" & iCommon & ChannelShortName & "'>"
                End If

                If rsAuthor("ClassID") <> -1 Then
                    strLink = strLink & Replace(Character_Class, "{$Text}", "<a href='" & GetClassUrl(rsAuthor("ParentDir"), rsAuthor("ClassDir"), rsAuthor("ClassID"), rsAuthor("ClassPurview")) & "'>" & rsAuthor("ClassName") & "</a>")
                End If

                If Left(ChannelUrl, 1) <> "/" Then
                    strLink = strLink & "<a href='"
                Else
                    strLink = strLink & "<a href='http://" & Trim(Request.ServerVariables("HTTP_HOST"))
                End If
                strLink = strLink & GetPhotoUrl(rsAuthor("ParentDir"), rsAuthor("ClassDir"), rsAuthor("UpdateTime"), rsAuthor("PhotoID"), rsAuthor("ClassPurview"), rsAuthor("InfoPurview"), rsAuthor("InfoPoint")) & "'>"
                strtmp = strtmp & ("<tr class='article_list_body'><td>" & strLink & TitleStr & "</a></td><td>" & GetAuthorInfo(rsAuthor("Author"), rsAuthor("ChannelID")) & "<td align='center' width='80'>" & FormatDateTime(rsAuthor("UpdateTime"), 2) & "</td></tr>")
                rsAuthor.MoveNext
                i = i + 1
                If i >= MaxPerPage Then Exit Do
            Loop
        End If
        rsAuthor.Close
        Set rsAuthor = Nothing
        strtmp = strtmp & "</table>"
        strtmp = strtmp & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "张图片", False)
    End Select
    GetItemList = strtmp
End Function

%>
