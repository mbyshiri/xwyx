<!--#include file="Admin_CreateCommon.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim PE_Content
Set PE_Content = New Article
PE_Content.Init
tmpPageTitle = strPageTitle    '保存页面标题到临时变量中，以做为栏目及内容页循环生成时初始值
tmpNavPath = strNavPath
ArticleID = Trim(Request("ArticleID"))
Select Case Action
Case "CreateArticle"
    Call CreateArticle
Case "CreateClass"
    Call CreateClass
Case "CreateSpecial"
    Call CreateSpecial
Case "CreateIndex"
    Call CreateIndex
Case "CreateArticle2"
    If AutoCreateType > 0 Then
        IsAutoCreate = True
        Call CreateArticle
        If ClassID > 0 Then
            ClassID = ParentPath & "," & ClassID
            Call CreateClass
        End If
        SpecialID = Trim(Request("SpecialID"))
        If SpecialID <> "" Then Call CreateSpecial
        '在生成首页前，要将栏目ID和专题ID置为0
        ClassID = 0
        arrChildID = 0
        SpecialID = 0
        Call CreateIndex

        Call CreateSiteIndex     '生成网站首页
        Call CreateSiteSpecial   '生成全站专题
    End If
Case "CreateOther" '定时生成创建除文章其他页
    TimingCreate = Trim(Request("TimingCreate"))
    TimingCreateNum = PE_CLng(Trim(Request("TimingCreateNum")))

    If Trim(Request("ChannelProperty")) <> "" Then
        CreateChannelItem = Split(Trim(Request("ChannelProperty")), ",")
        ChannelID = CreateChannelItem(0)
        CreateType = 2

        If CreateChannelItem(5) = "True" Then
            Call CreateClass
            Call CreateAllJS
        End If

        If CreateChannelItem(6) = "True" Then
            Call CreateSpecial
        End If

        If CreateChannelItem(7) = "True" Then
            Call CreateIndex
        End If

        If TimingCreateNum >= UBound(Split(TimingCreate, "$")) Then
            Call CreateSiteIndex    '生成网站首页
        End If


        TimingCreateNum = TimingCreateNum + 1
        strFileName = "Admin_Timing.asp?Action=DoTiming&TimingCreateNum=" & TimingCreateNum & "&TimingCreate=" & Trim(Request("TimingCreate"))
    Else    '采集后生成
        CreateNum = PE_CLng(Trim(Request("CreateNum")))
        Call CreateClass
        Call CreateSpecial
        Call CreateIndex
        Call CreateSiteIndex     '生成网站首页
        '生成所有JS
        Call CreateAllJS
        CreateNum = CreateNum + 1
        strFileName = "Admin_Collection.asp?Action=CreateItemHtml&CollectionCreateHTML=" & Trim(Request("CollectionCreateHTML")) & "&CreateNum=" & CreateNum & "&TimingCreate=" & Trim(Request("TimingCreate"))
    End If

    If Trim(Request("TimingCreate")) <> "" Or Trim(Request("CollectionCreateHTML")) <> "" Then
        Call Refresh(strFileName,5)		
        'Response.Write "<meta http-equiv=""refresh"" content=5;url='" & strFileName & "'>" & vbCrLf
    End If

Case Else
    FoundErr = True
    ErrMsg = ErrMsg & "<li>参数错误！</li>"
End Select

Call ShowProcess

Response.Write "</body></html>"
Set PE_Content = Nothing
Call CloseConn


Sub CreateArticle()
    'On Error Resume Next
    ChannelID = PE_CLng(Request("ChannelID"))

    Dim sql, strFields, ArticlePath
    Dim strArticleContent
    Dim tmpArticle, tmpTemplateID

    tmpTemplateID = 0

    sql = "select * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID

    If IsAutoCreate = False Then
        Response.Write "<b>正在生成" & ChannelShortName & "页面……请稍候！<font color='red'>在此过程中请勿刷新此页面！！！</font></b><br>"
        Response.Flush
    End If

    Select Case CreateType
    Case 1 '选定的文章
        If IsValidID(ArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请正确指定要生成的" & ChannelShortName & "ID</li>"
            Exit Sub
        End If
        If InStr(ArticleID, ",") > 0 Then
            sql = sql & " and ArticleID in (" & ArticleID & ")"
        Else
            sql = sql & " and ArticleID=" & ArticleID
        End If
        strUrlParameter = "&ArticleID=" & ArticleID
    Case 2 '选定的栏目
        ClassID = PE_CLng(Trim(Request("ClassID")))
        If ClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要生成的栏目ID</li>"
            Exit Sub
        End If
        Call GetClass
        If ClassPurview > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此栏目不是开放栏目，所以此栏目下的文章不能生成HTML！"
        End If
        If FoundErr = True Then Exit Sub
        If InStr(arrChildID, ",") > 0 Then
            sql = sql & " and ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and ClassID=" & ClassID
        End If
    Case 3 '所有文章
        
    Case 4 '最新的文章
        Dim TopNew
        TopNew = PE_CLng(Trim(Request("TopNew")))
        If TopNew <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定有效的数目！"
            Exit Sub
        End If
        sql = "select top " & TopNew & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID & ""
        strUrlParameter = "&TopNew=" & TopNew
    Case 5 '指定更新时间
        Dim BeginDate, EndDate
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        If Not (IsDate(BeginDate) And IsDate(EndDate)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入有效的日期！</li>"
            Exit Sub
        End If
        If SystemDatabaseType = "SQL" Then
            sql = sql & " and UpdateTime between '" & BeginDate & "' and '" & EndDate & "'"
        Else
            sql = sql & " and UpdateTime between #" & BeginDate & "# and #" & EndDate & "#"
        End If
        strUrlParameter = "&BeginDate=" & Replace(BeginDate,"/","-") & "&EndDate=" & Replace(EndDate,"/","-")
    Case 6 '指定ID范围
        Dim BeginID, EndID
        BeginID = Trim(Request("BeginID"))
        EndID = Trim(Request("EndID"))
        If Not (IsNumeric(BeginID) And IsNumeric(EndID)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入数字！</li>"
            Exit Sub
        End If
        sql = sql & " and ArticleID between " & BeginID & " and " & EndID & ""
        strUrlParameter = "&BeginID=" & BeginID & "&EndID=" & EndID
    Case 7 '采集生成文章
        TimingCreate = Trim(Request("TimingCreate"))
        CollectionCreateHTML = Trim(Request("CollectionCreateHTML"))
        CreateNum = PE_CLng(Trim(Request("CreateNum")))
        IsShowReturn = True

        If CollectionCreateHTML = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>请指定要生成的数目！"
            Exit Sub
        Else
            ChannelID = PE_CLng(Trim(Request("ChannelID")))
            ClassID = PE_CLng(Trim(Request("ClassID")))
            SpecialID = ReplaceBadChar(Trim(Request("SpecialID")))
            ArticleNum = PE_CLng(Trim(Request("ArticleNum")))

            sql = "select top " & ArticleNum & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and ChannelID=" & ChannelID & " and ClassID=" & ClassID & ""
        End If
        strUrlParameter = "&CollectionCreateHTML=" & CollectionCreateHTML & "&CreateNum=" & CreateNum & "&ArticleNum=" & ArticleNum & "&TimingCreate=" & TimingCreate

    Case 8 '定时生成文章
        TimingCreate = Trim(Request("TimingCreate"))
        ChannelProperty = Trim(Request("ChannelProperty"))
        TimingCreateNum = PE_CLng(Trim(Request("TimingCreateNum")))
        IsShowReturn = True
        arrChannelProperty = Split(ChannelProperty, ",")
        ChannelID = arrChannelProperty(0)
        CreateItemType = arrChannelProperty(2)
        CreateItemTopNewNum = arrChannelProperty(3)
        CreateItemDate = arrChannelProperty(4)
        Select Case CreateItemType
        Case 1
             sql = "select top " & CreateItemTopNewNum & " *  from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and ChannelID=" & ChannelID & " order by UpdateTime desc,ClassID asc,TemplateID asc,ArticleID asc"
        Case 2
            sql = sql & " DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<" & CreateItemDate & ""
        Case 3
			sql = "select top " & MaxPerPage_Create & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID & ""
            sql = sql & " and (CreateTime is null or CreateTime<=UpdateTime)"
        Case 4
            
        End Select
        strUrlParameter = "&TimingCreate=" & TimingCreate & "&TimingCreateNum=" & TimingCreateNum & "&ChannelProperty=" & Trim(Request("ChannelProperty"))
    Case 9 '所有未生成的文章
        sql = "select top " & MaxPerPage_Create & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID & ""
		sql = sql & " and (CreateTime is null or CreateTime<=UpdateTime)"
    Case Else
        Response.Write "参数错误！"
        Exit Sub
    End Select
    If CreateType = 4 Or CreateType = 7 Then
        sql = sql & " order by UpdateTime desc,ClassID,ArticleID"
    Else
        sql = sql & " order by ClassID,ArticleID"
    End If
    Set rsArticle = Server.CreateObject("ADODB.Recordset")
    rsArticle.Open sql, Conn, 1, 1
    If rsArticle.Bof And rsArticle.EOF Then
        TotalCreate = 0
		iTotalPage = 0
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    Else
        If CreateType = 9 Or (CreateType = 8 And CreateItemType = 3)Then
			TotalCreate = PE_Clng(Conn.Execute("select count(*) from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID & " and (CreateTime is null or CreateTime<=UpdateTime)")(0))
		Else
			TotalCreate = rsArticle.RecordCount
		End If
		
    End If

    PageTitle = "正文" '得到频道标题
    strFileName = ChannelUrl_ASPFile & "/ShowArticle.asp" '得到路径
    strTemplate = GetTemplate(ChannelID, 3, tmpTemplateID) '得到频道中正文的默认模板
    
    Call MoveRecord(rsArticle)
    Call ShowTotalCreate(ChannelItemUnit & ChannelShortName)
    Do While Not rsArticle.EOF
        FoundErr = False
        ArticleID = rsArticle("ArticleID")
        ClassID = rsArticle("ClassID")
        If CreateType = 7 Then ChannelID = rsArticle("ChannelID")
        strNavPath = tmpNavPath
        If ChannelID <> PrevChannelID Then
            Call GetChannel(ChannelID)
            PrevChannelID = ChannelID
        End If
        Call GetClass
        strPageTitle = tmpPageTitle
        iCount = iCount + 1

        If ClassPurview > 0 Or rsArticle("InfoPurview") > 0 Or rsArticle("InfoPoint") > 0 Then
            Response.Write "<li><font color='red'>ID为 " & rsArticle("ArticleID") & " 的" & ChannelShortName & "因为设置了阅读权限，所以没有生成。</font></li>"
            Response.Flush
        Else
            SpecialID = 0
            CurrentPage = 1
            ArticlePath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle("UpdateTime"))

            If CreateMultiFolder(ArticlePath) = False Then
                Response.Write "请检查服务器。系统不能创建生成文件所需要的文件夹。"
                Exit Sub
            End If
            ArticlePath = ArticlePath & GetItemFileName(FileNameType, ChannelDir, rsArticle("UpdateTime"), ArticleID)
                
            tmpFileName = ArticlePath & FileExt_Item

            '生成页面时判定转向连接
            If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then
                Call WriteToFile(tmpFileName, PE_Content.GetLinkUrlContent(rsArticle("LinkUrl"), ArticleID))
                Response.Write "<li>成功生成第 <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & "。&nbsp;&nbsp;ID：" & ArticleID & " &nbsp;&nbsp;标题：" & rsArticle("Title") & " &nbsp;&nbsp;地址：<a href='" & tmpFileName & "' target='_blank'>" & tmpFileName & "</a></li>" & vbCrLf
                Response.Flush
            Else
                ArticleUrl = GetArticleUrl(ParentDir, ClassDir, rsArticle("UpdateTime"), ArticleID, ClassPurview, rsArticle("InfoPurview"), rsArticle("InfoPoint"))

                SkinID = GetIDByDefault(rsArticle("SkinID"), DefaultItemSkin)
                TemplateID = GetIDByDefault(rsArticle("TemplateID"), DefaultItemTemplate)

                If Trim(rsArticle("TitleIntact")) <> "" Then
                    ArticleTitle = Replace(Replace(Replace(Replace(rsArticle("TitleIntact") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
                Else
                    ArticleTitle = Replace(Replace(Replace(Replace(rsArticle("Title") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
                End If

                If TemplateID <> tmpTemplateID Then
                    strTemplate = GetTemplate(ChannelID, 3, TemplateID)
                    tmpTemplateID = TemplateID
                End If
                strHtml = strTemplate
                Call PE_Content.GetHtml_Article
                tmpArticle = PE_Content.ReplaceContentLabel(strHtml)
                If InStr(tmpArticle, "{$ShowPageContent}") > 0 Then tmpArticle = Replace(tmpArticle, "{$ShowPageContent}", "")
                '写入生成地址
                Call WriteToFile(tmpFileName, tmpArticle)
                Response.Write "<li>成功生成第 <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & "。&nbsp;&nbsp;ID：" & ArticleID & " &nbsp;&nbsp;标题：" & rsArticle("Title") & " &nbsp;&nbsp;地址：<a href='" & tmpFileName & "' target='_blank'>" & tmpFileName & "</a></li>" & vbCrLf
                Response.Flush
                
                For CurrentPage = 2 To PE_Content.TotalPage
                    tmpFileName = ArticlePath & "_" & CurrentPage & FileExt_Item
                    tmpArticle = PE_Content.ReplaceContentLabel(strHtml)
                    If InStr(tmpArticle, "{$ShowPageContent}") > 0 Then tmpArticle = Replace(tmpArticle, "{$ShowPageContent}", "")
                    Call WriteToFile(tmpFileName, tmpArticle)
                    Response.Write "<br>&nbsp;&nbsp;&nbsp;成功生成第 <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & "的第 <font color='blue'>" & CurrentPage & "</font> 页：<a href='" & tmpFileName & "' target='_blank'>" & tmpFileName & "</a>" & vbCrLf
                    Response.Flush
                Next
            End If
            '生成内容结束，更新内容的生成时间
            Conn.Execute ("update PE_Article set CreateTime=" & PE_Now & " where ArticleID=" & ArticleID)

        End If
        If Response.IsClientConnected = False Then Exit Do
        If iCount Mod MaxPerPage_Create = 0 Then Exit Do
        rsArticle.MoveNext
    Loop
    rsArticle.Close
    Set rsArticle = Nothing
End Sub
%>
