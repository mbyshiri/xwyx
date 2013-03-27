<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim SoftID, SoftName, SoftUrl
Dim rsSoft

Class Soft

Private rsClass

Sub Init()
    FoundErr = False
    ErrMsg = ""
    PrevChannelID = ChannelID
    ChannelShortName = "软件"
    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
     
    '*****************************
    '读取语言包中的字符设置
    ChannelShortName = XmlText_Class("ChannelShortName", "软件")
    strListStr_Font = XmlText_Class("SoftList/UpdateTimeColor_New", "color=""red""")
    strTop = XmlText_Class("SoftList/t4", "固顶")
    strElite = XmlText_Class("SoftList/t3", "推荐")
    strCommon = XmlText_Class("SoftList/t5", "普通")
    strHot = XmlText_Class("SoftList/t7", "热点")
    strNew = XmlText_Class("SoftList/t6", "最新")
    strTop2 = XmlText_Class("SoftList/Top", " 顶")
    strElite2 = XmlText_Class("SoftList/Elite", " 荐")
    strHot2 = XmlText_Class("SoftList/Hot", " 热")
    Character_Author = XmlText("Soft", "Include/Author", "[{$Text}]")
    Character_Date = XmlText("Soft", "Include/Date", "[{$Text}]")
    Character_Hits = XmlText("Soft", "Include/Hits", "[{$Text}]")
    Character_Class = XmlText("Soft", "Include/ClassChar", "[{$Text}]")
    SearchResult_Content_NoPurview = XmlText("BaseText", "SearchPurviewContent", "此内容需要有指定权限才可以预览")
    SearchResult_ContentLenth = PE_CLng(XmlText_Class("ShowSearch/Content_Lenght", "200"))
    strList_Content_Div = XmlText_Class("SoftList/Content_DIV", "style=""padding:0px 20px""")
    strList_Title = R_XmlText_Class("SoftList/Title", "{$ChannelShortName}名称：{$Title}{$br}作&nbsp;&nbsp;&nbsp;&nbsp;者：{$Author}{$br}更新时间：{$UpdateTime}")
    strComment = XmlText_Class("SoftList/CommentLink", "<font color=""red"">评论</font>")
    '*****************************
    
    strPageTitle = SiteTitle
    
    Call GetChannel(ChannelID)
    HtmlDir = InstallDir & ChannelDir
    If Trim(ChannelName) <> "" And ShowNameOnPath <> False Then
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath' href='"
        If UseCreateHTML > 0 Then
            strNavPath = strNavPath & ChannelUrl & "/Index" & FileExt_Index
        Else
            strNavPath = strNavPath & ChannelUrl_ASPFile & "/Index.asp"
        End If
        strNavPath = strNavPath & "'>" & ChannelName & "</a>"
        strPageTitle = strPageTitle & " >> " & ChannelName
    End If
    
End Sub
'=================================================
'函数名：ShowChannelCount
'作  用：显示频道统计信息
'参  数：无
'=================================================
Private Function GetChannelCount()
    GetChannelCount = Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ChannelCount", "{$ChannelShortName}总数： {$ItemChecked_Channel} {$ChannelItemUnit}<br>待审{$ChannelShortName}： {$UnItemChecked} {$ChannelItemUnit}<br>评论总数： {$CommentCount_Channel} 条<br>专题总数： {$SpecialCount_Channel} 个<br>{$ChannelShortName}下载： {$HitsCount_Channel} 人次<br>"), "{$ItemChecked_Channel}", ItemChecked_Channel), "{$ChannelItemUnit}", ChannelItemUnit), "{$UnItemChecked}", treatAuditing("Soft", ChannelID)), "{$CommentCount_Channel}", CommentCount_Channel), "{$SpecialCount_Channel}", SpecialCount_Channel), "{$HitsCount_Channel}", "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?Action=Count'></script>")
End Function
'**************************************************
'函数名：treatAuditing
'作  用：待审核函数
'参  数：ModuleName ----表名
'        ChannelID ---- 频道ID
'返回值：待审核项目数
'**************************************************
Private Function treatAuditing(ByVal ModuleName, ByVal ChannelID)
    Dim trs
    Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & ChannelID & " and Status > -1 and Status < 3 and Deleted=" & PE_False & "")
    treatAuditing = trs(0)
    If IsNull(treatAuditing) Then treatAuditing = 0
    Set trs = Nothing
End Function

Private Function GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, Author, DateNum, OrderType, ShowClassName, IsPicUrl)
    Dim strSql, IDOrder
    iSpecialID = PE_CLng(iSpecialID)
    If IsValidID(iChannelID) = False Then
        iChannelID = 0
    Else
        iChannelID = ReplaceLabelBadChar(iChannelID)
    End If  
    If IsValidID(arrClassID) = False Then
        arrClassID = 0
    Else
        arrClassID = ReplaceLabelBadChar(arrClassID)
    End If	
    If iSpecialID > 0 Then
        strSql = strSql & " from PE_InfoS I inner join (PE_Soft S left join PE_Class C on S.ClassID=C.ClassID) on I.ItemID=S.SoftID"
    Else
        strSql = strSql & " from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID"
    End If
    strSql = strSql & " where S.Deleted=" & PE_False & " and S.Status=3"
    If iChannelID > 0 Then
        strSql = strSql & " and S.ChannelID=" & iChannelID
    End If
    If arrClassID <> "0" Then
        If InStr(arrClassID, ",") = 0 And IncludeChild = True Then
            Dim trs
            Set trs = Conn.Execute("select arrChildID from PE_Class where ClassID=" & PE_CLng(arrClassID) & "")
            If trs.BOF And trs.EOF Then
                arrClassID = "0"
            Else
                If IsNull(trs(0)) Or Trim(trs(0)) = "" Then
                    arrClassID = "0"
                Else
                    arrClassID = trs(0)
                End If
            End If
            Set trs = Nothing
        End If
        
        If InStr(arrClassID, ",") > 0 Then
            strSql = strSql & " and S.ClassID in (" & FilterArrNull(arrClassID, ",") & ")"
        Else
            If PE_CLng(arrClassID) > 0 Then strSql = strSql & " and S.ClassID=" & PE_CLng(arrClassID)
        End If
    End If
    If iSpecialID > 0 Then
        strSql = strSql & " and I.ModuleType=2 and I.SpecialID=" & iSpecialID
    End If
    If IsHot = True Then
        strSql = strSql & " and S.Hits>=" & HitsOfHot
    End If
    If IsElite = True Then
        strSql = strSql & " and S.Elite=" & PE_True
    End If
    If Trim(Author) <> "" Then
        strSql = strSql & " and S.Author='" & ReplaceBadChar(Author) & "'"
    End If
    If DateNum > 0 Then
        strSql = strSql & " and DateDiff(" & PE_DatePart_D & ",S.UpdateTime," & PE_Now & ")<" & DateNum
    End If

    If IsPicUrl = True Then
        strSql = strSql & " and S.SoftPicUrl<>'' "
    End If

    strSql = strSql & " order by S.OnTop " & PE_OrderType & ","
    Select Case PE_CLng(OrderType)
    Case 1, 2

    Case 3
        strSql = strSql & "S.UpdateTime desc,"
    Case 4
        strSql = strSql & "S.UpdateTime asc,"
    Case 5
        strSql = strSql & "S.Hits desc,"
    Case 6
        strSql = strSql & "S.Hits asc,"
    Case 7
        strSql = strSql & "S.CommentCount desc,"
    Case 8
        strSql = strSql & "S.CommentCount asc,"
    Case Else

    End Select
    If OrderType = 2 Then
        IDOrder = "asc"
    Else
        IDOrder = "desc"
    End If
    If iSpecialID > 0 Then
        strSql = strSql & "I.InfoID " & IDOrder
    Else
        strSql = strSql & "S.SoftID " & IDOrder
    End If
    GetSqlStr = strSql
End Function

'=================================================
'函数名：GetSoftList
'作  用：显示软件名称等信息
'参  数：
'0        iChannelID ---- 频道ID
'1        arrClassID ---- 栏目ID数组，0为所有栏目
'2        IncludeChild ---- 是否包含子栏目，仅当arrClassID为单个栏目ID时才有效，True----包含子栏目，False----不包含
'3        iSpecialID ---- 专题ID，0为所有下载（含非专题下载），如果为大于0，则只显示相应专题的下载
'4        UrlType ---- 链接地址类型，0为相对路径，1为带网址的绝对路径，不对外公开，4.03时为ShowAllSoft
'5        SoftNum ---- 下载数，若大于0，则只查询前几个下载
'6        IsHot ---- 是否是热门下载，True为只显示热门下载，False为显示所有下载
'7        IsElite ---- 是否是推荐下载，True为只显示推荐下载，False为显示所有下载
'8        Author ---- 作者姓名，如果不为空，则只显示指定作者的下载，用于作者软件集
'9        DateNum ---- 日期范围，如果大于0，则只显示最近几天内更新的下载
'10       OrderType ---- 排序方式，1--按下载ID降序，2--按下载ID升序，3--按更新时间降序，4--按更新时间升序，5--按点击数降序，6--按点击数升序，7--按评论数降序，8--按评论数升序
'11       ShowType ---- 显示方式，1为普通样式，2为表格式，3为各项独立式，4为输出DIV格式，5为输出RSS格式
'12       TitleLen ---- 标题最多字符数，一个汉字=两个英文字符，若为0，则显示完整标题
'13       ContentLen ---- 软件简介最多字符数，一个汉字=两个英文字符，为0时不显示。
'14       ShowClassName ---- 是否显示所属栏目名称，True为显示，False为不显示
'15       ShowPropertyType ---- 显示下载属性（固顶/推荐/普通）的方式，0为不显示，1为小图片，2为符号
'16       ShowAuthor ---- 是否显示下载作者，True为显示，False为不显示
'17       ShowDateType ---- 显示更新日期的样式，0为不显示，1为显示年月日，2为只显示月日，3为以“月-日”方式显示月日。
'18       ShowHits ---- 是否显示下载点击数，True为显示，False为不显示
'19       ShowHotSign ---- 是否显示热门下载标志，True为显示，False为不显示
'20       ShowNewSign ---- 是否显示新下载标志，True为显示，False为不显示
'21       ShowTips ---- 是否显示作者、更新日期、点击数等浮动提示信息，True为显示，False为不显示
'22       UsePage ---- 是否分页显示，True为分页显示，False为不分页显示，每页显示的软件数量由MaxPerPage指定
'23       OpenType ---- 下载打开方式，0为在原窗口打开，1为在新窗口打开
'24       Cols ---- 每行的列数。超过此列数就换行。
'25       CssNameA ---- 列表中文字链接调用的CSS类名
'26       CssName1 ---- 列表中奇数行的CSS效果的类名
'27       CssName2 ---- 列表中偶数行的CSS效果的类名
'=================================================
Public Function GetSoftList(iChannelID, arrClassID, IncludeChild, iSpecialID, UrlType, SoftNum, IsHot, IsElite, Author, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, UsePage, OpenType, Cols, CssNameA, CssName1, CssName2)
    Dim sqlInfo, rsInfoList, strInfoList, CssName, iCount, iNumber, InfoUrl
    Dim strProperty, strTitle, strLink, strAuthor, strUpdateTime, strHits, strHotSign, strNewSign, strContent, strClassName
    Dim TDWidth_Author, TdWidth_Date

    TDWidth_Author = 10 * AuthorInfoLen
    TdWidth_Date = GetTDWidth_Date(ShowDateType)

    iCount = 0
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)

    If ShowType = 5 Then UrlType = 1
    If TitleLen < 0 Or TitleLen > 200 Then TitleLen = 50
    If IsNull(CssNameA) Then CssNameA = "listA"
    If IsNull(CssName1) Then CssName1 = "listbg"
    If IsNull(CssName2) Then CssName2 = "listbg2"

    FoundErr = False
    If iChannelID <> PrevChannelID Or ChannelID = 0 Then
        Call GetChannel(iChannelID)
    End If
    PrevChannelID = iChannelID
    If FoundErr = True Then
        GetSoftList = ErrMsg
        Exit Function
    End If

    sqlInfo = "select"
    If SoftNum > 0 Then
        sqlInfo = sqlInfo & " top " & SoftNum
    End If
    sqlInfo = sqlInfo & " S.ChannelID,S.ClassID,S.SoftID,S.SoftName,S.SoftVersion,S.Author,S.UpdateTime,S.Hits,S.OnTop,S.Elite"
    If ContentLen > 0 Then
        sqlInfo = sqlInfo & ",S.SoftIntro"
    End If
    sqlInfo = sqlInfo & ",C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview"
    sqlInfo = sqlInfo & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, Author, DateNum, OrderType, ShowClassName, False)
    Set rsInfoList = Server.CreateObject("ADODB.Recordset")
    rsInfoList.Open sqlInfo, Conn, 1, 1
    If rsInfoList.BOF And rsInfoList.EOF Then
        If UsePage = True Then totalPut = 0
        If ShowType < 5 Then
            strInfoList = GetInfoList_StrNoItem(arrClassID, iSpecialID, IsHot, IsElite, strHot, strElite)
        End If
        rsInfoList.Close
        Set rsInfoList = Nothing
        GetSoftList = strInfoList
        Exit Function
    End If
    If UsePage = True And ShowType < 5 Then
        totalPut = rsInfoList.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                iMod = 0
                If CurrentPage > UpdatePages Then
                    iMod = totalPut Mod MaxPerPage
                    If iMod <> 0 Then iMod = MaxPerPage - iMod
                End If
                rsInfoList.Move (CurrentPage - 1) * MaxPerPage - iMod
            Else
                CurrentPage = 1
            End If
        End If
    End If

    CssName = CssName1

    If ShowType = 5 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
    If ShowType = 2 Or Cols > 1 Then
        strInfoList = "<table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr>"
    Else
        strInfoList = ""
    End If

    Do While Not rsInfoList.EOF
        If iChannelID = 0 Then
            If rsInfoList("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsInfoList("ChannelID"))
                PrevChannelID = rsInfoList("ChannelID")
            End If
        End If
        If UsePage = True Then
            iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1
        Else
            iNumber = iCount + 1
        End If


        ChannelUrl = UrlPrefix(UrlType, ChannelUrl) & ChannelUrl
        ChannelUrl_ASPFile = UrlPrefix(UrlType, ChannelUrl_ASPFile) & ChannelUrl_ASPFile
        InfoUrl = GetSoftUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("UpdateTime"), rsInfoList("SoftID"))
        If rsInfoList("SoftVersion") = "" Or IsNull(rsInfoList("SoftVersion")) then
            strTitle = GetInfoList_GetStrTitle(rsInfoList("SoftName"), TitleLen, 0, "")
        Else		
            strTitle = GetInfoList_GetStrTitle(rsInfoList("SoftName") & " " & rsInfoList("SoftVersion"), TitleLen, 0, "")
        End If	
        If ShowType < 5 Then

            strProperty = GetInfoList_GetStrProperty(ShowPropertyType, rsInfoList("OnTop"), rsInfoList("Elite"), iNumber, strCommon, strTop, strElite)
            strHotSign = GetInfoList_GetStrHotSign(ShowHotSign, rsInfoList("Hits"), strHot)
            strNewSign = GetInfoList_GetStrNewSign(ShowNewSign, rsInfoList("UpdateTime"), strNew)
            strAuthor = GetSubStr(rsInfoList("Author"), AuthorInfoLen, True)
            strUpdateTime = GetInfoList_GetStrUpdateTime(rsInfoList("UpdateTime"), ShowDateType)
            strHits = rsInfoList("Hits")
            If ShowType = 3 Or ShowType = 4 Then
                strAuthor = GetInfoList_GetStrAuthor_Xml(ShowAuthor, strAuthor)
                strUpdateTime = GetInfoList_GetStrUpdateTime_Xml(ShowDateType, strUpdateTime)
                strHits = GetInfoList_GetStrHits_Xml(ShowHits, strHits)
            End If

            strLink = ""
            If ShowClassName = True Then
                strLink = strLink & GetInfoList_GetStrClassLink(Character_Class, CssNameA, rsInfoList("ClassID"), rsInfoList("ClassName"), GetClassUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("ClassID"), rsInfoList("ClassPurview")))
            End If
            strLink = strLink & GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, CssNameA, strTitle, InfoUrl, rsInfoList("SoftName"), rsInfoList("Author"), rsInfoList("UpdateTime"))

            strContent = ""
            Select Case PE_CLng(ShowType)
            Case 1, 3, 4
                If ContentLen > 0 Then
                    strContent = strContent & "<div " & strList_Content_Div & ">"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("SoftIntro"), "")
                    strContent = strContent & "</div>"
                End If
            Case 2
                If ContentLen > 0 Then
                    strContent = strContent & "<tr><td colspan=""10"" class=""" & CssName & """>"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("SoftIntro"), "")
                    strContent = strContent & "</td></tr>"
                End If
            End Select

        ElseIf ShowType = 5 Then

            strTitle = xml_nohtml(strTitle)
            strLink = InfoUrl
            If ContentLen > 0 Then
                strContent = Left(Replace(Replace(xml_nohtml(rsInfoList("SoftIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen)
            End If
            strAuthor = GetInfoList_GetStrAuthor_RSS(Author)
            If ShowClassName = True And rsInfoList("ClassID") <> -1 Then
                strClassName = xml_nohtml(rsInfoList("ClassName"))
            Else
                strClassName = ""
            End If
            strUpdateTime = GetInfoList_GetStrUpdateTime(rsInfoList("UpdateTime"), ShowDateType)

        End If

        Select Case PE_CLng(ShowType)
        Case 1
            If Cols > 1 Then
                strInfoList = strInfoList & "<td valign=""top"" class=""" & CssName & """>"
            End If
            strInfoList = strInfoList & strProperty & "&nbsp;" & strLink
            strInfoList = strInfoList & GetInfoList_GetStrAuthorDateHits(ShowAuthor, ShowDateType, ShowHits, rsInfoList("Author"), strUpdateTime, strHits, rsInfoList("ChannelID"))
            strInfoList = strInfoList & strHotSign & strNewSign & strContent
            strInfoList = strInfoList & "<br />"

            iCount = iCount + 1
            If Cols > 1 Then
                strInfoList = strInfoList & "</td>"
                If iCount Mod Cols = 0 Then
                    strInfoList = strInfoList & "</tr><tr>"
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                End If
            End If
        Case 2
            If strProperty <> "" Then
                strInfoList = strInfoList & "<td width=""10"" valign=""top"" class=""" & CssName & """>" & strProperty & "</td>"
            End If
            strInfoList = strInfoList & "<td class=""" & CssName & """>" & strLink & strHotSign & strNewSign & "</td>"
            If ShowAuthor = True Then
                strInfoList = strInfoList & "<td align=""center"" class=""" & CssName & """ width=""" & TDWidth_Author & """>" & strAuthor & "</td>"
            End If
            If ShowDateType > 0 Then
                strInfoList = strInfoList & "<td align=""right"" class=""" & CssName & """ width=""" & TdWidth_Date & """>" & strUpdateTime & "</td>"
            End If
            If ShowHits = True Then
                strInfoList = strInfoList & "<td align=""center"" class=""" & CssName & """ width=""40"">" & strHits & "</td>"
            End If

            iCount = iCount + 1
            If (iCount Mod Cols = 0) Or ContentLen > 0 Then
                strInfoList = strInfoList & "</tr>"
                strInfoList = strInfoList & strContent
                strInfoList = strInfoList & "<tr>"
                If iCount Mod (Cols * 2) = 0 Then
                    CssName = CssName1
                Else
                    CssName = CssName2
                End If
            End If
        Case 3
            If Cols > 1 Then
                strInfoList = strInfoList & "<td valign=""top"" class=""" & CssName & """>"
            End If
            strInfoList = strInfoList & strProperty & "&nbsp;" & strLink
            strInfoList = strInfoList & strAuthor & strUpdateTime & strHits
            strInfoList = strInfoList & strHotSign & strNewSign & strContent
            strInfoList = strInfoList & "<br />"

            iCount = iCount + 1
            If Cols > 1 Then
                strInfoList = strInfoList & "</td>"
                If iCount Mod Cols = 0 Then
                    strInfoList = strInfoList & "</tr><tr>"
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                End If
            End If
        Case 4 '输出DIV
            strInfoList = strInfoList & "<div class=""" & CssName & """>"
            strInfoList = strInfoList & strProperty & "&nbsp;" & strLink
            strInfoList = strInfoList & strAuthor & strUpdateTime & strHits
            strInfoList = strInfoList & strHotSign & strNewSign & strContent
            strInfoList = strInfoList & "</div>"

            iCount = iCount + 1
            If iCount Mod 2 = 0 Then
                CssName = CssName1
            Else
                CssName = CssName2
            End If
        Case 5 '输出RSS
            strInfoList = strInfoList & GetInfoList_GetStrRSS(strTitle, strLink, strContent, strAuthor, strClassName, strUpdateTime)
            iCount = iCount + 1
        End Select
        rsInfoList.MoveNext
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    If ShowType = 2 Or Cols > 1 Then
        strInfoList = strInfoList & "</tr></table>"
    End If
    rsInfoList.Close
    Set rsInfoList = Nothing
    If ShowType = 5 And RssCodeType = False Then strInfoList = unicode(strInfoList)
    GetSoftList = strInfoList
End Function


'=================================================
'函数名：GetPicSoft
'作  用：显示图片下载
'参  数：
'0        iChannelID ---- 频道ID
'1        arrClassID ---- 栏目ID数组，0为所有栏目
'2        IncludeChild ---- 是否包含子栏目，仅当arrClassID为单个栏目ID时才有效，True----包含子栏目，False----不包含
'3        iSpecialID ---- 专题ID，0为所有下载（含非专题下载），如果为大于0，则只显示相应专题的下载
'4        SoftNum ---- 最多显示多少个软件
'5        IsHot ---- 是否是热门下载
'6        IsElite ---- 是否是推荐下载
'7        DateNum ---- 日期范围，如果大于0，则只显示最近几天内更新的下载
'8        OrderType ---- 排序方式，1--按下载ID降序，2--按下载ID升序，3--按更新时间降序，4--按更新时间升序，5--按点击数降序，6--按点击数升序，7--按评论数降序，8--按评论数升序
'9        ShowType ---- 显示方式。1为图片+标题+内容简介：上下排列；2为（图片+标题：上下排列）+内容简介：左右排列，3为图片+（标题+内容简介：上下排列）：左右排列，4为输出DIV格式，5为输出RSS格式
'10       ImgWidth ---- 图片宽度
'11       ImgHeight ---- 图片高度
'12       TitleLen ---- 标题最多字符数，一个汉字=两个英文字符。若为0，则不显示标题；若为-1，则显示完整标题
'13       ContentLen ---- 内容最多字符数，一个汉字=两个英文字符。若为0，则不显示内容简介
'14       ShowTips ---- 是否显示作者、更新时间、点击数等提示信息，True为显示，False为不显示
'15       Cols ---- 每行的列数。超过此列数就换行。
'16       UrlType ---- 链接地址类型，0为相对路径，1为带网址的绝对路径。
'=================================================
Public Function GetPicSoft(iChannelID, arrClassID, IncludeChild, iSpecialID, SoftNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType)
    Dim sqlPic, rsPic, iCount, strPic, strLink, strAuthor, InfoUrl
    Dim strSoftPicUrl, strLink_SoftPicUrl, strTitle, strLink_Title, strContent, strLink_Content

    iCount = 0
    SoftNum = PE_CLng(SoftNum)
    ShowType = PE_CLng(ShowType)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)

    If SoftNum < 0 Or SoftNum >= 100 Then SoftNum = 10
    If ShowType < 1 And ShowType > 5 Then ShowType = 2
    If ImgWidth < 0 Or ImgWidth > 1000 Then ImgWidth = 150
    If ImgHeight < 0 Or ImgHeight > 1000 Then ImgHeight = 150
    If ShowType = 5 Then UrlType = 1
    If Cols <= 0 Then Cols = 5

    FoundErr = False
    If iChannelID <> PrevChannelID Or ChannelID = 0 Then
        Call GetChannel(iChannelID)
    End If
    PrevChannelID = iChannelID
    If FoundErr = True Then
        GetPicSoft = ErrMsg
        Exit Function
    End If

    sqlPic = "select"
    If SoftNum > 0 Then
        sqlPic = sqlPic & " top " & SoftNum
    End If
    sqlPic = sqlPic & " S.ChannelID,S.ClassID,S.SoftID,S.SoftName,S.SoftVersion,S.Author,S.UpdateTime,S.Hits,S.SoftPicUrl"
    If ContentLen > 0 Then
        sqlPic = sqlPic & ",S.SoftIntro"
    End If
    sqlPic = sqlPic & ",C.ClassName,C.ClassDir,C.ParentDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    If ShowType < 4 Then strPic = "<table width='100%' cellpadding='0' cellspacing='5' border='0' align='center'><tr valign='top'>"
    If rsPic.BOF And rsPic.EOF Then
        If SoftNum = 0 Then totalPut = 0
        If ShowType < 4 Then
            strPic = strPic & "<td align='center'><img class='pic2' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>" & R_XmlText_Class("PicSoft/NoFound", "没有任何图片{$ChannelShortName}") & "</td></tr></table>"
        ElseIf ShowType = 4 Then
            strPic = strPic & "<div class=""pic_soft""><img class='pic2' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>" & R_XmlText_Class("PicSoft/NoFound", "没有任何图片{$ChannelShortName}") & "</div>"
        End If
        rsPic.Close
        Set rsPic = Nothing
        GetPicSoft = strPic
        Exit Function
    End If

    If SoftNum = 0 And ShowType < 5 Then
        totalPut = rsPic.RecordCount
        If totalPut > 0 Then
            If CurrentPage < 1 Then
                CurrentPage = 1
            End If
            If (CurrentPage - 1) * MaxPerPage > totalPut Then
                If (totalPut Mod MaxPerPage) = 0 Then
                    CurrentPage = totalPut \ MaxPerPage
                Else
                    CurrentPage = totalPut \ MaxPerPage + 1
                End If
            End If
            If CurrentPage > 1 Then
                If (CurrentPage - 1) * MaxPerPage < totalPut Then
                    iMod = 0
                    If CurrentPage > UpdatePages Then
                        iMod = totalPut Mod MaxPerPage
                        If iMod <> 0 Then iMod = MaxPerPage - iMod
                    End If
                    rsPic.Move (CurrentPage - 1) * MaxPerPage - iMod
                Else
                    CurrentPage = 1
                End If
            End If
        End If
    End If
    
    If ShowType = 5 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
    Do While Not rsPic.EOF
        If iChannelID = 0 Then
            If rsPic("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsPic("ChannelID"))
                PrevChannelID = rsPic("ChannelID")
            End If
        End If

        ChannelUrl = UrlPrefix(UrlType, ChannelUrl) & ChannelUrl
        ChannelUrl_ASPFile = UrlPrefix(UrlType, ChannelUrl_ASPFile) & ChannelUrl_ASPFile
        If ShowType < 5 Then
            InfoUrl = GetSoftUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("SoftID"))
            strSoftPicUrl = GetSoftPicUrl(rsPic("SoftPicUrl"), ImgWidth, ImgHeight)
            strLink_SoftPicUrl = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strSoftPicUrl, InfoUrl, rsPic("SoftName"), rsPic("Author"), rsPic("UpdateTime"))

            If ShowType = 4 Then
                strPic = strPic & "<div class=""pic_soft"">" & vbCrLf
                strPic = strPic & "<div class=""pic_soft_img"">" & strLink_SoftPicUrl & "</a></div>" & vbCrLf
            Else
                strPic = strPic & "<td align='center'>"
                strPic = strPic & strLink_SoftPicUrl
            End If

            If TitleLen <> 0 Then
                If rsPic("SoftVersion") = "" Or IsNull(rsPic("SoftVersion")) then
                    strTitle = GetInfoList_GetStrTitle(rsPic("SoftName"), TitleLen, 0, "")
                Else		
                    strTitle = GetInfoList_GetStrTitle(rsPic("SoftName") & " " & rsPic("SoftVersion"), TitleLen, 0, "")
                End If				
                strLink_Title = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strTitle, InfoUrl, rsPic("SoftName"), rsPic("Author"), rsPic("UpdateTime"))
                Select Case PE_CLng(ShowType)
                Case 1, 2
                    strPic = strPic & "<br>" & strLink_Title
                Case 3
                    strPic = strPic & "</td><td valign='top' align='left'>" & strLink_Title
                Case 4
                    strPic = strPic & "<div class=""pic_soft_title"">" & strLink_Title & "</div>" & vbCrLf
                End Select
            End If
            If ContentLen > 0 Then
                strContent = Left(Replace(Replace(nohtml(rsPic("SoftIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen) & "……"
                strLink_Content = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strContent, InfoUrl, rsPic("SoftName"), rsPic("Author"), rsPic("UpdateTime"))
                Select Case PE_CLng(ShowType)
                Case 1, 3
                    strPic = strPic & "<br><div align='left'>" & strLink_Content & "</div>"
                Case 2
                    strPic = strPic & "</td><td valign='top' align='left'>" & strLink_Content
                Case 4
                    strPic = strPic & "<div class=""pic_soft_content"">" & strLink_Content & "</div>" & vbCrLf
                End Select
            End If
            If ShowType = 4 Then
                strPic = strPic & "</div>" & vbCrLf
            Else
                strPic = strPic & "</td>"
            End If
        Else
            If rsPic("SoftVersion") = "" Or IsNull(rsPic("SoftVersion")) then
                strTitle = GetInfoList_GetStrTitle(rsPic("SoftName"), TitleLen, 0, "")
            Else		
                strTitle = GetInfoList_GetStrTitle(rsPic("SoftName") & " " & rsPic("SoftVersion"), TitleLen, 0, "")
            End If	
            strLink = GetSoftUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("SoftID"))
            strAuthor = GetInfoList_GetStrAuthor_RSS(rsPic("Author"))
            If ContentLen > 0 Then
                strContent = Left(Replace(Replace(xml_nohtml(rsPic("SoftIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen)
            End If
            strPic = strPic & GetInfoList_GetStrRSS(xml_nohtml(strTitle), strLink, strContent, strAuthor, xml_nohtml(rsPic("ClassName")), rsPic("UpdateTime"))
        End If
        rsPic.MoveNext
        iCount = iCount + 1
        If SoftNum = 0 And iCount >= MaxPerPage Then Exit Do
        If ((iCount Mod Cols = 0) And (Not rsPic.EOF)) And ShowType < 4 Then strPic = strPic & "</tr><tr valign='top'>"
    Loop

    If ShowType < 4 Then strPic = strPic & "</tr></table>"
    rsPic.Close
    Set rsPic = Nothing
    If ShowType = 5 And RssCodeType = False Then strPic = unicode(strPic)
    GetPicSoft = strPic
End Function

'=================================================
'函数名：GetSlidePicSoft
'作  用：以幻灯片效果显示图片软件
'参  数：
'0        iChannelID ---- 频道ID
'1        arrClassID ---- 栏目ID数组，0为所有栏目
'2        IncludeChild ---- 是否包含子栏目，仅当arrClassID为单个栏目ID时才有效，True----包含子栏目，False----不包含
'3        iSpecialID ---- 专题ID，0为所有软件（含非专题软件），如果为大于0，则只显示相应专题的软件
'4        SoftNum ---- 最多显示多少个软件
'5        IsHot ---- 是否是热门软件
'6        IsElite ---- 是否是推荐软件
'7        DateNum ---- 日期范围，如果大于0，则只显示最近几天内更新的软件
'8        OrderType ---- 排序方式，1--按软件ID降序，2--按软件ID升序，3--按更新时间降序，4--按更新时间升序，5--按点击数降序，6--按点击数升序，7--按评论数降序，8--按评论数升序
'9        ImgWidth ---- 图片宽度
'10       ImgHeight ---- 图片高度
'11       TitleLen ---- 软件标题字数限制，0为不显示，-1为显示完整标题
'12       iTimeOut ---- 效果变换间隔时间，以毫秒为单位
'13       effectID ---- 图片转换效果，0至22指定某一种特效，23表示随机效果
'=================================================
Public Function GetSlidePicSoft(iChannelID, arrClassID, IncludeChild, iSpecialID, SoftNum, IsHot, IsElite, DateNum, OrderType, ImgWidth, ImgHeight, TitleLen, iTimeOut, effectID)
    Dim sqlPic, rsPic, i, strPic
    Dim SoftPicUrl, strTitle

    SoftNum = PE_CLng(SoftNum)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)

    If SoftNum <= 0 Or SoftNum > 100 Then SoftNum = 10
    If ImgWidth < 0 Or ImgWidth > 1000 Then ImgWidth = 150
    If ImgHeight < 0 Or ImgHeight > 1000 Then ImgHeight = 150
    If iTimeOut < 1000 Or iTimeOut > 100000 Then iTimeOut = 5000
    If effectID < 0 Or effectID > 23 Then effectID = 23

    FoundErr = False
    If iChannelID <> PrevChannelID Or ChannelID = 0 Then
        Call GetChannel(iChannelID)
    End If
    PrevChannelID = iChannelID
    If FoundErr = True Then
        GetSlidePicSoft = ErrMsg
        Exit Function
    End If

    sqlPic = "select top " & SoftNum & " S.ChannelID,S.ClassID,S.SoftID,S.SoftName,S.UpdateTime,S.SoftPicUrl"
    sqlPic = sqlPic & ",C.ClassName,C.ClassDir,C.ParentDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)

    Dim ranNum
    Randomize
    ranNum = Int(900 * Rnd) + 100
    strPic = "<script language=JavaScript>" & vbCrLf
    strPic = strPic & "<!--" & vbCrLf
    strPic = strPic & "var SlidePic_" & ranNum & " = new SlidePic_Soft(""SlidePic_" & ranNum & """);" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Width    = " & ImgWidth & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Height   = " & ImgHeight & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".TimeOut  = " & iTimeOut & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Effect   = " & effectID & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".TitleLen = " & TitleLen & ";" & vbCrLf

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    Do While Not rsPic.EOF
        If iChannelID = 0 Then
            If rsPic("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsPic("ChannelID"))
                PrevChannelID = rsPic("ChannelID")
            End If
        End If
        If LCase(Left(rsPic("SoftPicUrl"), Len("UploadSoftPic"))) = "uploadsoftpic" Then
            SoftPicUrl = ChannelUrl & "/" & rsPic("SoftPicUrl")
        Else
            SoftPicUrl = rsPic("SoftPicUrl")
        End If
        If TitleLen = -1 Then
            strTitle = rsPic("SoftName")
        Else
            strTitle = GetSubStr(rsPic("SoftName"), TitleLen, ShowSuspensionPoints)
        End If
        
        strPic = strPic & "var oSP = new objSP_Soft();" & vbCrLf
        strPic = strPic & "oSP.ImgUrl         = """ & SoftPicUrl & """;" & vbCrLf
        strPic = strPic & "oSP.LinkUrl        = """ & GetSoftUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("SoftID")) & """;" & vbCrLf
        strPic = strPic & "oSP.Title         = """ & strTitle & """;" & vbCrLf
        strPic = strPic & "SlidePic_" & ranNum & ".Add(oSP);" & vbCrLf
        
        rsPic.MoveNext
    Loop
    strPic = strPic & "SlidePic_" & ranNum & ".Show();" & vbCrLf
    strPic = strPic & "//-->" & vbCrLf
    strPic = strPic & "</script>" & vbCrLf
    
    rsPic.Close
    Set rsPic = Nothing
    GetSlidePicSoft = strPic
End Function

Private Function JS_SlidePic()
    Dim strJS, LinkTarget
    LinkTarget = XmlText_Class("SlidePicSoft/LinkTarget", "_blank")
    strJS = strJS & "<script language=""JavaScript"">" & vbCrLf
    strJS = strJS & "<!--" & vbCrLf
    strJS = strJS & "function objSP_Soft() {this.ImgUrl=""""; this.LinkUrl=""""; this.Title="""";}" & vbCrLf
    strJS = strJS & "function SlidePic_Soft(_id) {this.ID=_id; this.Width=0;this.Height=0; this.TimeOut=5000; this.Effect=23; this.TitleLen=0; this.PicNum=-1; this.Img=null; this.Url=null; this.Title=null; this.AllPic=new Array(); this.Add=SlidePic_Soft_Add; this.Show=SlidePic_Soft_Show; this.LoopShow=SlidePic_Soft_LoopShow;}" & vbCrLf
    strJS = strJS & "function SlidePic_Soft_Add(_SP) {this.AllPic[this.AllPic.length] = _SP;}" & vbCrLf
    strJS = strJS & "function SlidePic_Soft_Show() {" & vbCrLf
    strJS = strJS & "  if(this.AllPic[0] == null) return false;" & vbCrLf
    strJS = strJS & "  document.write(""<div align='center'><a id='Url_"" + this.ID + ""' href='' target='" & LinkTarget & "'><img id='Img_"" + this.ID + ""' style='width:"" + this.Width + ""px; height:"" + this.Height + ""px; filter: revealTrans(duration=2,transition=23);' src='javascript:null' border='0'></a>"");" & vbCrLf
    strJS = strJS & "  if(this.TitleLen != 0) {document.write(""<br><span id='Title_"" + this.ID + ""'></span></div>"");}" & vbCrLf
    strJS = strJS & "  else{document.write(""</div>"");}" & vbCrLf
    strJS = strJS & "  this.Img = document.getElementById(""Img_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Url = document.getElementById(""Url_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Title = document.getElementById(""Title_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.LoopShow();" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function SlidePic_Soft_LoopShow() {" & vbCrLf
    strJS = strJS & "  if(this.PicNum<this.AllPic.length-1) this.PicNum++ ; " & vbCrLf
    strJS = strJS & "  else this.PicNum=0; " & vbCrLf
    strJS = strJS & "  this.Img.filters.revealTrans.Transition=this.Effect; " & vbCrLf
    strJS = strJS & "  this.Img.filters.revealTrans.apply(); " & vbCrLf
    strJS = strJS & "  this.Img.src=this.AllPic[this.PicNum].ImgUrl;" & vbCrLf
    strJS = strJS & "  this.Img.filters.revealTrans.play();" & vbCrLf
    strJS = strJS & "  this.Url.href=this.AllPic[this.PicNum].LinkUrl;" & vbCrLf
    strJS = strJS & "  if(this.Title) this.Title.innerHTML=""<a href=""+this.AllPic[this.PicNum].LinkUrl+"" target='" & LinkTarget & "'>""+this.AllPic[this.PicNum].Title+""</a>"";" & vbCrLf
    strJS = strJS & "  this.Img.timer=setTimeout(this.ID+"".LoopShow()"",this.TimeOut);" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "//-->" & vbCrLf
    strJS = strJS & "</script>" & vbCrLf
    JS_SlidePic = strJS
End Function

Private Function GetSoftPicUrl(SoftPicUrl, SoftPicWidth, SoftPicHeight)
    Dim strSoftPicUrl, FileType, strPicUrl

    If SoftPicUrl = "" Then
        strSoftPicUrl = strSoftPicUrl & "<img src='" & strPicUrl & strInstallDir & "images/nopic.gif' "
        If SoftPicWidth > 0 Then strSoftPicUrl = strSoftPicUrl & " width='" & SoftPicWidth & "'"
        If SoftPicHeight > 0 Then strSoftPicUrl = strSoftPicUrl & " height='" & SoftPicHeight & "'"
        strSoftPicUrl = strSoftPicUrl & " border='0'>"
    Else
        FileType = LCase(Mid(SoftPicUrl, InStrRev(SoftPicUrl, ".") + 1))
        If LCase(Left(SoftPicUrl, Len("UploadSoftPic"))) = "uploadsoftpic" Then
            strPicUrl = ChannelUrl & "/" & SoftPicUrl
        Else
            strPicUrl = SoftPicUrl
        End If
        If FileType = "swf" Then
            strSoftPicUrl = strSoftPicUrl & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' "
            If SoftPicWidth > 0 Then strSoftPicUrl = strSoftPicUrl & " width='" & SoftPicWidth & "'"
            If SoftPicHeight > 0 Then strSoftPicUrl = strSoftPicUrl & " height='" & SoftPicHeight & "'"
            strSoftPicUrl = strSoftPicUrl & "><param name='movie' value='" & strPicUrl & "'><param name='quality' value='high'><embed src='" & strPicUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' "
            If SoftPicWidth > 0 Then strSoftPicUrl = strSoftPicUrl & " width='" & SoftPicWidth & "'"
            If SoftPicHeight > 0 Then strSoftPicUrl = strSoftPicUrl & " height='" & SoftPicHeight & "'"
            strSoftPicUrl = strSoftPicUrl & "></embed></object>"
        ElseIf FileType = "gif" Or FileType = "jpg" Or FileType = "jpeg" Or FileType = "jpe" Or FileType = "bmp" Or FileType = "png" Then
            strSoftPicUrl = strSoftPicUrl & "<img class='pic2' src='" & strPicUrl & "' "
            If SoftPicWidth > 0 Then strSoftPicUrl = strSoftPicUrl & " width='" & SoftPicWidth & "'"
            If SoftPicHeight > 0 Then strSoftPicUrl = strSoftPicUrl & " height='" & SoftPicHeight & "'"
            strSoftPicUrl = strSoftPicUrl & " border='0'>"
        Else
            strSoftPicUrl = strSoftPicUrl & "<img class='pic2' src='" & strInstallDir & "images/nopic.gif' "
            If SoftPicWidth > 0 Then strSoftPicUrl = strSoftPicUrl & " width='" & SoftPicWidth & "'"
            If SoftPicHeight > 0 Then strSoftPicUrl = strSoftPicUrl & " height='" & SoftPicHeight & "'"
            strSoftPicUrl = strSoftPicUrl & " border='0'>"
        End If
    End If
    GetSoftPicUrl = strSoftPicUrl
End Function

Private Function GetSearchResultIDArr(iChannelID)
    Dim sqlSearch, rsSearch
    Dim rsField
    Dim SoftNum, arrSoftID

    If PE_CLng(SearchResultNum) > 0 Then
        sqlSearch = "select top " & PE_CLng(SearchResultNum) & " SoftID "
    Else
        sqlSearch = "select SoftID "
    End If
    sqlSearch = sqlSearch & " from PE_Soft where Deleted=" & PE_False & " and Status=3"
    If iChannelID > 0 Then
        sqlSearch = sqlSearch & " and ChannelID=" & iChannelID & " "
    End If
    If ClassID > 0 Then
        If Child > 0 Then
            sqlSearch = sqlSearch & " and ClassID in (" & arrChildID & ")"
        Else
            sqlSearch = sqlSearch & " and ClassID=" & ClassID
        End If
    End If
    If SpecialID > 0 Then
        sqlSearch = sqlSearch & " and SoftID in (select ItemID from PE_InfoS where SpecialID=" & SpecialID & ")"
    End If
    If strField <> "" Then  '普通搜索
        Select Case strField
            Case "Title", "SoftName"
                sqlSearch = sqlSearch & SetSearchString("SoftName")
            Case "Content", "SoftIntro"
                sqlSearch = sqlSearch & SetSearchString("SoftIntro")
            Case "Author"
                sqlSearch = sqlSearch & SetSearchString("Author")
            Case "Inputer"
                sqlSearch = sqlSearch & SetSearchString("Inputer")
            Case "Keywords"
                sqlSearch = sqlSearch & SetSearchString("Keyword")
            Case Else  '自定义字段
                Set rsField = Conn.Execute("select Title from PE_Field where (ChannelID=-2 or ChannelID=" & iChannelID & ") and FieldName='" & ReplaceBadChar(strField) & "'")
                If rsField.BOF And rsField.EOF Then
                    sqlSearch = sqlSearch & SetSearchString("Title")
                Else
                    sqlSearch = sqlSearch & SetSearchString(ReplaceBadChar(strField))
                End If
                rsField.Close
                Set rsField = Nothing
        End Select
    Else   '高级搜索
        '定义高级搜索变量
        Dim SoftName, SoftIntro, Author, CopyFrom, Keyword2, LowInfoPoint, HighInfoPoint, BeginDate, EndDate, Inputer, SoftLanguage, SoftType, SoftVersion, CopyrightType
        SoftName = Trim(Request("SoftName"))
        SoftIntro = Trim(Request("SoftIntro"))
        Author = Trim(Request("Author"))
        CopyFrom = Trim(Request("CopyFrom"))
        Keyword2 = Trim(Request("Keywords"))
        LowInfoPoint = PE_CLng(Request("LowInfoPoint"))
        HighInfoPoint = PE_CLng(Request("HighInfoPoint"))
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        Inputer = Trim(Request("Inputer"))
        SoftLanguage = Trim(Request("SoftLanguage"))
        SoftType = Trim(Request("SoftType"))
        SoftVersion = Trim(Request("SoftVersion"))
        CopyrightType = Trim(Request("CopyrightType"))
        strFileName = "Search.asp?ModuleName=Soft&ClassID=" & ClassID & "&SpecialID=" & SpecialID
        If SoftName <> "" Then
            SoftName = ReplaceBadChar(SoftName)
            strFileName = strFileName & "&SoftName=" & SoftName
            sqlSearch = sqlSearch & " and SoftName like '%" & SoftName & "%' "
        End If
        If SoftIntro <> "" Then
            SoftIntro = ReplaceBadChar(SoftIntro)
            strFileName = strFileName & "&SoftIntro=" & SoftIntro
            sqlSearch = sqlSearch & " and SoftIntro like '%" & SoftIntro & "%'"
        End If
        If SoftLanguage <> "" Then
            SoftLanguage = ReplaceBadChar(SoftLanguage)
            strFileName = strFileName & "&SoftLanguage=" & SoftLanguage
            sqlSearch = sqlSearch & " and SoftLanguage like '%" & SoftLanguage & "%' "
        End If
        If SoftType <> "" Then
            SoftType = ReplaceBadChar(SoftType)
            strFileName = strFileName & "&SoftType=" & SoftType
            sqlSearch = sqlSearch & " and SoftType like '%" & SoftType & "%' "
        End If
        If SoftVersion <> "" Then
            SoftVersion = ReplaceBadChar(SoftVersion)
            strFileName = strFileName & "&SoftVersion=" & SoftVersion
            sqlSearch = sqlSearch & " and SoftVersion like '%" & SoftVersion & "%' "
        End If
        If CopyrightType <> "" Then
            CopyrightType = ReplaceBadChar(CopyrightType)
            strFileName = strFileName & "&CopyrightType=" & CopyrightType
            sqlSearch = sqlSearch & " and CopyrightType like '%" & CopyrightType & "%' "
        End If
        If Author <> "" Then
            Author = ReplaceBadChar(Author)
            strFileName = strFileName & "&Author=" & Author
            sqlSearch = sqlSearch & " and Author like '%" & Author & "%' "
        End If
        If CopyFrom <> "" Then
            CopyFrom = ReplaceBadChar(CopyFrom)
            strFileName = strFileName & "&CopyFrom=" & CopyFrom
            sqlSearch = sqlSearch & " and CopyFrom like '%" & CopyFrom & "%' "
        End If
        If Inputer <> "" Then
            Inputer = ReplaceBadChar(Inputer)
            strFileName = strFileName & "&Inputer=" & Inputer
            sqlSearch = sqlSearch & " and Inputer='" & Inputer & "' "
        End If
        If Keyword2 <> "" Then
            Keyword2 = ReplaceBadChar(Keyword2)
            strFileName = strFileName & "&Keywords=" & Keyword2
            sqlSearch = sqlSearch & " and Keyword like '%" & Keyword2 & "%' "
        End If
    
        If LowInfoPoint > 0 Then
            strFileName = strFileName & "&LowInfoPoint=" & LowInfoPoint
            sqlSearch = sqlSearch & " and InfoPoint >=" & LowInfoPoint
        End If
        If HighInfoPoint > 0 Then
            strFileName = strFileName & "&HighInfoPoint=" & HighInfoPoint
            sqlSearch = sqlSearch & " and InfoPoint <=" & HighInfoPoint
        End If

        If IsDate(BeginDate) Then
            strFileName = strFileName & "&BeginDate=" & BeginDate
            If SystemDatabaseType = "SQL" Then
                sqlSearch = sqlSearch & " and UpdateTime >= '" & BeginDate & "'"
            Else
                sqlSearch = sqlSearch & " and UpdateTime >= #" & BeginDate & "#"
            End If
        End If
        If IsDate(EndDate) Then
            strFileName = strFileName & "&EndDate=" & EndDate
            If SystemDatabaseType = "SQL" Then
                sqlSearch = sqlSearch & " and UpdateTime <= '" & EndDate & "'"
            Else
                sqlSearch = sqlSearch & " and UpdateTime <= #" & EndDate & "#"
            End If
        End If

        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-2 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            If Trim(Request(rsField("FieldName"))) <> "" Then
                strFileName = strFileName & "&" & Trim(rsField("FieldName")) & "=" & ReplaceBadChar(Trim(Request(rsField("FieldName"))))
                sqlSearch = sqlSearch & " and " & Trim(rsField("FieldName")) & " like '%" & ReplaceBadChar(Trim(Request(rsField("FieldName")))) & "%' "
            End If
            rsField.MoveNext
        Loop
        Set rsField = Nothing
        
    End If
    sqlSearch = sqlSearch & " order by SoftID desc"
    arrSoftID = ""
    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1
    If rsSearch.BOF And rsSearch.EOF Then
        totalPut = 0
    Else
        totalPut = rsSearch.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsSearch.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        SoftNum = 0
        Do While Not rsSearch.EOF
            If arrSoftID = "" Then
                arrSoftID = rsSearch(0)
            Else
                arrSoftID = arrSoftID & "," & rsSearch(0)
            End If
            SoftNum = SoftNum + 1
            If SoftNum >= MaxPerPage Then Exit Do
            rsSearch.MoveNext
        Loop
    End If
    rsSearch.Close
    Set rsSearch = Nothing

    GetSearchResultIDArr = arrSoftID
End Function

'=================================================
'函数名：GetSearchResult
'作  用：分页显示搜索结果
'参  数：无
'=================================================
Private Function GetSearchResult(iChannelID)
    Dim sqlSearch, rsSearch, iCount, SoftNum, arrSoftID, strSearchResult, Content
    strSearchResult = ""
    arrSoftID = GetSearchResultIDArr(iChannelID)
    If arrSoftID = "" Then
        GetSearchResult = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "没有或没有找到任何{$ChannelShortName}") & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If

    SoftNum = 1
    sqlSearch = "select S.ChannelID,S.SoftID,S.SoftName,S.SoftVersion,S.Author,S.UpdateTime,S.Hits,S.SoftIntro,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where SoftID in (" & arrSoftID & ") order by SoftID desc"
    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1
    Do While Not rsSearch.EOF
        If iChannelID = 0 Then
            If rsSearch("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsSearch("ChannelID"))
                PrevChannelID = rsSearch("ChannelID")
            End If
        End If
        
        strSearchResult = strSearchResult & "<b>" & CStr(MaxPerPage * (CurrentPage - 1) + SoftNum) & ".</b> "
        
        strSearchResult = strSearchResult & "[<a class='LinkSearchResult' href='" & GetClassUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("ClassID"), rsSearch("ClassPurview")) & "' target='_blank'>" & rsSearch("ClassName") & "</a>] "
        
        strSearchResult = strSearchResult & "<a class='LinkSearchResult' href='" & GetSoftUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("UpdateTime"), rsSearch("SoftID")) & "' target='_blank'>"
        
        If strField = "SoftName" Then
            strSearchResult = strSearchResult & "<b>" & Replace(ReplaceText(rsSearch("SoftName"), 2) & "", "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "</b>"
        Else
            strSearchResult = strSearchResult & "<b>" & ReplaceText(rsSearch("SoftName"), 2) & "</b>"
        End If
        strSearchResult = strSearchResult & " " & rsSearch("SoftVersion")
        strSearchResult = strSearchResult & "</a>"
        If strField = "Author" Then
            strSearchResult = strSearchResult & "&nbsp;[" & Replace(rsSearch("Author") & "", "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "]"
        Else
            strSearchResult = strSearchResult & "&nbsp;[" & rsSearch("Author") & "]"
        End If
        strSearchResult = strSearchResult & "[" & FormatDateTime(rsSearch("UpdateTime"), 1) & "][" & rsSearch("Hits") & "]"
        strSearchResult = strSearchResult & "<br>"
        
        Content = Left(Replace(Replace(ReplaceText(nohtml(rsSearch("SoftIntro")), 1), ">", "&gt;"), "<", "&lt;"), SearchResult_ContentLenth)
        If strField = "Content" Then
            strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & Replace(Content, "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "……</div>"
        Else
            strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & Content & "……</div>"
        End If
        strSearchResult = strSearchResult & "<br>"
        SoftNum = SoftNum + 1
        rsSearch.MoveNext
    Loop
    rsSearch.Close
    Set rsSearch = Nothing
    GetSearchResult = strSearchResult
End Function

Public Function GetSearchResult2(iChannelID, strValue)   '得到自定义列表的版面设计的HTML代码
    Dim strCustom, strParameter
	strCustom = strValue
    regEx.Pattern = "【SearchResultList\((.*?)\)】([\s\S]*?)【\/SearchResultList】"
    Set Matches = regEx.Execute(strCustom)
    For Each Match In Matches
        strParameter = Replace(Match.SubMatches(0), Chr(34), " ")
        strCustom = PE_Replace(strCustom, Match.value, GetSearchResultLabel(strParameter, Match.SubMatches(1), iChannelID))
    Next
    GetSearchResult2 = strCustom
End Function

Private Function GetSearchResultLabel(strTemp, strList, iChannelID)
    Dim sqlSearch, rsSearch, iCount, SoftNum, arrSoftID, Content
    Dim arrTemp
    Dim strSoftPic, strPicTemp, arrPicTemp
    Dim ItemNum, arrClassID, IsHot, IsElite, Author, DateNum, OrderType, UsePage, OpenType, TitleLen, ContentLen
    Dim iCols, iColsHtml, iRows, iRowsHtml, iNumber
    Dim rsCustom, strCustomList, strLink
    Dim rsField, ArrField, iField
    
    iCount = 0
    strCustomList = ""
        
    If strTemp = "" Or strList = "" Then GetSearchResultLabel = "": Exit Function

    iCols = 1: iRows = 1: iColsHtml = "": iRowsHtml = ""
    regEx.Pattern = "【(Cols|Rows)=(\d{1,2})\s*(?:\||｜)(.+?)】"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        If LCase(Match.SubMatches(0)) = "cols" Then
            If Match.SubMatches(1) > 1 Then iCols = Match.SubMatches(1)
            iColsHtml = Match.SubMatches(2)
        ElseIf LCase(Match.SubMatches(0)) = "rows" Then
            If Match.SubMatches(1) > 1 Then iRows = Match.SubMatches(1)
            iRowsHtml = Match.SubMatches(2)
        End If
        strList = regEx.Replace(strList, "")
    Next
    
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) <> 2 Then
        GetSearchResultLabel = "自定义列表标签：【SearchResultList(参数列表)】列表内容【/SearchResultList】的参数个数不对。请检查模板中的此标签。"
        Exit Function
    End If
    
    TitleLen = arrTemp(0)
    UsePage = arrTemp(1)
    ContentLen = arrTemp(2)
    
    arrSoftID = GetSearchResultIDArr(iChannelID)
    If arrSoftID = "" Then
        GetSearchResultLabel = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "没有或没有找到任何{$ChannelShortName}") & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If
    
    Set rsField = Conn.Execute("select FieldName,LabelName from PE_Field where ChannelID=-2 or ChannelID=" & ChannelID & "")
    If Not (rsField.BOF And rsField.EOF) Then
        ArrField = rsField.getrows(-1)
    End If
    Set rsField = Nothing

    sqlSearch = "select S.ChannelID,S.SoftID,S.SoftName,S.SoftVersion,S.Author,S.SoftPicUrl,"
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlSearch = sqlSearch & "S." & ArrField(0, iField) & ","
        Next
    End If
    sqlSearch = sqlSearch & "S.UpdateTime,S.DemoUrl,S.RegUrl,S.Hits,S.DayHits,S.InfoPoint,S.WeekHits,"
    sqlSearch = sqlSearch & "S.MonthHits,S.SoftLanguage,S.SoftType,S.SoftIntro,S.OperatingSystem,S.OnTop,S.Keyword,"
    sqlSearch = sqlSearch & "S.Elite,S.Stars,S.SoftSize,S.CopyrightType,S.DownloadUrl,C.ClassID,C.ClassName,C.ParentDir,"
    sqlSearch = sqlSearch & "C.ClassDir,C.ClassPurview,C.ReadMe from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where SoftID in (" & arrSoftID & ") order by SoftID desc"
    SoftNum = 1
    Set rsCustom = Server.CreateObject("ADODB.Recordset")
    rsCustom.Open sqlSearch, Conn, 1, 1
    Do While Not rsCustom.EOF
        If iChannelID = 0 Then
            If rsCustom("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsCustom("ChannelID"))
                PrevChannelID = rsCustom("ChannelID")
            End If
        End If
        
        strTemp = strList

        iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1

        strTemp = PE_Replace(strTemp, "{$Number}", iNumber)
        strTemp = PE_Replace(strTemp, "{$ClassID}", rsCustom("ClassID"))
        strTemp = PE_Replace(strTemp, "{$ClassName}", rsCustom("ClassName"))
        strTemp = PE_Replace(strTemp, "{$ParentDir}", rsCustom("ParentDir"))
        strTemp = PE_Replace(strTemp, "{$ClassDir}", rsCustom("ClassDir"))
        strTemp = PE_Replace(strTemp, "{$Readme}", rsCustom("ReadMe"))
        If InStr(strTemp, "{$ClassUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ClassUrl}", GetClassUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("ClassID"), rsCustom("ClassPurview")))

        strLink = "<a href='" & GetSoftUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("SoftID")) & "'>"
    
        If InStr(strTemp, "{$SoftUrl}") > 0 Then strTemp = Replace(strTemp, "{$SoftUrl}", GetSoftUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("SoftID")))
        strTemp = PE_Replace(strTemp, "{$SoftID}", rsCustom("SoftID"))
        If TitleLen > 0 Then
            strTemp = Replace(strTemp, "{$SoftName}", Left(rsCustom("SoftName"), TitleLen))
        Else
            strTemp = Replace(strTemp, "{$SoftName}", rsCustom("SoftName"))
        End If
        strTemp = Replace(strTemp, "{$SoftNameOriginal}", rsCustom("SoftName"))
        strTemp = PE_Replace(strTemp, "{$SoftVersion}", strLink & rsCustom("SoftVersion") & "</a>")
        If InStr(strTemp, "{$SoftProperty}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftProperty}", GetPropertyPic(rsCustom("OnTop"), rsCustom("Elite")))
        strTemp = PE_Replace(strTemp, "{$SoftSize}", rsCustom("SoftSize"))
        If InStr(strTemp, "{$SoftSize_M}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftSize_M}", Round(rsCustom("SoftSize") / 1024, 2))
        strTemp = PE_Replace(strTemp, "{$Keyword}", GetKeywords(",", rsCustom("Keyword")))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        If InStr(strTemp, "{$CopyrightType}") > 0 Then strTemp = PE_Replace(strTemp, "{$CopyrightType}", PE_HTMLEncode(rsCustom("CopyrightType")))
        If InStr(strTemp, "{$Stars}") > 0 Then strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        If ContentLen > 0 Then
            If InStr(strTemp, "{$SoftIntro}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftIntro}", Left(nohtml(rsCustom("SoftIntro")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$SoftIntro}", "")
        End If
        strTemp = PE_Replace(strTemp, "{$OperatingSystem}", rsCustom("OperatingSystem"))
        If InStr(strTemp, "{$SoftType}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftType}", PE_HTMLEncode(rsCustom("SoftType")))
        If InStr(strTemp, "{$SoftLanguage}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftLanguage}", PE_HTMLEncode(rsCustom("SoftLanguage")))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$DayHits}", rsCustom("DayHits"))
        strTemp = PE_Replace(strTemp, "{$WeekHits}", rsCustom("WeekHits"))
        strTemp = PE_Replace(strTemp, "{$MonthHits}", rsCustom("MonthHits"))
        strTemp = PE_Replace(strTemp, "{$Author}", rsCustom("Author"))
        If InStr(strTemp, "{$InfoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$InfoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$SoftPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        strTemp = PE_Replace(strTemp, "{$SoftAuthor}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$AuthorEmail}", "")
        strTemp = PE_Replace(strTemp, "{$DemoUrl}", rsCustom("DemoUrl"))
        strTemp = PE_Replace(strTemp, "{$RegUrl}", rsCustom("RegUrl"))
        
        'strTemp = PE_Replace(strTemp, "{$DownloadUrl}", Mid(rsCustom("DownloadUrl"), InStr(rsCustom("DownloadUrl"),"|")))
        '替换下载图片
        regEx.Pattern = "\{\$SoftPic\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strSoftPic = GetSoftPicUrl(Trim(rsCustom("SoftPicUrl")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)))
            strTemp = Replace(strTemp, Match.value, strSoftPic)
        Next
        
        If IsArray(ArrField) Then
            For iField = 0 To UBound(ArrField, 2)
                strTemp = PE_Replace(strTemp, ArrField(1, iField), PE_HTMLEncode(rsCustom(Trim(ArrField(0, iField)))))
            Next
        End If

        strCustomList = strCustomList & strTemp
        rsCustom.MoveNext
        iCount = iCount + 1
        If iCols > 1 And iCount Mod iCols = 0 Then strCustomList = strCustomList & iColsHtml
        If iRows > 1 And iCount Mod iCols * iRows = 0 Then strCustomList = strCustomList & iRowsHtml
        If iCount >= MaxPerPage Then Exit Do
    Loop
    rsCustom.Close
    Set rsCustom = Nothing
    
    GetSearchResultLabel = strCustomList
End Function

'=================================================
'函数名：GetCorrelative
'作  用：显示相关下载
'参  数：SoftNum  ----最多显示多少个下载
'        TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'        OrderType ---- 排序方式，1--按软件ID降序，2--按软件ID升序，3--按更新时间降序，4--按更新时间升序，5--按点击数降序，6--按点击数升序，7--按评论数降序，8--按评论数升序
'        OpenType ---- 软件打开方式，0为在原窗口打开，1为在新窗口打开
'        Cols ---- 每行的列数。超过此列数就换行。
'=================================================
Private Function GetCorrelative(SoftNum, TitleLen, OrderType, OpenType, Cols)
    Dim rsCorrelative, sqlCorrelative, strCorrelative, strSoftTitle, iCols, iTemp
    Dim strKey, arrKey, i, MaxNum
    iTemp = 1
    If PE_CLng(Cols) <> 0 Then
        iCols = PE_CLng(Cols)
    Else
        iCols = 1
    End If

    If SoftNum > 0 And SoftNum <= 100 Then
        sqlCorrelative = "select top " & SoftNum
    Else
        sqlCorrelative = "Select Top 5 "
    End If
    strKey = Mid(rsSoft("Keyword"), 2, Len(rsSoft("Keyword")) - 2)
    If InStr(strKey, "|") > 1 Then
        arrKey = Split(strKey, "|")
        MaxNum = UBound(arrKey)
        If MaxNum > 2 Then MaxNum = 2
        strKey = "((S.Keyword like '%|" & arrKey(0) & "|%')"
        For i = 1 To MaxNum
            strKey = strKey & " or (S.Keyword like '%|" & arrKey(i) & "|%')"
        Next
        strKey = strKey & ")"
    Else
        strKey = "(S.Keyword like '%|" & strKey & "|%')"
    End If
    sqlCorrelative = sqlCorrelative & " S.SoftID,S.SoftName,S.SoftVersion,S.Author,S.UpdateTime,S.Hits,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where S.ChannelID=" & ChannelID & " and S.Deleted=" & PE_False & " and S.Status=3"

    sqlCorrelative = sqlCorrelative & " and " & strKey & " and S.SoftID<>" & SoftID & " Order by "
    Select Case PE_CLng(OrderType)
    Case 1
        sqlCorrelative = sqlCorrelative & "S.SoftID desc"
    Case 2
        sqlCorrelative = sqlCorrelative & "S.SoftID asc"
    Case 3
        sqlCorrelative = sqlCorrelative & "S.UpdateTime desc"
    Case 4
        sqlCorrelative = sqlCorrelative & "S.UpdateTime asc"
    Case 5
        sqlCorrelative = sqlCorrelative & "S.Hits desc"
    Case 6
        sqlCorrelative = sqlCorrelative & "S.Hits asc"
    Case 7
        sqlCorrelative = sqlCorrelative & "S.CommentCount desc"
    Case 8
        sqlCorrelative = sqlCorrelative & "S.CommentCount asc"
    Case Else
        sqlCorrelative = sqlCorrelative & "S.SoftID desc"
    End Select

    Set rsCorrelative = Conn.Execute(sqlCorrelative)
    If TitleLen < 0 Or TitleLen > 255 Then TitleLen = 50
    If rsCorrelative.BOF And rsCorrelative.EOF Then
        strCorrelative = R_XmlText_Class("ShowSoft/NoCorrelative", "没有相关{$ChannelShortName}")
    Else
        Do While Not rsCorrelative.EOF
            strSoftTitle = rsCorrelative("SoftName")
            If Trim(rsCorrelative("SoftVersion")) <> "" Then
                strSoftTitle = strSoftTitle & " " & Trim(rsCorrelative("SoftVersion"))
            End If
            strSoftTitle = GetSubStr(strSoftTitle, TitleLen, ShowSuspensionPoints)
            strCorrelative = strCorrelative & "<a class='LinkSoftCorrelative' href='" & GetSoftUrl(rsCorrelative("ParentDir"), rsCorrelative("ClassDir"), rsCorrelative("UpdateTime"), rsCorrelative("SoftID")) & "'"
            If Trim(rsCorrelative("SoftVersion")) <> "" Then
                strCorrelative = strCorrelative & " title='" & Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("SoftList/Title2", "{$ChannelShortName}名称：{$SoftName}{$br}{$ChannelShortName}版本：{$SoftVersion}{$br}作&nbsp;&nbsp;&nbsp;&nbsp;者：{$Author}{$br}更新时间：{$UpdateTime}"), "{$SoftName}", rsCorrelative("SoftName")), "{$SoftVersion}", rsCorrelative("SoftVersion")), "{$Author}", rsCorrelative("Author")), "{$UpdateTime}", rsCorrelative("UpdateTime")), "{$br}", vbCrLf)
            Else
                strCorrelative = strCorrelative & " title='" & Replace(Replace(Replace(Replace(R_XmlText_Class("SoftList/Title", "{$ChannelShortName}名称：{$SoftName}{$br}作&nbsp;&nbsp;&nbsp;&nbsp;者：{$Author}{$br}更新时间：{$UpdateTime}"), "{$SoftName}", rsCorrelative("SoftName")), "{$Author}", rsCorrelative("Author")), "{$UpdateTime}", rsCorrelative("UpdateTime")), "{$br}", vbCrLf)
            End If
            If OpenType = 0 Then
                strCorrelative = strCorrelative & "' target=""_self"">"
            Else
                strCorrelative = strCorrelative & "' target=""_blank"">"
            End If
            strCorrelative = strCorrelative & strSoftTitle & "</a>"
            If (iTemp Mod iCols) = 0 Then
                strCorrelative = strCorrelative & "<br>"
            Else
                strCorrelative = strCorrelative & "&nbsp;&nbsp;"
            End If
            rsCorrelative.MoveNext
            iTemp = iTemp + 1
        Loop
    End If
    rsCorrelative.Close
    Set rsCorrelative = Nothing
    GetCorrelative = strCorrelative
End Function


Private Function GetHits()
    Dim strHits
    If UseCreateHTML > 0 Then
        strHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?HitsType=0&SoftID=" & SoftID & "'></script>"
    Else
        strHits = rsSoft("Hits")
    End If
    GetHits = strHits
End Function

Private Function GetDayHits()
    Dim strHits
    If UseCreateHTML > 0 Then
        strHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?HitsType=1&SoftID=" & SoftID & "'></script>"
    Else
        strHits = rsSoft("DayHits")
    End If
    GetDayHits = strHits
End Function

Private Function GetWeekHits()
    Dim strHits
    If UseCreateHTML > 0 Then
        strHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?HitsType=2&SoftID=" & SoftID & "'></script>"
    Else
        strHits = rsSoft("WeekHits")
    End If
    GetWeekHits = strHits
End Function

Private Function GetMonthHits()
    Dim strHits
    If UseCreateHTML > 0 Then
        strHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?HitsType=3&SoftID=" & SoftID & "'></script>"
    Else
        strHits = rsSoft("MonthHits")
    End If
    GetMonthHits = strHits
End Function

Private Function GetSoftLink()
    GetSoftLink = Replace(Replace(R_XmlText_Class("ShowSoft/SoftLink", "<a href='{$DemoUrl}' target='_blank'>演示地址</a>&nbsp;&nbsp;<a href='{$RegUrl}' target='_blank'>注册地址</a>"), "{$DemoUrl}", rsSoft("DemoUrl")), "{$RegUrl}", rsSoft("RegUrl"))
End Function

Private Function GetSoftProperty()
    Dim strProperty
    If rsSoft("OnTop") = True Then
        strProperty = strProperty & XmlText_Class("ShowSoft/OnTop", "<font color=blue>顶</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsSoft("Hits") >= HitsOfHot Then
        strProperty = strProperty & XmlText_Class("ShowSoft/Hot", "<font color=red>热</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsSoft("Elite") = True Then
        strProperty = strProperty & XmlText_Class("ShowSoft/Elite", "<font color=green>荐</font>")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;"
    End If
    GetSoftProperty = strProperty
End Function

Private Function GetStars(Stars)
    GetStars = "<font color='" & XmlText_Class("ShowSoft/Star_Color", "#009900") & "'>" & String(Stars, XmlText_Class("ShowSoft/Star", "★")) & "</font>"
End Function

Private Function GetDownloadUrlList(ShowModule)
    Dim DownloadUrls, arrDownloadUrls, arrUrls, iTemp, strUrls
    Dim rsDownServer, sqlDownServer, ShowServerName, iShowModule, iCols
    iShowModule = PE_CLng(ShowModule)
    DownloadUrls = rsSoft("DownloadUrl")
    If DownloadUrls = "" Then
        GetDownloadUrlList = ""
        Exit Function
    End If
    strUrls = ""
    If InStr(DownloadUrls, "@@@") > 0 Then
    '构建镜像服务器下载地址列表
        arrDownloadUrls = Trim(Replace(DownloadUrls, "@@@", "")) '得PE_Soft中的Url地址
        sqlDownServer = "select * from PE_DownServer where ChannelID=" & ChannelID
        Set rsDownServer = Server.CreateObject("adodb.recordset")
        rsDownServer.Open sqlDownServer, Conn, 1, 3
        If rsDownServer.BOF Or rsDownServer.EOF Then
           strUrls = "对不起，未找到任何镜像服务器信息！"
        End If
        iCols = 1

        Do While Not rsDownServer.EOF
            If rsDownServer("ShowType") = 0 Then
               ShowServerName = rsDownServer("ServerName")
            Else
               ShowServerName = "<img src='" & rsDownServer("ServerLogo") & "' border=0>"
            End If
            '设置了下载点数的处理（PE_DownServer中有收费相关字段）
            'If EnableProtect = True Or ClassPurview > 0 Or rsSoft("InfoPoint") > 0 Or rsDownServer("InfoPoint") > 0 Then
                strUrls = strUrls & "<a href='" & ChannelUrl_ASPFile & "/ShowSoftDown.asp?UrlID=" & rsDownServer("ServerID") & "&SoftID=" & rsSoft("SoftID") & "' target='_blank'>" & ShowServerName & "</a>"
            'Else
            '    strUrls = strUrls & "<a href='" & rsDownServer("ServerUrl") & arrDownloadUrls & "' target='_blank'>" & ShowServerName & "</a>"
            'End If
            If iShowModule = 0 Then
                strUrls = strUrls & "&nbsp;&nbsp;"
            Else
                If (iCols Mod iShowModule) <> 0 Then
                    strUrls = strUrls & "&nbsp;&nbsp;"
                Else
                    strUrls = strUrls & "<br>"
                End If
            End If
            iCols = iCols + 1
            rsDownServer.MoveNext
        Loop
        GetDownloadUrlList = strUrls
        rsDownServer.Close
        Set rsDownServer = Nothing
    Else
        iCols = 0
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        For iTemp = 0 To UBound(arrDownloadUrls)
            iCols = iCols + 1
            arrUrls = Split(arrDownloadUrls(iTemp), "|")
            If UBound(arrUrls) >= 1 Then
                If arrUrls(1) <> "" And arrUrls(1) <> "http://" Then
                    'If EnableProtect = True Or ClassPurview > 0 Or rsSoft("InfoPoint") > 0 Then
                        strUrls = strUrls & "<a href='" & ChannelUrl_ASPFile & "/ShowSoftDown.asp?UrlID=" & iTemp + 1 & "&SoftID=" & rsSoft("SoftID") & "' target='_blank'>" & arrUrls(0) & "</a>"
                    'Else
                    '    If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                    '        strUrls = strUrls & "<a href='" & ChannelUrl & "/" & UploadDir & "/" & arrUrls(1) & "' target='_blank'>" & arrUrls(0) & "</a>"
                    '    Else
                    '        strUrls = strUrls & "<a href='" & GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|") & "' target='_blank'>" & arrUrls(0) & "</a>"
                    '    End If
                    'End If
                    If iShowModule = 0 Then
                        strUrls = strUrls & "&nbsp;&nbsp;"
                    Else
                        If (iCols Mod iShowModule) <> 0 Then
                            strUrls = strUrls & "&nbsp;&nbsp;"
                        Else
                            strUrls = strUrls & "<br>"
                        End If
                    End If
                End If
            End If
        Next
        GetDownloadUrlList = strUrls
    End If
End Function

Private Function GetDownloadUrlList1(ShowModule,TempDownloadUrls,tempSoftID)
    
    Dim DownloadUrls, arrDownloadUrls, arrUrls, iTemp, strUrls
    Dim rsDownServer, sqlDownServer, ShowServerName, iShowModule, iCols
    iShowModule = PE_CLng(ShowModule)
    DownloadUrls = TempDownloadUrls
    If DownloadUrls = "" Then
        GetDownloadUrlList1 = ""
        Exit Function
    End If
    strUrls = ""
    If InStr(DownloadUrls, "@@@") > 0 Then
    '构建镜像服务器下载地址列表
        arrDownloadUrls = Trim(Replace(DownloadUrls, "@@@", "")) '得PE_Soft中的Url地址
        sqlDownServer = "select * from PE_DownServer where ChannelID=" & ChannelID
        Set rsDownServer = Server.CreateObject("adodb.recordset")
        rsDownServer.Open sqlDownServer, Conn, 1, 3
        If rsDownServer.BOF Or rsDownServer.EOF Then
           strUrls = "对不起，未找到任何镜像服务器信息！"
        End If
        iCols = 1

        Do While Not rsDownServer.EOF
            If rsDownServer("ShowType") = 0 Then
               ShowServerName = rsDownServer("ServerName")
            Else
               ShowServerName = "<img src='" & rsDownServer("ServerLogo") & "' border=0>"
            End If
            '设置了下载点数的处理（PE_DownServer中有收费相关字段）
            'If EnableProtect = True Or ClassPurview > 0 Or rsSoft("InfoPoint") > 0 Or rsDownServer("InfoPoint") > 0 Then
                strUrls = strUrls & "<a href='" & ChannelUrl_ASPFile & "/ShowSoftDown.asp?UrlID=" & rsDownServer("ServerID") & "&SoftID=" & tempSoftID & "' target='_blank'>" & ShowServerName & "</a>"
            'Else
            '    strUrls = strUrls & "<a href='" & rsDownServer("ServerUrl") & arrDownloadUrls & "' target='_blank'>" & ShowServerName & "</a>"
            'End If
            If iShowModule = 0 Then
                strUrls = strUrls & "&nbsp;&nbsp;"
            Else
                If (iCols Mod iShowModule) <> 0 Then
                    strUrls = strUrls & "&nbsp;&nbsp;"
                Else
                    strUrls = strUrls & "<br>"
                End If
            End If
            iCols = iCols + 1
            rsDownServer.MoveNext
        Loop
        GetDownloadUrlList = strUrls
        rsDownServer.Close
        Set rsDownServer = Nothing
    Else
        iCols = 0
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        For iTemp = 0 To UBound(arrDownloadUrls)
            iCols = iCols + 1
            arrUrls = Split(arrDownloadUrls(iTemp), "|")
            If UBound(arrUrls) >= 1 Then
                If arrUrls(1) <> "" And arrUrls(1) <> "http://" Then
                    'If EnableProtect = True Or ClassPurview > 0 Or rsSoft("InfoPoint") > 0 Then
                        strUrls = strUrls & "<a href='" & ChannelUrl_ASPFile & "/ShowSoftDown.asp?UrlID=" & iTemp + 1 & "&SoftID=" & tempSoftID & "' target='_blank'>" & arrUrls(0) & "</a>"
                    'Else
                    '    If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                    '        strUrls = strUrls & "<a href='" & ChannelUrl & "/" & UploadDir & "/" & arrUrls(1) & "' target='_blank'>" & arrUrls(0) & "</a>"
                    '    Else
                    '        strUrls = strUrls & "<a href='" & GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|") & "' target='_blank'>" & arrUrls(0) & "</a>"
                    '    End If
                    'End If
                    If iShowModule = 0 Then
                        strUrls = strUrls & "&nbsp;&nbsp;"
                    Else
                        If (iCols Mod iShowModule) <> 0 Then
                            strUrls = strUrls & "&nbsp;&nbsp;"
                        Else
                            strUrls = strUrls & "<br>"
                        End If
                    End If
                End If
            End If
        Next
        GetDownloadUrlList1 = strUrls
    End If
End Function

Public Function GetDownloadUrl()
    GetDownloadUrl = "ErrorDownloadUrl"
    Call Init
    Dim UrlID, SoftID, ComeUrl, cUrl, ConsumePoint
    Dim rsDownServer, sqlDownServer

    SoftID = PE_CLng(Trim(Request("SoftID")))
    UrlID = PE_CLng(Trim(Request("UrlID")))
    If SoftID = 0 Then
        FoundErr = True
        Response.Write "<br><li>请指定" & ChannelShortName & "ID！</li>"
        Exit Function
    End If
    If UrlID <= 0 Then UrlID = 1

    If EnableProtect = True Then
        ComeUrl = Replace(LCase(Trim(Request.ServerVariables("HTTP_REFERER"))), "http://", "")
        cUrl = LCase(Trim(Request.ServerVariables("SERVER_NAME")))
        If ComeUrl <> "" And ChannelUrl_ASPFile = ChannelUrl Then
            If Left(ComeUrl, Len(cUrl)) <> cUrl Then
                FoundErr = True
                Response.Write "<br><li>请勿非法盗链或下载本站软件！</li>"
                Exit Function
            End If
        End If
    End If
    
    Dim sqlSoft, rsSoft, LastHitTime
    sqlSoft = "select S.SoftID,S.SoftName,S.InfoPoint,S.DividePercent,S.Inputer,S.InfoPurview,S.DownloadUrl,S.LastHitTime,S.ChargeType,S.PitchTime,S.ReadTimes,S.arrGroupID,C.ClassID,C.ClassPurview,C.ParentID,C.ParentPath from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where S.Deleted=" & PE_False & " and S.Status=3 and S.SoftID=" & SoftID & " and S.ChannelID=" & ChannelID & ""
    Set rsSoft = Server.CreateObject("ADODB.Recordset")
    rsSoft.Open sqlSoft, Conn, 1, 3
    If rsSoft.BOF And rsSoft.BOF Then
        FoundErr = True
        Response.Write "<br><li>找不到指定的" & ChannelShortName & "！</li>"
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Function
    End If

    LastHitTime = rsSoft("LastHitTime")

    Dim DownloadUrl, DownloadUrls, arrDownloadUrls, arrUrls, iTemp
    Dim InfoPurview, InfoPoint, arrGroupID, ChargeType, PitchTime, ReadTimes, DividePercent

    DownloadUrls = rsSoft("DownloadUrl")
    iTemp = UrlID - 1
    InfoPurview = PE_CLng(rsSoft("InfoPurview"))
    InfoPoint = rsSoft("InfoPoint")
    arrGroupID = rsSoft("arrGroupID")
    ChargeType = PE_CLng(rsSoft("ChargeType"))
    PitchTime = PE_CLng(rsSoft("PitchTime"))
    ReadTimes = PE_CLng(rsSoft("ReadTimes"))
    DividePercent = PE_CDbl(rsSoft("DividePercent"))
    If InStr(DownloadUrls, "@@@") > 0 Then
    '构建镜像服务器下载地址
        arrDownloadUrls = Trim(Replace(DownloadUrls, "@@@", "")) '得PE_Soft中的Url地址
        sqlDownServer = "select * from PE_DownServer where ServerID= " & UrlID & " and ChannelID=" & ChannelID
        Set rsDownServer = Server.CreateObject("adodb.recordset")
        rsDownServer.Open sqlDownServer, Conn, 1, 3
        If rsDownServer.BOF Or rsDownServer.EOF Then
            FoundErr = True
            Response.Write "对不起，未找到任何镜像服务器信息或已经被删除！"
            Exit Function
        Else
            DownloadUrl = rsDownServer("ServerUrl") & arrDownloadUrls
            If InfoPurview = 0 And InfoPoint <= 0 Then
                InfoPurview = PE_CLng(rsDownServer("InfoPurview"))
                InfoPoint = rsDownServer("InfoPoint")
                arrGroupID = rsDownServer("arrGroupID")
                ChargeType = PE_CLng(rsDownServer("ChargeType"))
                PitchTime = PE_CLng(rsDownServer("PitchTime"))
                ReadTimes = PE_CLng(rsDownServer("ReadTimes"))
                DividePercent = PE_CDbl(rsDownServer("DividePercent"))
            End If
        End If
        rsDownServer.Close
        Set rsDownServer = Nothing
    Else
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        If UBound(arrDownloadUrls) >= iTemp Then
            arrUrls = Split(arrDownloadUrls(iTemp), "|")
            If UBound(arrUrls) >= 1 Then
                If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                    DownloadUrl = ChannelUrl & "/" & UploadDir & "/" & arrUrls(1)
                Else
                    DownloadUrl = GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|")
                End If
            End If
        End If
    End If
    
    If DownloadUrl = "" Or DownloadUrl = "http://" Then
        FoundErr = True
        Response.Write "<br><li>找不到有效下载地址！</li>"
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Function
    End If

    Dim ClassPurview, PurviewChecked, ParentID, ParentPath
    ClassPurview = PE_CLng(rsSoft("ClassPurview"))
    ParentID = PE_CLng(rsSoft("ParentID"))
    ParentPath = rsSoft("ParentPath")


    If ClassPurview > 0 Or InfoPurview > 0 Or InfoPoint > 0 Then
        Dim ErrMsg_NoLogin, ErrMsg_PurviewCheckedErr, ErrMsg_PurviewCheckedErr2, ErrMsg_NoMail, ErrMsg_NoCheck, ErrMsg_NeedPoint, ErrMsg_UsePoint, ErrMsg_OutTime, ErrMsg_Overflow_Total, ErrMsg_Overflow_Today
        ErrMsg_NoLogin = Replace(Replace(Replace(R_XmlText_Class("SoftContent/Nologin", "<br>&nbsp;&nbsp;&nbsp;&nbsp;你还没注册？或者没有登录？这{$ItemUnit}要求至少是本站的注册会员才能下载！<br><br>&nbsp;&nbsp;&nbsp;&nbsp;如果你还没注册，请赶紧<a href='{$InstallDir}Reg/User_Reg.asp'><font color=red>点此注册</font></a>吧！<br><br>&nbsp;&nbsp;&nbsp;&nbsp;如果你已经注册但还没登录，请赶紧<a href='{$InstallDir}User/User_Login.asp'><font color=red>点此登录</font></a>吧！<br><br>"), "{$ItemUnit}", ChannelItemUnit & ChannelShortName), "{$ChannelItemUnit}", ChannelItemUnit), "{$InstallDir}", strInstallDir)
        If UserLogined <> True Then
            FoundErr = True
            ErrMsg = ErrMsg & ErrMsg_NoLogin
        Else
            Call GetUser(UserName)
            ErrMsg_PurviewCheckedErr = XmlText("BaseText", "PurviewCheckedErr", "<li>对不起，您没有查看此栏目内容的权限！</li>")
            ErrMsg_PurviewCheckedErr2 = XmlText("BaseText", "PurviewCheckedErr2", "<li>对不起，您没有查看此信息的权限！</li>")
            ErrMsg_NoMail = "<li>" & R_XmlText_Class("SoftContent/NoMail", "对不起，您尚未通过邮件验证，不能查看此{$ChannelShortName}") & "</li>"
            ErrMsg_NoCheck = "<li>" & R_XmlText_Class("SoftContent/NoCheck", "对不起，您尚未通过管理员审核，不能查看收费{$ChannelShortName}") & "</li>"
            ErrMsg_NeedPoint = Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("SoftContent/NeedPoint", "<p align='center'><br><br>对不起，下载本软件需要消耗 <b><font color=red>{$NeedPoint}</font></b> {$PointUnit}{$PointName}！而你目前只有 <b><font color=blue>{$NowPoint}</font></b> {$PointUnit}{$PointName}可用。{$PointName}数不足，无法下载本软件。请与我们联系进行充值。</p>"), "{$InfoPoint}", InfoPoint), "{$NeedPoint}", InfoPoint), "{$NowPoint}", UserPoint), "{$PointName}", PointName), "{$PointUnit}", PointUnit)
            ErrMsg_UsePoint = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("SoftContent/UsePoint", "<p align='center'><br><br>下载本软件需要消耗 <b><font color=red>{$InfoPoint}</font></b> {$PointUnit}{$PointName}！你目前尚有 <b><font color=blue>{$NowPoint}</font></b> {$PointUnit}{$PointName}可用。下载本软件后，你将剩下 <b><font color=green>{$FinalPoint}</font></b> {$PointUnit}{$PointName}<br><br>你确实愿意花费 <b><font color=red>{$InfoPoint}</font></b> {$PointUnit}{$PointName}来下载本软件吗？<br><br><a href='{$FileName}?Pay=yes&UrlID={$UrlID}&SoftID={$SoftID}'>我愿意</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='{$InstallDir}index.asp'>我不愿意</a></p>"), "{$InfoPoint}", InfoPoint), "{$NowPoint}", UserPoint), "{$FinalPoint}", UserPoint - InfoPoint), "{$FileName}", strFileName), "{$SoftID}", SoftID), "{$InstallDir}", strInstallDir), "{$PointName}", PointName), "{$PointUnit}", PointUnit), "{$UrlID}", UrlID)
            ErrMsg_OutTime = R_XmlText_Class("SoftContent/OutTime", "<p align='center'><br><br><font color=red>对不起，本软件为收费内容，而您的有效期已经过期，所以无法下载本软件。请与我们联系进行充值。</font></p>")
            ErrMsg_Overflow_Total = "<li>" & R_XmlText_Class("SoftContent/Overflow_Total", "你已经达到或超过有效期内所能查看的信息总数！") & "</li>"
            ErrMsg_Overflow_Today = "<li>" & R_XmlText_Class("SoftContent/Overflow_Today", "你已经达到或超过今天所能查看的信息总数！") & "</li>"
            Select Case InfoPurview
            Case 0
                If ClassPurview > 0 Then
                    
                    ClassID = rsSoft("ClassID")
                    Call GetClass

                    If ParentID > 0 Then
                        PurviewChecked = CheckPurview_Class(arrClass_View, ChannelDir & "all," & ParentPath & "," & ClassID)
                    Else
                        PurviewChecked = CheckPurview_Class(arrClass_View, ChannelDir & "all," & ClassID)
                    End If
                    If PurviewChecked = False Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_PurviewCheckedErr
                    End If
                Else
                    PurviewChecked = True
                End If
            Case 1
                If GroupType < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & ErrMsg_NoMail
                ElseIf GroupType = 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & ErrMsg_NoCheck
                Else
                    PurviewChecked = True
                End If
            Case 2
                PurviewChecked = FoundInArr(arrGroupID, GroupID, ",")
                If PurviewChecked = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & ErrMsg_PurviewCheckedErr2
                End If
            End Select
            If PurviewChecked = True Then
                If InfoPoint > 0 And InfoPoint < 9999 Then
                    If GroupType < 1 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_NoMail
                    ElseIf GroupType = 1 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_NoCheck
                    Else
                        Dim trs, ValidConsumeLogID, DividePoint

                        If UserChargeType = 0 Then   '点数优先
                            ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, SoftID, ChargeType, PitchTime, ReadTimes)
                            If ValidConsumeLogID = 0 Then   '如果没有找到记录消费，则要开始计费
                                If UserPoint < InfoPoint Then  '如果用户的点数小于要扣的点数
                                    FoundErr = True
                                    ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                Else
                                    If LCase(Trim(Request("Pay"))) = "yes" Then  '如果用户确认要扣点
                                        Conn.Execute "update PE_User set UserPoint=UserPoint-" & InfoPoint & " where UserName='" & UserName & "'"
                                        Call AddConsumeLog("System", ModuleType, UserName, SoftID, InfoPoint, 2, "用于下载收费" & ChannelShortName & "：" & rsSoft("SoftName"))
                                        If DividePercent <= 0 Then
                                            DividePoint = 0
                                        ElseIf DividePercent > 0 And DividePercent < 100 Then
                                            DividePoint = PE_CLng(InfoPoint * DividePercent / 100)
                                        Else
                                            DividePoint = InfoPoint
                                        End If
                                        If DividePoint > 0 Then
                                            Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsSoft("Inputer") & "'"
                                            Call AddConsumeLog("System", ModuleType, rsSoft("Inputer"), 0, DividePoint, 1, "从“" & rsSoft("SoftName") & "”的收费中分成")
                                        End If
                                    Else    '在用户没有确认前，先进行扣费提示
                                        FoundErr = True
                                        ErrMsg = ErrMsg & ErrMsg_UsePoint
                                    End If
                                End If
                            Else   '如果找到了消费记录，直接更新消费记录的消费次数
                                Conn.Execute ("update PE_ConsumeLog set Times=Times+1,IP='" & UserTrueIP & "' where LogID=" & ValidConsumeLogID & "")
                            End If
                        Else
                            If ValidDays <= 0 Then  '过期
                                If UserChargeType = 1 Or UserChargeType = 2 Then '有效期优先，或者是同时判断点券和有效期：点券用完或有效期到期后，就不可查看收费内容
                                    FoundErr = True
                                    ErrMsg = ErrMsg & ErrMsg_OutTime
                                Else
                                    '过期后按照点券优先来算
                                    ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, SoftID, ChargeType, PitchTime, ReadTimes)
                                    If ValidConsumeLogID = 0 Then   '如果没有找到记录消费，则要开始计费
                                        If UserPoint < InfoPoint Then  '如果用户的点数小于要扣的点数
                                            FoundErr = True
                                            ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                        Else
                                            If LCase(Trim(Request("Pay"))) = "yes" Then  '如果用户确认要扣点
                                                Conn.Execute "update PE_User set UserPoint=UserPoint-" & InfoPoint & " where UserName='" & UserName & "'"
                                                Call AddConsumeLog("System", ModuleType, UserName, SoftID, InfoPoint, 2, "用于下载收费" & ChannelShortName & "：" & rsSoft("SoftName"))
                                                If DividePercent <= 0 Then
                                                    DividePoint = 0
                                                ElseIf DividePercent > 0 And DividePercent < 100 Then
                                                    DividePoint = PE_CLng(InfoPoint * DividePercent / 100)
                                                Else
                                                    DividePoint = InfoPoint
                                                End If
                                                If DividePoint > 0 Then
                                                    Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsSoft("Inputer") & "'"
                                                    Call AddConsumeLog("System", ModuleType, rsSoft("Inputer"), 0, DividePoint, 1, "从“" & rsSoft("SoftName") & "”的收费中分成")
                                                End If
                                            Else    '在用户没有确认前，先进行扣费提示
                                                FoundErr = True
                                                ErrMsg = ErrMsg & ErrMsg_UsePoint
                                            End If
                                        End If
                                    Else   '如果找到了消费记录，直接更新消费记录的消费次数
                                        Conn.Execute ("update PE_ConsumeLog set Times=Times+1,IP='" & UserTrueIP & "' where LogID=" & ValidConsumeLogID & "")
                                    End If
                                End If
                            Else   '有效期内
                                '则根据有效期内的扣费方式进行
                                If PE_CLng(UserSetting(15)) > 0 Then   'PE_CLng(UserSetting(15))：有效期内，查看收费内容是否扣点和记录，0为不扣点，1为不扣点，但做记录，2为扣点
                                    '查找消费记录
                                    ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, SoftID, ChargeType, PitchTime, ReadTimes)
                                    If ValidConsumeLogID = 0 Then    '未找到消费记录
                                        If PE_CLng(UserSetting(16)) > 0 Then   '有效期内总共可以查看多少条信息
                                            Set trs = Conn.Execute("select count(0) from PE_ConsumeLog where UserName='" & UserName & "' and Income_Payout=2 and InfoID>0")
                                            If PE_CLng(trs(0)) >= PE_CLng(UserSetting(16)) Then
                                                FoundErr = True
                                                ErrMsg = ErrMsg & ErrMsg_Overflow_Total
                                            End If
                                            Set trs = Nothing
                                        End If
                                        If PE_CLng(UserSetting(17)) > 0 Then    '有效期内每天可以下载多少条信息
                                            Set trs = Conn.Execute("select count(0) from PE_ConsumeLog where UserName='" & UserName & "' and Income_Payout=2 and InfoID>0 and DateDiff(" & PE_DatePart_D & ",LogTime," & PE_Now & ")<1")
                                            If PE_CLng(trs(0)) >= PE_CLng(UserSetting(17)) Then
                                                FoundErr = True
                                                ErrMsg = ErrMsg & ErrMsg_Overflow_Today
                                            End If
                                            Set trs = Nothing
                                        End If
                                        If FoundErr = False Then
                                            If PE_CLng(UserSetting(15)) = 1 Then  '不扣点，但做记录
                                                Call AddConsumeLog("System", ModuleType, UserName, SoftID, 0, 2, "有效期内下载收费" & ChannelShortName & "：" & rsSoft("SoftName") & "，应扣点数：" & InfoPoint & "")
                                            Else  '扣点
                                                If UserPoint >= InfoPoint Then   '如果点数足够
                                                    Conn.Execute "update PE_User set UserPoint=UserPoint-" & InfoPoint & " where UserName='" & UserName & "'"
                                                    Call AddConsumeLog("System", ModuleType, UserName, SoftID, InfoPoint, 2, "有效期内下载收费" & ChannelShortName & "：" & rsSoft("SoftName"))
                                                    If DividePercent <= 0 Then
                                                        DividePoint = 0
                                                    ElseIf DividePercent > 0 And DividePercent < 100 Then
                                                        DividePoint = PE_CLng(InfoPoint * DividePercent / 100)
                                                    Else
                                                        DividePoint = InfoPoint
                                                    End If
                                                    If DividePoint > 0 Then
                                                        Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsSoft("Inputer") & "'"
                                                        Call AddConsumeLog("System", ModuleType, rsSoft("Inputer"), 0, DividePoint, 1, "从“" & rsSoft("SoftName") & "”的收费中分成")
                                                    End If
                                                Else   '点数不够扣时
                                                    If UserChargeType = 2 Then '点数用完或有效期到期后，就不可下载收费内容。
                                                        FoundErr = True
                                                        ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                                    Else   '有效期优先或有效期过期和点数用完，才不可下载收费内容，此时只需记录
                                                        Call AddConsumeLog("System", ModuleType, UserName, SoftID, 0, 2, "有效期内下载收费" & ChannelShortName & "：" & rsSoft("SoftName") & "，应扣点数：" & InfoPoint & "")
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else     '找到消费记录，只更新其消费次数
                                        Conn.Execute ("update PE_ConsumeLog set Times=Times+1,IP='" & UserTrueIP & "' where LogID=" & ValidConsumeLogID & "")
                                    End If
                                Else   '有效期内，下载收费内容不扣点数，也不做记录。
                                    '不做任何处理
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If


    rsSoft.Close
    Set rsSoft = Nothing
    If FoundErr = True Then
        Response.Write ErrMsg
        Exit Function
    End If
    
    Dim sqlHits
    sqlHits = "update PE_Soft set Hits=Hits+1,LastHitTime=" & PE_Now & ""
    If DateDiff("D", LastHitTime, Now()) <= 0 Then
        sqlHits = sqlHits & ",DayHits=DayHits+1"
    Else
        sqlHits = sqlHits & ",DayHits=1"
    End If
    If DateDiff("ww", LastHitTime, Now()) <= 0 Then
        sqlHits = sqlHits & ",WeekHits=WeekHits+1"
    Else
        sqlHits = sqlHits & ",WeekHits=1"
    End If
    If DateDiff("m", LastHitTime, Now()) <= 0 Then
        sqlHits = sqlHits & ",MonthHits=MonthHits+1"
    Else
        sqlHits = sqlHits & ",MonthHits=1"
    End If
    sqlHits = sqlHits & " where SoftID=" & SoftID
    Conn.Execute (sqlHits)

    GetDownloadUrl = DownloadUrl
End Function

Private Function GetPropertyPic(OnTop, IsElite)
    If OnTop = True Then
        GetPropertyPic = "<img src='" & strInstallDir & "images/Soft_ontop.gif' alt='" & strTop & ChannelShortName & "'>"
    ElseIf IsElite = True Then
        GetPropertyPic = "<img src='" & strInstallDir & "images/Soft_elite.gif' alt='" & strElite & ChannelShortName & "'>"
    Else
        GetPropertyPic = "<img src='" & strInstallDir & "images/Soft_common.gif' alt='" & strCommon & ChannelShortName & "'>"
    End If
End Function


Public Function GetCustomFromTemplate(strValue)   '得到自定义列表的版面设计的HTML代码
    Dim strCustom, strParameter
	strCustom = strValue
    regEx.Pattern = "【SoftList\((.*?)\)】([\s\S]*?)【\/SoftList】"
    Set Matches = regEx.Execute(strCustom)
    For Each Match In Matches
        strParameter = Replace(Match.SubMatches(0), Chr(34), " ")
        strCustom = PE_Replace(strCustom, Match.value, GetCustomFromLabel(strParameter, Match.SubMatches(1)))
    Next
    GetCustomFromTemplate = strCustom
End Function

Public Function GetListFromTemplate(ByVal strValue)
    Dim strList
    strList = strValue
    regEx.Pattern = "\{\$GetSoftList\((.*?)\)\}"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        strList = PE_Replace(strList, Match.value, GetListFromLabel(Match.SubMatches(0)))
    Next
    GetListFromTemplate = strList
End Function

Public Function GetPicFromTemplate(ByVal strValue)
    Dim strPicList
    strPicList = strValue
    regEx.Pattern = "\{\$GetPicSoft\((.*?)\)\}"
    Set Matches = regEx.Execute(strPicList)
    For Each Match In Matches
        strPicList = PE_Replace(strPicList, Match.value, GetPicFromLabel(Match.SubMatches(0)))
    Next
    GetPicFromTemplate = strPicList
End Function

Public Function GetSlidePicFromTemplate(ByVal strValue)
    Dim strSlidePic, InitSlideJS
    InitSlideJS = False
    strSlidePic = strValue
    regEx.Pattern = "\{\$GetSlidePicSoft\((.*?)\)\}"
    Set Matches = regEx.Execute(strSlidePic)
    For Each Match In Matches
        If InitSlideJS = False Then
            strSlidePic = PE_Replace(strSlidePic, Match.value, JS_SlidePic & GetSlidePicFromLabel(Match.SubMatches(0)))
            InitSlideJS = True
        Else
            strSlidePic = PE_Replace(strSlidePic, Match.value, GetSlidePicFromLabel(Match.SubMatches(0)))
        End If
    Next
    GetSlidePicFromTemplate = strSlidePic
End Function

Private Function GetSlidePicFromLabel(ByVal strSource)
    Dim arrTemp, tChannelID, arrClassID, tSpecialID
    If strSource = "" Then
        GetSlidePicFromLabel = ""
        Exit Function
    End If
    
    arrTemp = Split(strSource, ",")
    
    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        tChannelID = ChannelID
    Case Else
        tChannelID = PE_CLng(arrTemp(0))
    End Select
    
    Select Case Trim(arrTemp(1))
    Case "arrChildID"
        arrClassID = arrChildID
    Case "ClassID"
        arrClassID = ClassID
    Case Else
        arrClassID = arrTemp(1)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")

    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(3))
    End Select
    
    Select Case UBound(arrTemp)
    Case 12
        GetSlidePicFromLabel = GetSlidePicSoft(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), -1)
    Case 13
        GetSlidePicFromLabel = GetSlidePicSoft(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)))
    Case Else
        GetSlidePicFromLabel = "函数式标签：{$GetSlidePicSoft(参数列表)}的参数个数不对。请检查模板中的此标签。"
    End Select
End Function

Private Function GetPicFromLabel(ByVal strSource)
    Dim arrTemp, tChannelID, arrClassID, tSpecialID
    If strSource = "" Then
        GetPicFromLabel = ""
        Exit Function
    End If

    strSource = FillInArrStr(strSource, "0", 17)
    arrTemp = Split(strSource, ",")
    
    If UBound(arrTemp) <> 16 Then
        GetPicFromLabel = "函数式标签：{$GetPicSoft(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Exit Function
    End If
    
    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        tChannelID = ChannelID
    Case Else
        tChannelID = PE_CLng(arrTemp(0))
    End Select
    
    Select Case Trim(arrTemp(1))
    Case "rsClass_arrChildID"
        If IsObject(rsClass) Then
            arrClassID = rsClass("arrChildID")
        Else
            arrClassID = arrChildID
        End If
    Case "arrChildID"
        arrClassID = arrChildID
    Case "ClassID"
        arrClassID = ClassID
    Case Else
        arrClassID = arrTemp(1)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")

    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(3))
    End Select
    
    GetPicFromLabel = GetPicSoft(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CBool(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CLng(arrTemp(16)))
End Function

Private Function GetListFromLabel(ByVal strSource)
    Dim arrTemp
    Dim tChannelID, SoftNum, arrClassID, tSpecialID, OrderType, OpenType
    If strSource = "" Then
        GetListFromLabel = ""
        Exit Function
    End If
    
    strSource = Replace(strSource, Chr(34), "")
    strSource = FillInArrStr(strSource, "1,listA,listbg,listbg2", 28)
    arrTemp = Split(strSource, ",")
    If UBound(arrTemp) + 1 < 28 Then
        GetListFromLabel = "函数式标签：{$GetSoftList(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Exit Function
    End If

    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        tChannelID = ChannelID
    Case Else
        tChannelID = PE_CLng(arrTemp(0))
    End Select

    Select Case Trim(arrTemp(1))
    Case "rsClass_arrChildID"
        If IsObject(rsClass) Then
            arrClassID = rsClass("arrChildID")
        Else
            arrClassID = arrChildID
        End If
    Case "arrChildID"
        arrClassID = arrChildID
    Case "ClassID"
        arrClassID = ClassID
    Case Else
        arrClassID = arrTemp(1)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")
    
    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(3))
    End Select
    
    Select Case Trim(arrTemp(5))
    Case "rsClass_TopNumber"
        SoftNum = 8
    Case "TopNumber"
        SoftNum = 8
    Case Else
        SoftNum = PE_CLng(arrTemp(5))
    End Select
    

    Select Case Trim(arrTemp(10))
    Case "rsClass_ItemListOrderType"
        OrderType = rsClass("ItemListOrderType")
    Case "ItemListOrderType"
        OrderType = ItemListOrderType
    Case Else
        OrderType = PE_CLng(arrTemp(10))
    End Select

    Select Case Trim(arrTemp(23))
    Case "rsClass_ItemOpenType"
        OpenType = rsClass("ItemOpenType")
    Case "ItemOpenType"
        OpenType = ItemOpenType
    Case Else
        OpenType = PE_CLng(arrTemp(23))
    End Select
    
    GetListFromLabel = GetSoftList(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), SoftNum, PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), arrTemp(8), PE_CLng(arrTemp(9)), OrderType, PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CBool(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CLng(arrTemp(17)), PE_CBool(arrTemp(18)), PE_CBool(arrTemp(19)), PE_CBool(arrTemp(20)), PE_CBool(arrTemp(21)), PE_CBool(arrTemp(22)), OpenType, PE_CLng(arrTemp(24)), Trim(arrTemp(25)), Trim(arrTemp(26)), Trim(arrTemp(27)))
End Function

Private Function GetCustomFromLabel(strTemp, strList)
    Dim arrTemp
    Dim strSoftPic, strPicTemp, arrPicTemp
    Dim iChannelID, arrClassID, IncludeChild, iSpecialID, ItemNum, IsHot, IsElite, Author, DateNum, OrderType, UsePage, TitleLen, ContentLen
    Dim iCols, iColsHtml, iRows, iRowsHtml, iNumber
    Dim IncludePic    
    If strTemp = "" Or strList = "" Then GetCustomFromLabel = "": Exit Function

    iCols = 1: iRows = 1: iColsHtml = "": iRowsHtml = ""
    regEx.Pattern = "【(Cols|Rows)=(\d{1,2})\s*(?:\||｜)(.+?)】"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        If LCase(Match.SubMatches(0)) = "cols" Then
            If Match.SubMatches(1) > 1 Then iCols = Match.SubMatches(1)
            iColsHtml = Match.SubMatches(2)
        ElseIf LCase(Match.SubMatches(0)) = "rows" Then
            If Match.SubMatches(1) > 1 Then iRows = Match.SubMatches(1)
            iRowsHtml = Match.SubMatches(2)
        End If
        strList = regEx.Replace(strList, "")
    Next
    
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) = 10 Then
        strTemp = "ChannelID," & arrTemp(1) & ",False,0," & arrTemp(0) & "," & arrTemp(3) & "," & arrTemp(4) & "," & arrTemp(5) & "," & arrTemp(6) & "," & arrTemp(7) & "," & arrTemp(8) & ",0," & arrTemp(10)
        arrTemp = Split(strTemp, ",")
    End If
    If UBound(arrTemp) <> 13 and UBound(arrTemp) <> 12 Then
        GetCustomFromLabel = "自定义列表标签：【SoftList(参数列表)】列表内容【/SoftList】的参数个数不对。请检查模板中的此标签。"
        Exit Function
    End If


    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        iChannelID = ChannelID
    Case Else
        iChannelID = PE_CLng(arrTemp(0))
    End Select
    Select Case Trim(arrTemp(1))
    Case "rsClass_arrChildID"
        If IsObject(rsClass) Then
            arrClassID = rsClass("arrChildID")
        Else
            arrClassID = arrChildID
        End If
    Case "arrChildID"
        arrClassID = arrChildID
    Case "ClassID"
        arrClassID = ClassID
    Case Else
        arrClassID = arrTemp(1)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")
    IncludeChild = PE_CBool(arrTemp(2))
    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        iSpecialID = SpecialID
    Case Else
        iSpecialID = PE_CLng(arrTemp(3))
    End Select
    ItemNum = PE_CLng(arrTemp(4))
    IsHot = PE_CBool(arrTemp(5))
    IsElite = PE_CBool(arrTemp(6))
    Author = Replace(Replace(Replace(Trim(arrTemp(7)), "?", ""), "&quot;", ""), Chr(34), "")
    DateNum = PE_CLng(arrTemp(8))
    Select Case Trim(arrTemp(9))
    Case "rsClass_ItemListOrderType"
        OrderType = rsClass("ItemListOrderType")
    Case "ItemListOrderType"
        OrderType = ItemListOrderType
    Case Else
        OrderType = PE_CLng(arrTemp(9))
    End Select
    UsePage = PE_CBool(arrTemp(10))
    TitleLen = PE_CLng(arrTemp(11))
    ContentLen = PE_CLng(arrTemp(12))
    If UBound(arrTemp) = 13  then
        IncludePic = PE_CBool(arrTemp(13))
    Else
        IncludePic = False	    
    End If

    FoundErr = False
    If iChannelID <> PrevChannelID Or ChannelID = 0 Then
        Call GetChannel(iChannelID)
    End If
    PrevChannelID = iChannelID
    If FoundErr = True Then
        GetCustomFromLabel = ErrMsg
        Exit Function
    End If

    Dim rsField, ArrField, iField
    Set rsField = Conn.Execute("select FieldName,LabelName,FieldType from PE_Field where ChannelID=-2 or ChannelID=" & ChannelID & "")
    If Not (rsField.BOF And rsField.EOF) Then
        ArrField = rsField.getrows(-1)
    End If
    Set rsField = Nothing

    Dim sqlCustom, rsCustom, iCount, strCustomList, strThisClass, strLink
    iCount = 0
    sqlCustom = ""
    strThisClass = ""
    strCustomList = ""
    
    sqlCustom = "select "
    If ItemNum > 0 Then
        sqlCustom = sqlCustom & "top " & ItemNum & " "
    End If
    If ContentLen > 0 Then
        sqlCustom = sqlCustom & "S.SoftIntro,"
    End If
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlCustom = sqlCustom & "S." & ArrField(0, iField) & ","
        Next
    End If
    sqlCustom = sqlCustom & "S.ChannelID,S.SoftID,S.ClassID,S.SoftName,S.DownloadUrl,S.SoftVersion,S.Author,S.Keyword,S.UpdateTime,S.CopyrightType,S.OperatingSystem,S.SoftType,S.SoftLanguage"
    sqlCustom = sqlCustom & ",S.Stars,S.Inputer,S.Hits,S.DayHits,S.WeekHits,S.MonthHits,S.OnTop,S.Elite,S.Status,S.SoftSize,S.DemoUrl,S.RegUrl,S.SoftPicUrl,S.InfoPoint"
    sqlCustom = sqlCustom & ",C.ClassName,C.ParentDir,C.ClassDir,C.Readme,C.ClassPurview"
    sqlCustom = sqlCustom & ",S.DownloadUrl"
    sqlCustom = sqlCustom & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, Author, DateNum, OrderType, False, IncludePic)

    Set rsCustom = Server.CreateObject("ADODB.Recordset")
    rsCustom.Open sqlCustom, Conn, 1, 1
    If rsCustom.BOF And rsCustom.EOF Then
        totalPut = 0
        strCustomList = GetInfoList_StrNoItem(arrClassID, iSpecialID, IsHot, IsElite, strHot, strElite)
        rsCustom.Close
        Set rsCustom = Nothing
        GetCustomFromLabel = strCustomList
        Exit Function
    End If

    If UsePage = True Then
        totalPut = rsCustom.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                iMod = 0
                If CurrentPage > UpdatePages Then
                    iMod = totalPut Mod MaxPerPage
                    If iMod <> 0 Then iMod = MaxPerPage - iMod
                End If
                rsCustom.Move (CurrentPage - 1) * MaxPerPage - iMod
            Else
                CurrentPage = 1
            End If
        End If
    End If
    Do While Not rsCustom.EOF
        If iChannelID = 0 Then
            If rsCustom("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsCustom("ChannelID"))
                PrevChannelID = rsCustom("ChannelID")
            End If
        End If
        strTemp = strList

        If UsePage = True Then
            iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1
        Else
            iNumber = iCount + 1
        End If

        strTemp = PE_Replace(strTemp, "{$Number}", iNumber)
        strTemp = PE_Replace(strTemp, "{$ClassID}", rsCustom("ClassID"))
        strTemp = PE_Replace(strTemp, "{$ClassName}", rsCustom("ClassName"))
        strTemp = PE_Replace(strTemp, "{$ParentDir}", rsCustom("ParentDir"))
        strTemp = PE_Replace(strTemp, "{$ClassDir}", rsCustom("ClassDir"))
        strTemp = PE_Replace(strTemp, "{$Readme}", rsCustom("ReadMe"))
        If InStr(strTemp, "{$ClassUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ClassUrl}", GetClassUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("ClassID"), rsCustom("ClassPurview")))

        strLink = "<a href='" & GetSoftUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("SoftID")) & "' "
        If OrderType = 0 Then
            strLink = strLink & " >"
        Else
            strLink = strLink & "  target=""_blank"" >"
        End If

        If InStr(strTemp, "{$SoftUrl}") > 0 Then strTemp = Replace(strTemp, "{$SoftUrl}", GetSoftUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("SoftID")))
        strTemp = PE_Replace(strTemp, "{$SoftID}", rsCustom("SoftID"))
        If TitleLen > 0 Then
            strTemp = Replace(strTemp, "{$SoftName}", GetSubStr(rsCustom("SoftName"), TitleLen, False))
        Else
            strTemp = Replace(strTemp, "{$SoftName}", rsCustom("SoftName"))
        End If
        strTemp = Replace(strTemp, "{$SoftNameOriginal}", rsCustom("SoftName"))
        If InStr(strTemp, "{$SoftVersion}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftVersion}", strLink & rsCustom("SoftVersion") & "</a>")
        If InStr(strTemp, "{$SoftProperty}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftProperty}", GetPropertyPic(rsCustom("OnTop"), rsCustom("Elite")))
        strTemp = PE_Replace(strTemp, "{$SoftSize}", rsCustom("SoftSize"))
        If InStr(strTemp, "{$SoftSize_M}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftSize_M}", Round(rsCustom("SoftSize") / 1024, 2))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        If InStr(strTemp, "{$CopyrightType}") > 0 Then strTemp = PE_Replace(strTemp, "{$CopyrightType}", PE_HTMLEncode(rsCustom("CopyrightType")))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        If ContentLen > 0 Then
            If InStr(strTemp, "{$SoftIntro}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftIntro}", Left(nohtml(rsCustom("SoftIntro")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$SoftIntro}", "")
        End If
        strTemp = PE_Replace(strTemp, "{$OperatingSystem}", rsCustom("OperatingSystem"))
        If InStr(strTemp, "{$SoftType}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftType}", PE_HTMLEncode(rsCustom("SoftType")))
        If InStr(strTemp, "{$SoftLanguage}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftLanguage}", PE_HTMLEncode(rsCustom("SoftLanguage")))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$DayHits}", rsCustom("DayHits"))
        strTemp = PE_Replace(strTemp, "{$WeekHits}", rsCustom("WeekHits"))
        strTemp = PE_Replace(strTemp, "{$MonthHits}", rsCustom("MonthHits"))
        If InStr(strTemp, "{$InfoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$InfoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$SoftPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$SoftPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        
        strTemp = PE_Replace(strTemp, "{$SoftAuthor}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$AuthorEmail}", "")
        strTemp = PE_Replace(strTemp, "{$DemoUrl}", rsCustom("DemoUrl"))
        strTemp = PE_Replace(strTemp, "{$RegUrl}", rsCustom("RegUrl"))
        'strTemp = PE_Replace(strTemp, "{$DownloadUrl}", Mid(rsCustom("DownloadUrl"), InStr(rsCustom("DownloadUrl"),"|")))
       ' Set rsSoft = Conn.Execute("Select * from PE_Soft where SoftID=" & rsCustom("SoftID"))
        If InStr(strTemp, "{$DownloadUrl}") > 0 Then strTemp = Replace(strTemp, "{$DownloadUrl}", GetDownloadUrlList1(0,rsCustom("DownloadUrl"),rsCustom("SoftID")))
        '替换下载图片
        regEx.Pattern = "\{\$SoftPic\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strSoftPic = GetSoftPicUrl(Trim(rsCustom("SoftPicUrl")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)))
            strTemp = Replace(strTemp, Match.value, strSoftPic)
        Next
        
        If IsArray(ArrField) Then
            For iField = 0 To UBound(ArrField, 2)
                Select Case ArrField(2, iField)
                Case 8,9
                    strTemp = PE_Replace(strTemp, ArrField(1, iField), PE_HTMLDecode(rsCustom(Trim(ArrField(0, iField)))))
                Case 4
                    If PE_HTMLDecode(rsCustom(Trim(ArrField(0, iField))))="" or IsNull(PE_HTMLDecode(rsCustom(Trim(ArrField(0, iField))))) or PE_HTMLDecode(rsCustom(Trim(ArrField(0, iField))))="http://" Then
                        strTemp = PE_Replace(strTemp, ArrField(1, iField), "")	
                    Else 
                        strTemp = PE_Replace(strTemp, ArrField(1, iField), "<img  class='fieldImg' src='" &PE_HTMLDecode(rsCustom(Trim(ArrField(0, iField))))&"' border=0>")	
                    End If
                Case Else
                    strTemp = PE_Replace(strTemp, ArrField(1, iField), PE_HTMLEncode(rsCustom(Trim(ArrField(0, iField)))))				
                End Select 
           Next
        End If

        strCustomList = strCustomList & strTemp
        rsCustom.MoveNext
        iCount = iCount + 1
        If iCols > 1 And iCount Mod iCols = 0 Then strCustomList = strCustomList & iColsHtml
        If iRows > 1 And iCount Mod iCols * iRows = 0 Then strCustomList = strCustomList & iRowsHtml
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    rsCustom.Close
    Set rsCustom = Nothing
    
    GetCustomFromLabel = strCustomList
End Function


Private Function GetBrowseTimes()
    If UseCreateHTML > 0 Then
        GetBrowseTimes = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetBrowseTimes.asp?SoftID=" & SoftID & "'></script>"
    Else
        GetBrowseTimes = rsSoft("BrowseTimes")
    End If
End Function



Public Sub GetHtml_Index()
    Dim strTemp, arrTemp, iCols, iClassID
    Dim SoftList_ChildClass, SoftList_ChildClass2

    Call GetChannel(ChannelID)
    ClassID = 0

    strHtml = GetTemplate(ChannelID, 1, Template_Index)
    Call ReplaceCommonLabel

    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    
    If InStr(strHtml, "{$ShowChannelCount}") > 0 Then strHtml = Replace(strHtml, "{$ShowChannelCount}", GetChannelCount())
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
        strHtml = Replace(strHtml, "{$RssHot}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "&Hot=1' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
        strHtml = Replace(strHtml, "{$RssElite}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "&Elite=1' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
        strHtml = Replace(strHtml, "{$RssHot}", "")
        strHtml = Replace(strHtml, "{$RssElite}", "")
    End If

    '得到子栏目列表的版面设计的HTML代码
    regEx.Pattern = "【SoftList_ChildClass】([\s\S]*?)【\/SoftList_ChildClass】"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        SoftList_ChildClass = Match.SubMatches(0)
        strHtml = regEx.Replace(strHtml, "{$SoftList_ChildClass}")
        
        '得到每行显示的列数
        iCols = 1
        regEx.Pattern = "【Cols=(\d{1,2})】"
        Set Matches2 = regEx.Execute(SoftList_ChildClass)
        SoftList_ChildClass = regEx.Replace(SoftList_ChildClass, "")
        For Each Match2 In Matches2
            If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
        Next
     
        '开始循环，得到所有子栏目列表的HTML代码
        SoftList_ChildClass2 = ""
        iClassID = 0
        Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=0 and ShowOnIndex=" & PE_True & " order by RootID")
        Do While Not rsClass.EOF
            strTemp = SoftList_ChildClass
            
            strTemp = GetCustomFromTemplate(strTemp)
            strTemp = GetListFromTemplate(strTemp)
            strTemp = GetPicFromTemplate(strTemp)
            
            strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), rsClass("ClassPurview")))
            strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", rsClass("Readme"))
            strTemp = PE_Replace(strTemp, "{$rsClass_Tips}", rsClass("Tips"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", rsClass("ClassID"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", rsClass("ClassName"))
            strTemp = Replace(strTemp, "{$ShowClassAD}", "")
            
            rsClass.MoveNext
            iClassID = iClassID + 1
            If iClassID Mod iCols = 0 And Not rsClass.EOF Then
                SoftList_ChildClass2 = SoftList_ChildClass2 & strTemp
                If iCols > 1 Then SoftList_ChildClass2 = SoftList_ChildClass2 & "</tr><tr>"
            Else
                SoftList_ChildClass2 = SoftList_ChildClass2 & strTemp
                If iCols > 1 Then SoftList_ChildClass2 = SoftList_ChildClass2 & "<td width='1'></td>"
            End If
        Loop
        rsClass.Close
        Set rsClass = Nothing

        strHtml = Replace(strHtml, "{$SoftList_ChildClass}", SoftList_ChildClass2)
    Next
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If UseCreateHTML = 0 Then
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    Else
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(ChannelUrl & "/", 0, FileExt_Index, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(ChannelUrl & "/", 0, FileExt_Index, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
    End If
    
End Sub

Public Sub GetHtml_Class()
    Dim strTemp, iCols, iClassID

    If Child > 0 And ClassShowType <> 2 Then
        strHtml = arrTemplate(0)
    Else
        strHtml = arrTemplate(1)
    End If
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommonLabel
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    
    strHtml = PE_Replace(strHtml, "{$Meta_Keywords_Class}", Meta_Keywords_Class)
    strHtml = PE_Replace(strHtml, "{$Meta_Description_Class}", Meta_Description_Class)
    strHtml = CustomContent("Class", Custom_Content_Class, strHtml)
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
    If Child > 0 Then    '如果当前栏目有子栏目
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(0, 0, 3, 3, 0, True))
    Else
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(ParentID, 0, 3, 3, 0, True))
    End If
    
    Dim SoftList_CurrentClass, SoftList_CurrentClass2, SoftList_ChildClass, SoftList_ChildClass2
    If Child > 0 And ClassShowType <> 2 Then    '如果当前栏目有子栏目
        ItemCount = PE_CLng(Conn.Execute("select Count(*) from PE_Soft where ClassID=" & ClassID & "")(0))
        If ItemCount <= 0 Then     '如果当前栏目没有内容
            regEx.Pattern = "【SoftList_CurrentClass】([\s\S]*?)【\/SoftList_CurrentClass】"
            strHtml = regEx.Replace(strHtml, "") '再去掉显示当前栏目的只属于本栏目的内容列表
        Else      '如果当前栏目有子栏目并且当前栏目有内容，则需要显示出来。
            regEx.Pattern = "【SoftList_CurrentClass】([\s\S]*?)【\/SoftList_CurrentClass】"
            Set Matches = regEx.Execute(strHtml)
            For Each Match In Matches
                SoftList_CurrentClass = Match.SubMatches(0)
                strHtml = regEx.Replace(strHtml, "{$SoftList_CurrentClass}")
                
                strTemp = SoftList_CurrentClass
                strTemp = GetCustomFromTemplate(strTemp)
                strTemp = GetListFromTemplate(strTemp)
                strTemp = GetPicFromTemplate(strTemp)
                
                strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
                strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", ReadMe)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", ClassName)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", ClassID)
                
                strHtml = Replace(strHtml, "{$SoftList_CurrentClass}", strTemp)
            Next
        End If
        
        '得到子栏目列表的版面设计的HTML代码
        regEx.Pattern = "【SoftList_ChildClass】([\s\S]*?)【\/SoftList_ChildClass】"
        Set Matches = regEx.Execute(strHtml)
        For Each Match In Matches
            SoftList_ChildClass = Match.SubMatches(0)
            strHtml = regEx.Replace(strHtml, "{$SoftList_ChildClass}")
            
            '得到每行显示的列数
            iCols = 1
            regEx.Pattern = "【Cols=(\d{1,2})】"
            Set Matches2 = regEx.Execute(SoftList_ChildClass)
            SoftList_ChildClass = regEx.Replace(SoftList_ChildClass, "")
            For Each Match2 In Matches2
                If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
            Next
            
            '开始循环，得到所有子栏目列表的HTML代码
            iClassID = 0
            Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=" & ClassID & " and IsElite=" & PE_True & " and ClassType=1 order by RootID,OrderID")
            Do While Not rsClass.EOF
                strTemp = SoftList_ChildClass
                
                strTemp = GetCustomFromTemplate(strTemp)
                strTemp = GetListFromTemplate(strTemp)
                strTemp = GetPicFromTemplate(strTemp)
                
                strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), rsClass("ClassPurview")))
                strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", rsClass("Readme"))
                strTemp = PE_Replace(strTemp, "{$rsClass_Tips}", rsClass("Tips"))
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", rsClass("ClassName"))
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", rsClass("ClassID"))
                strTemp = Replace(strTemp, "{$ShowClassAD}", "")
            
                rsClass.MoveNext
                iClassID = iClassID + 1
                If iClassID Mod iCols = 0 And Not rsClass.EOF Then
                    SoftList_ChildClass2 = SoftList_ChildClass2 & strTemp
                    If iCols > 1 Then SoftList_ChildClass2 = SoftList_ChildClass2 & "</tr><tr>"
                Else
                    SoftList_ChildClass2 = SoftList_ChildClass2 & strTemp
                    If iCols > 1 Then SoftList_ChildClass2 = SoftList_ChildClass2 & "<td width='1'></td>"
                End If
            Loop
            rsClass.Close
            Set rsClass = Nothing

            strHtml = Replace(strHtml, "{$SoftList_ChildClass}", SoftList_ChildClass2)
        Next
    End If

    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    Dim strPath
    strPath = ChannelUrl & GetListPath(StructureType, ListFileType, ParentDir, ClassDir)

    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$ClassPicUrl}", ClassPicUrl)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = Replace(strHtml, "{$ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
    strHtml = Replace(strHtml, "{$ClassListUrl}", GetClass_1Url(ParentDir, ClassDir, ClassID, ClassPurview))
    
    If ClassPurview > 1 Then
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    Else
        Select Case UseCreateHTML
        Case 0, 2
            If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
            If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
        Case 1
            If ListFileType > 0 Then
                If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, ClassID, FileExt_List, "", totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
                If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, ClassID, FileExt_List, "", totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
            Else
                If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, 0, FileExt_List, "", totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
                If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, 0, FileExt_List, "", totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
            End If
        Case 3
            If ListFileType > 0 Then
                If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, ClassID, FileExt_List, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
                If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, ClassID, FileExt_List, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
            Else
                If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, 0, FileExt_List, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
                If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, 0, FileExt_List, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
            End If
        End Select
    End If

End Sub


Public Sub GetHtml_Soft()
    strHtml = GetCustomFromTemplate(strHtml)  '必须先解析自定义列表标签
    
    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$SoftID}", SoftID)
    Call ReplaceCommonLabel   '解析通用标签，包含自定义标签

    strHtml = GetCustomFromTemplate(strHtml)  '必须先解析自定义列表标签
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)

    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$SoftID}", SoftID)
    strHtml = Replace(strHtml, "{$PageTitle}", ReplaceText(SoftName, 2))
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    
    If InStr(strHtml, "{$MY_") > 0 Then
        Dim rsField
        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-2 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            strHtml = PE_Replace(strHtml, rsField("LabelName"), PE_HTMLEncode(rsSoft(Trim(rsField("FieldName")))))
            rsField.MoveNext
        Loop
        Set rsField = Nothing
    End If
    
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    If InStr(strHtml, "{$ClassUrl}") > 0 Then strHtml = PE_Replace(strHtml, "{$ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
    strHtml = CustomContent("Class", Custom_Content_Class, strHtml)
    
    If InStr(strHtml, "{$SoftName}") > 0 Then strHtml = PE_Replace(strHtml, "{$SoftName}", ReplaceText(rsSoft("SoftName"), 2))
    strHtml = PE_Replace(strHtml, "{$SoftVersion}", rsSoft("SoftVersion"))
    strHtml = PE_Replace(strHtml, "{$SoftSize}", rsSoft("SoftSize"))
    strHtml = PE_Replace(strHtml, "{$SoftSize_M}", Round(rsSoft("SoftSize") / 1024, 2))
    strHtml = PE_Replace(strHtml, "{$DecompressPassword}", rsSoft("DecompressPassword"))
    strHtml = PE_Replace(strHtml, "{$OperatingSystem}", rsSoft("OperatingSystem"))
   
    If InStr(strHtml, "{$Hits}") > 0 Then strHtml = Replace(strHtml, "{$Hits}", GetHits())
    If InStr(strHtml, "{$DayHits}") > 0 Then strHtml = Replace(strHtml, "{$DayHits}", GetDayHits())
    If InStr(strHtml, "{$WeekHits}") > 0 Then strHtml = Replace(strHtml, "{$WeekHits}", GetWeekHits())
    If InStr(strHtml, "{$MonthHits}") > 0 Then strHtml = Replace(strHtml, "{$MonthHits}", GetMonthHits())
    
    If InStr(strHtml, "{$Author}") > 0 Then strHtml = Replace(strHtml, "{$Author}", GetAuthorInfo(rsSoft("Author"), ChannelID))
    If InStr(strHtml, "{$CopyFrom}") > 0 Then strHtml = Replace(strHtml, "{$CopyFrom}", GetCopyFromInfo(rsSoft("CopyFrom"), ChannelID))
    If InStr(strHtml, "{$SoftLink}") > 0 Then strHtml = Replace(strHtml, "{$SoftLink}", GetSoftLink())
    If InStr(strHtml, "{$SoftType}") > 0 Then strHtml = Replace(strHtml, "{$SoftType}", PE_HTMLEncode(rsSoft("SoftType")))
    If InStr(strHtml, "{$SoftLanguage}") > 0 Then strHtml = Replace(strHtml, "{$SoftLanguage}", PE_HTMLEncode(rsSoft("SoftLanguage")))
    If InStr(strHtml, "{$CopyrightType}") > 0 Then strHtml = Replace(strHtml, "{$CopyrightType}", PE_HTMLEncode(rsSoft("CopyrightType")))
    If InStr(strHtml, "{$SoftProperty}") > 0 Then strHtml = Replace(strHtml, "{$SoftProperty}", GetSoftProperty())
    If InStr(strHtml, "{$Stars}") > 0 Then strHtml = Replace(strHtml, "{$Stars}", GetStars(rsSoft("Stars")))
    If InStr(strHtml, "{$SoftPicUrl}") > 0 Then strHtml = Replace(strHtml, "{$SoftPicUrl}", GetSoftPicUrl(rsSoft("SoftPicUrl"), 150, 0))
    If InStr(strHtml, "{$UpdateDate}") > 0 Then strHtml = Replace(strHtml, "{$UpdateDate}", FormatDateTime(rsSoft("UpdateTime"), 2))
    strHtml = Replace(strHtml, "{$UpdateTime}", rsSoft("UpdateTime"))
    strHtml = PE_Replace(strHtml, "{$Editor}", rsSoft("Editor"))
    strHtml = Replace(strHtml, "{$Inputer}", rsSoft("Inputer"))
    
    If InStr(strHtml, "{$SoftIntro}") > 0 Then strHtml = Replace(strHtml, "{$SoftIntro}", ReplaceKeyLink(ReplaceText(rsSoft("SoftIntro"), 1)))
	
    '替换{$ArticleIntro(Type,InfoLength)}标签
    Dim strSoftIntro
    regEx.Pattern = "\{\$SoftIntro\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strSoftIntro= "函数式标签：{$SoftIntro(参数列表)}的参数个数不对。请检查模板中的此标签。"

        Else
            Select Case PE_Clng(arrTemp(0))
            Case 1
                strSoftIntro = ReplaceKeyLink(ReplaceText(rsSoft("SoftIntro"), 1))
            Case 2
                If PE_Clng(arrTemp(1))>0 then
                    strSoftIntro = GetSubStr(nohtml(rsSoft("SoftIntro")),PE_Clng(arrTemp(1)),False)
                Else
                    strSoftIntro = nohtml(rsSoft("SoftIntro"))
                End IF
            End Select
        End If
        strHtml = Replace(strHtml, Match.Value, strSoftIntro)
	Next
	
    
    If InStr(strHtml, "{$CorrelativeSoft}") > 0 Then strHtml = Replace(strHtml, "{$CorrelativeSoft}", GetCorrelative(5, 50, 1, 0, 1))

    strHtml = PE_Replace(strHtml, "{$SoftAuthor}", rsSoft("Author"))
    strHtml = Replace(strHtml, "{$AuthorEmail}", "")
    strHtml = PE_Replace(strHtml, "{$DemoUrl}", rsSoft("DemoUrl"))
    strHtml = PE_Replace(strHtml, "{$RegUrl}", rsSoft("RegUrl"))
    If InStr(strHtml, "{$InfoPoint}") > 0 Then strHtml = PE_Replace(strHtml, "{$InfoPoint}", GetInfoPoint(rsSoft("InfoPoint")))
    If InStr(strHtml, "{$SoftPoint}") > 0 Then strHtml = PE_Replace(strHtml, "{$SoftPoint}", GetInfoPoint(rsSoft("InfoPoint")))

    If InStr(strHtml, "{$DownloadUrl}") > 0 Then strHtml = Replace(strHtml, "{$DownloadUrl}", GetDownloadUrlList(0))
    If InStr(strHtml, "{$Vote}") > 0 Then strHtml = Replace(strHtml, "{$Vote}", GetVoteOfContent(SoftID)) '投票标签
    strHtml = Replace(strHtml, "{$Rss}", "")
    If InStr(strHtml, "{$Keyword}") > 0 Then strHtml = PE_Replace(strHtml, "{$Keyword}", GetKeywords(",", rsSoft("Keyword")))
    If InStr(strHtml, "{$BrowseTimes}") > 0 Then strHtml = Replace(strHtml, "{$BrowseTimes}", GetBrowseTimes()) '浏览量标签

    Dim arrTemp
    Dim strSoftPic
    '替换下载图片
    regEx.Pattern = "\{\$SoftPic\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strSoftPic = GetSoftPicUrl(Trim(rsSoft("SoftPicUrl")), PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
        strHtml = Replace(strHtml, Match.value, strSoftPic)
    Next

    Dim strCorrelativeSoft
    regEx.Pattern = "\{\$CorrelativeSoft\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        Select Case UBound(arrTemp)
        Case 1
            strCorrelativeSoft = GetCorrelative(arrTemp(0), arrTemp(1), 1, 0, 1)
        Case 4
            strCorrelativeSoft = GetCorrelative(arrTemp(0), arrTemp(1), arrTemp(2), arrTemp(3), arrTemp(4))
        Case Else
            strCorrelativeSoft = "函数式标签：{$CorrelativeSoft(参数列表)}的参数个数不对。请检查模板中的此标签。"
        End Select
        strHtml = Replace(strHtml, Match.value, strCorrelativeSoft)
    Next
    
    '替换下载地址DownloadUrl
    Dim strDownloadUrl
    regEx.Pattern = "\{\$DownloadUrl\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        If Match.SubMatches(0) = "" Then
            strDownloadUrl = "函数式标签：{$DownloadUrl(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strDownloadUrl = GetDownloadUrlList(Match.SubMatches(0))
        End If
        strHtml = Replace(strHtml, Match.value, strDownloadUrl)
    Next

End Sub

Public Sub GetHtml_Special()
    strHtml = PE_Replace(strHtml, "{$SpecialID}", SpecialID)
    Call ReplaceCommonLabel
    strHtml = PE_Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = PE_Replace(strHtml, "{$ShowPath}", ShowPath())
    strHtml = PE_Replace(strHtml, "{$SpecialID}", SpecialID)
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)
    strHtml = PE_Replace(strHtml, "{$SpecialPicUrl}", SpecialPicUrl)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = CustomContent("Special", Custom_Content_Special, strHtml)
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "&SpecialID=" & SpecialID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
    
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    Dim strPath
    strPath = ChannelUrl & "/Special/" & SpecialDir
    
    Select Case UseCreateHTML
    Case 0, 2
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    Case 1
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, 0, FileExt_List, "", totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, 0, FileExt_List, "", totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
    Case 3
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, 0, FileExt_List, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, 0, FileExt_List, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
    End Select
    
End Sub

Public Sub GetHtml_SpecialList()
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "&SpecialID=" & SpecialID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
    strHtml = PE_Replace(strHtml, "{$GetAllSpecial}", GetAllSpecial)
    
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个专题", False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个专题", False))
    
End Sub

Public Sub GetHtml_Search()
    Dim SearchChannelID
    SearchChannelID = ChannelID
    If ChannelID > 0 Then
        strHtml = GetTemplate(ChannelID, 5, 0)
    Else
        strHtml = GetTemplate(ChannelID, 3, 0)
        ChannelID = PE_CLng(Conn.Execute("select min(ChannelID) from PE_Channel where ModuleType=2 and Disabled=" & PE_False & "")(0))
        Call GetChannel(ChannelID)
    End If
    Select Case strField
    Case "Title"
        strField = "SoftName"
    Case "Content"
        strField = "SoftIntro"
    End Select
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())

    If strField <> "" Then
        regEx.Pattern = "【SearchForm】([\s\S]*?)【\/SearchForm】"
        Set Matches = regEx.Execute(strHtml)
        strHtml = regEx.Replace(strHtml, "")
    Else
        If Trim(Request.ServerVariables("QUERY_STRING")) <> "" Then
            regEx.Pattern = "【SearchForm】([\s\S]*?)【\/SearchForm】"
            Set Matches = regEx.Execute(strHtml)
            strHtml = regEx.Replace(strHtml, "")
        Else
            regEx.Pattern = "【ShowResult】([\s\S]*?)【\/ShowResult】"
            Set Matches = regEx.Execute(strHtml)
            strHtml = regEx.Replace(strHtml, "")
        End If
    End If

    Call GetClass
    MaxPerPage = MaxPerPage_SearchResult
    If InStr(strHtml, "{$ResultTitle}") > 0 Then strHtml = Replace(strHtml, "{$ResultTitle}", GetResultTitle())
    If InStr(strHtml, "{$SearchResult}") > 0 Then strHtml = Replace(strHtml, "{$SearchResult}", GetSearchResult(SearchChannelID))
    strHtml = GetSearchResult2(SearchChannelID, strHtml)
    
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)

    strHtml = Replace(strHtml, "【ShowResult】", "")
    strHtml = Replace(strHtml, "【/ShowResult】", "")
    strHtml = Replace(strHtml, "【SearchForm】", "")
    strHtml = Replace(strHtml, "【/SearchForm】", "")
    strHtml = Replace(strHtml, "{$Keyword}", Keyword)
    
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)

    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage_SearchResult, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage_SearchResult, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    
End Sub

Public Sub GetHtml_List()
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)
    
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    
End Sub

Public Sub ShowFavorite()
    Response.Write "<table width='100%' cellpadding='2' cellspacing='1' border='0' class='border'>"
    Response.Write "  <tr class='title' align='center'><td width='30'>选中</td><td>" & ChannelShortName & "名称</td><td width='100'>作者</td><td width='80'>更新时间</td><td width='80'>操作</td></tr>"
    
    Dim sqlFavorite, rsFavorite, iCount, strLink
    iCount = 0
    
    sqlFavorite = "select S.ChannelID,S.SoftID,S.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,S.SoftName,S.SoftVersion,S.Author,S.UpdateTime from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where S.Deleted=" & PE_False & " and S.Status=3"
    sqlFavorite = sqlFavorite & " and SoftID in (select InfoID from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & ")"
    sqlFavorite = sqlFavorite & " order by S.SoftID desc"

    Set rsFavorite = Server.CreateObject("ADODB.Recordset")
    rsFavorite.Open sqlFavorite, Conn, 1, 1
    If rsFavorite.BOF And rsFavorite.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td height='50' colspan='20' align='center'>没有收藏任何" & ChannelShortName & "</td></tr>"
    Else
        totalPut = rsFavorite.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsFavorite.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Do While Not rsFavorite.EOF
            strLink = "[<a href='" & GetClassUrl(rsFavorite("ParentDir"), rsFavorite("ClassDir"), rsFavorite("ClassID"), rsFavorite("ClassPurview")) & "'>" & rsFavorite("ClassName") & "</a>] "
            strLink = strLink & "<a href='" & GetSoftUrl(rsFavorite("ParentDir"), rsFavorite("ClassDir"), rsFavorite("UpdateTime"), rsFavorite("SoftID")) & "' target='_blank'>" & rsFavorite("SoftName") & " " & rsFavorite("SoftVersion") & "</a>"
            
            Response.Write "<tr class='tdbg'>"
            Response.Write "<td align='center' width='30'><input type='checkbox' name='InfoID' value='" & rsFavorite("SoftID") & "'></td>"
            Response.Write "<td align='left'>" & strLink & "</td>"
            Response.Write "<td width='100' align='center'>" & rsFavorite("Author") & "</td>"
            Response.Write "<td width='80' align='right'>" & Year(rsFavorite("UpdateTime")) & "-" & Right("0" & Month(rsFavorite("UpdateTime")), 2) & "-" & Right("0" & Day(rsFavorite("UpdateTime")), 2) & "</td>"
            Response.Write "<td width='80' align='center'><a href='User_Favorite.asp?Action=Remove&ChannelID=" & ChannelID & "&InfoID=" & rsFavorite("SoftID") & "' onclick=""return confirm('确实不再收藏此" & ChannelShortName & "吗？');"">取消收藏</a></td>"
            Response.Write "</tr>"
            
            iCount = iCount + 1
            If iCount >= MaxPerPage Then Exit Do
            rsFavorite.MoveNext
        Loop
    End If
    rsFavorite.Close
    Set rsFavorite = Nothing
    Response.Write "</table>"
    Response.Write ShowPage("User_Favorite.asp?ChannelID=" & ChannelID & "", totalPut, 20, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False)
End Sub


Function XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    XmlText_Class = XmlText("Soft", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Soft", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>
