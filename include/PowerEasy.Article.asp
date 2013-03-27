<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim ArticleID, ArticleTitle, ArticleUrl
Dim rsArticle

Class Article

Private ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4
Private rsClass
Private strTempContent, strContentPageTitleArr
Public totalPage

'��ʼ����Ҫ�õ���һЩ����
Public Sub Init()
    FoundErr = False
    ErrMsg = ""
    PrevChannelID = ChannelID
    ChannelShortName = "����"
        
    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
     
    '*****************************
    '��ȡ���԰��е��ַ�����
    strListStr_Font = XmlText_Class("ArticleList/UpdateTimeColor_New", "color=""red""")
    strTop = XmlText_Class("ArticleList/t4", "�̶�")
    strElite = XmlText_Class("ArticleList/t3", "�Ƽ�")
    strCommon = XmlText_Class("ArticleList/t5", "��ͨ")
    strHot = XmlText_Class("ArticleList/t7", "�ȵ�")
    strNew = XmlText_Class("ArticleList/t6", "����")
    strTop2 = XmlText_Class("ArticleList/Top", " ��")
    strElite2 = XmlText_Class("ArticleList/Elite", " ��")
    strHot2 = XmlText_Class("ArticleList/Hot", " ��")
    ArticlePro1 = XmlText_Class("ArticlePro1", "[ͼ��]")
    ArticlePro2 = XmlText_Class("ArticlePro2", "[��ͼ]")
    ArticlePro3 = XmlText_Class("ArticlePro3", "[�Ƽ�]")
    ArticlePro4 = XmlText_Class("ArticlePro4", "[ע��]")
    Character_Author = XmlText("Article", "Include/Author", "[{$Text}]")
    Character_Date = XmlText("Article", "Include/Date", "[{$Text}]")
    Character_Hits = XmlText("Article", "Include/Hits", "[{$Text}]")
    Character_Class = XmlText("Article", "Include/ClassChar", "[{$Text}]")
    SearchResult_Content_NoPurview = XmlText("BaseText", "SearchPurviewContent", "��������Ҫ��ָ��Ȩ�޲ſ���Ԥ��")
    SearchResult_ContentLenth = PE_CLng(XmlText_Class("ShowSearch/Content_Lenght", "200"))
    strList_Content_Div = XmlText_Class("ArticleList/Content_DIV", "style=""padding:0px 20px""")
    strList_Title = R_XmlText_Class("ArticleList/Title", "{$ChannelShortName}���⣺{$Title}{$br}��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�{$Author}{$br}����ʱ�䣺{$UpdateTime}")
    strComment = XmlText_Class("ArticleList/CommentLink", "<font color=""red"">����</font>")
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
'��������ShowChannelCount
'��  �ã���ʾƵ��ͳ����Ϣ
'��  ������
'=================================================
Private Function GetChannelCount()
    GetChannelCount = Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ChannelCount", "{$ChannelShortName}������ {$ItemChecked_Channel} {$ChannelItemUnit}<br>����{$ChannelShortName}�� {$UnItemChecked} {$ChannelItemUnit}<br>���������� {$CommentCount_Channel} ��<br>ר�������� {$SpecialCount_Channel} ��<br>{$ChannelShortName}�Ķ��� {$HitsCount_Channel} �˴�<br>"), "{$ItemChecked_Channel}", ItemChecked_Channel), "{$ChannelItemUnit}", ChannelItemUnit), "{$UnItemChecked}", treatAuditing("Article", ChannelID)), "{$CommentCount_Channel}", CommentCount_Channel), "{$SpecialCount_Channel}", SpecialCount_Channel), "{$HitsCount_Channel}", "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?Action=Count'></script>")
End Function
'**************************************************
'��������treatAuditing
'��  �ã�����˺���
'��  ����ModuleName ----����
'        ChannelID ---- Ƶ��ID
'����ֵ���������Ŀ��
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
        strSql = strSql & " from PE_InfoS I inner join (PE_Article A left join PE_Class C on A.ClassID=C.ClassID) on I.ItemID=A.ArticleID"
    Else
        strSql = strSql & " from PE_Article A left join PE_Class C on A.ClassID=C.ClassID"
    End If
    strSql = strSql & " where A.Deleted=" & PE_False & " and A.Status=3 and A.ReceiveType=0"
	
    If InStr(iChannelID, ",") > 0 Then
        strSql = strSql & " and A.ChannelID in (" & FilterArrNull(iChannelID, ",") & ")"
    Else
        If PE_CLng(iChannelID) > 0 Then strSql = strSql & " and A.ChannelID=" & PE_CLng(iChannelID)
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
            strSql = strSql & " and A.ClassID in (" & FilterArrNull(arrClassID, ",") & ")"
        Else
            If PE_CLng(arrClassID) > 0 Then strSql = strSql & " and A.ClassID=" & PE_CLng(arrClassID)
        End If
    End If
    If iSpecialID > 0 Then
        strSql = strSql & " and I.ModuleType=1 and I.SpecialID=" & iSpecialID
    End If
    If IsHot = True Then
        strSql = strSql & " and A.Hits>=" & HitsOfHot
    End If
    If IsElite = True Then
        strSql = strSql & " and A.Elite=" & PE_True
    End If
    If Trim(Author) <> "" Then
        strSql = strSql & " and A.Author='" & ReplaceBadChar(Author) & "'"
    End If
    If DateNum > 0 Then
        strSql = strSql & " and DateDiff(" & PE_DatePart_D & ",A.UpdateTime," & PE_Now & ")<" & DateNum
    End If

    If IsPicUrl = True Then
        strSql = strSql & " and A.DefaultPicUrl<>'' "
    End If

    strSql = strSql & " order by A.OnTop " & PE_OrderType & ","
    Select Case PE_CLng(OrderType)
    Case 1, 2
    
    Case 3
        strSql = strSql & "A.UpdateTime desc,"
    Case 4
        strSql = strSql & "A.UpdateTime asc,"
    Case 5
        strSql = strSql & "A.Hits desc,"
    Case 6
        strSql = strSql & "A.Hits asc,"
    Case 7
        strSql = strSql & "A.CommentCount desc,"
    Case 8
        strSql = strSql & "A.CommentCount asc,"
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
        strSql = strSql & "A.ArticleID " & IDOrder
    End If
    GetSqlStr = strSql
End Function

'=================================================
'��������GetArticleList
'��  �ã���ʾ���±������Ϣ
'��  ����
'0        iChannelID ---- Ƶ��ID
'1        arrClassID ---- ��ĿID���飬0Ϊ������Ŀ
'2        IncludeChild ---- �Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'3        iSpecialID ---- ר��ID��0Ϊ�������£�����ר�����£������Ϊ����0����ֻ��ʾ��Ӧר�������
'4        UrlType ---- ���ӵ�ַ���ͣ�0Ϊ���·����1Ϊ����ַ�ľ���·���������⹫����4.03ʱΪShowAllArticle
'5        ArticleNum ---- ��������������0����ֻ��ѯǰ��ƪ����
'6        IsHot ---- �Ƿ����������£�TrueΪֻ��ʾ�������£�FalseΪ��ʾ��������
'7        IsElite ---- �Ƿ����Ƽ����£�TrueΪֻ��ʾ�Ƽ����£�FalseΪ��ʾ��������
'8        Author ---- ���������������Ϊ�գ���ֻ��ʾָ�����ߵ����£����������ļ�
'9        DateNum ---- ���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�����
'10       OrderType ---- ����ʽ��1--������ID����2--������ID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'11       ShowType ---- ��ʾ��ʽ��1Ϊ��ͨ��ʽ��2Ϊ���ʽ��3Ϊ�������ʽ��4Ϊ���ܶ���ʽ��5Ϊ���DIV��6Ϊ���RSS
'12       TitleLen ---- ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ��������
'13       ContentLen ---- ������������ַ�����һ������=����Ӣ���ַ���Ϊ0ʱ����ʾ�������������϶࣬���ܻᵼ���������
'14       ShowClassName ---- �Ƿ���ʾ������Ŀ���ƣ�TrueΪ��ʾ��FalseΪ����ʾ
'15       ShowPropertyType ---- ��ʾ�������ԣ��̶�/�Ƽ�/��ͨ���ķ�ʽ��0Ϊ����ʾ��1ΪСͼƬ��2Ϊ���ţ�3--9ΪСͼƬ��10Ϊ���
'16       ShowIncludePic ---- �Ƿ���ʾ��[ͼ��]��������TrueΪ��ʾ��FalseΪ����ʾ
'17       ShowAuthor ---- �Ƿ���ʾ�������ߣ�TrueΪ��ʾ��FalseΪ����ʾ
'18       ShowDateType ---- ��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���գ�3Ϊ�ԡ���-�ա���ʽ��ʾ���ա�
'19       ShowHits ---- �Ƿ���ʾ���µ������TrueΪ��ʾ��FalseΪ����ʾ
'20       ShowHotSign ---- �Ƿ���ʾ�������±�־��TrueΪ��ʾ��FalseΪ����ʾ
'21       ShowNewSign ---- �Ƿ���ʾ�����±�־��TrueΪ��ʾ��FalseΪ����ʾ
'22       ShowTips ---- �Ƿ���ʾ���ߡ��������ڡ�������ȸ�����ʾ��Ϣ��TrueΪ��ʾ��FalseΪ����ʾ
'23       ShowCommentLink ---- �Ƿ���ʾ�������ӣ�TrueΪ��ʾ��FalseΪ����ʾ����ѡ��ֻ�е���Ӧ�����ں�̨�����ˡ���ʾ�������ӡ�����Ч��
'24       UsePage ---- �Ƿ��ҳ��ʾ��TrueΪ��ҳ��ʾ��FalseΪ����ҳ��ʾ��ÿҳ��ʾ������������MaxPerPageָ��
'25       OpenType ---- ���´򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�
'26       Cols ---- ÿ�е������������������ͻ��С�
'27       CssNameA ---- �б����������ӵ��õ�CSS����
'28       CssName1 ---- �б��������е�CSSЧ��������
'29       CssName2 ---- �б���ż���е�CSSЧ��������
'=================================================
Public Function GetArticleList(iChannelID, arrClassID, IncludeChild, iSpecialID, UrlType, ArticleNum, IsHot, IsElite, Author, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowIncludePic, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, ShowCommentLink, UsePage, OpenType, Cols, CssNameA, CssName1, CssName2)
    Dim sqlInfo, rsInfoList, strInfoList, CssName, iCount, iNumber, InfoUrl
    Dim strProperty, strTitle, strLink, strAuthor, strUpdateTime, strHits, strHotSign, strNewSign, strContent, strClassName
    Dim iTitleLen, strCommentLink
    Dim TDWidth_Author, TdWidth_Date

    TDWidth_Author = 10 * AuthorInfoLen
    TdWidth_Date = GetTdWidth_Date(ShowDateType)

    iCount = 0
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)

    If ShowType = 6 Then UrlType = 1
    If TitleLen < 0 Or TitleLen > 200 Then TitleLen = 50
    If IsNull(CssNameA) Then CssNameA = "listA"
    If IsNull(CssName1) Then CssName1 = "listbg"
    If IsNull(CssName2) Then CssName2 = "listbg2"

    FoundErr = False
    If (PE_Clng(iChannelID) <> 0 and Instr(iChannelID,",")=0) and (PE_Clng(iChannelID)<>PrevChannelID Or ChannelID = 0) Then
        Call GetChannel(PE_Clng(iChannelID))
        PrevChannelID = iChannelID		 
    End If
    If FoundErr = True Then
        GetArticleList = ErrMsg
        Exit Function
    End If

    sqlInfo = "select"
    If ArticleNum > 0 Then
        If ShowType = 4 Then
            sqlInfo = sqlInfo & " top " & ArticleNum * 4
        Else
            sqlInfo = sqlInfo & " top " & ArticleNum
        End If
    End If
    sqlInfo = sqlInfo & " A.ChannelID,A.ClassID,A.ArticleID,A.Title,A.TitleFontColor,A.TitleFontType,A.ShowCommentLink,A.IncludePic,A.Author,A.UpdateTime,A.Hits,A.OnTop,A.Elite,A.InfoPurview,A.InfoPoint"
    If ContentLen > 0 Then
        sqlInfo = sqlInfo & ",A.Intro,A.Content"
    End If
    sqlInfo = sqlInfo & ",C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview"
    sqlInfo = sqlInfo & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, Author, DateNum, OrderType, ShowClassName, False)
    Set rsInfoList = Server.CreateObject("ADODB.Recordset")
    rsInfoList.Open sqlInfo, Conn, 1, 1
    If rsInfoList.BOF And rsInfoList.EOF Then
        If UsePage = True Then totalPut = 0
        If ShowType < 6 Then
            strInfoList = GetInfoList_StrNoItem(arrClassID, iSpecialID, IsHot, IsElite, strHot, strElite)
        End If
        rsInfoList.Close
        Set rsInfoList = Nothing
        GetArticleList = strInfoList
        Exit Function
    End If
    If UsePage = True And ShowType < 6 Then
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

    If ShowType = 6 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
    If ShowType = 2 Or ShowType = 4 Or (Cols > 1 and ShowType<>5) Then
        strInfoList = "<table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr>"
    Else
        strInfoList = ""
    End If

    Dim CurrentTitleLen, isfirst, rownum, outend
    CurrentTitleLen = 0
    isfirst = True
    rownum = 1
    outend = False
    Do While Not rsInfoList.EOF
        'If iChannelID = 0 Then
            If rsInfoList("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsInfoList("ChannelID"))
                PrevChannelID = rsInfoList("ChannelID")
            End If
       ' End If
        If UsePage = True Then
            iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1
        Else
            iNumber = iCount + 1
        End If

        ChannelUrl = UrlPrefix(UrlType, ChannelUrl) & ChannelUrl
        ChannelUrl_ASPFile = UrlPrefix(UrlType, ChannelUrl_ASPFile) & ChannelUrl_ASPFile
        InfoUrl = GetArticleUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("UpdateTime"), rsInfoList("ArticleID"), rsInfoList("ClassPurview"), rsInfoList("InfoPurview"), rsInfoList("InfoPoint"))
        If ShowType < 6 And ShowType <> 4 Then

            strProperty = GetInfoList_GetStrProperty(ShowPropertyType, rsInfoList("OnTop"), rsInfoList("Elite"), iNumber, strCommon, strTop, strElite)
            strHotSign = GetInfoList_GetStrHotSign(ShowHotSign, rsInfoList("Hits"), strHot)
            strNewSign = GetInfoList_GetStrNewSign(ShowNewSign, rsInfoList("UpdateTime"), strNew)
            strCommentLink = GetInfoList_GetStrCommentLink(ShowCommentLink, rsInfoList("ShowCommentLink"), rsInfoList("ArticleID"))
            strAuthor = GetSubStr(rsInfoList("Author"), AuthorInfoLen, True)
            strUpdateTime = GetInfoList_GetStrUpdateTime(rsInfoList("UpdateTime"), ShowDateType)
            strHits = rsInfoList("Hits")
            If ShowType = 3 Or ShowType = 5 Then
                strAuthor = GetInfoList_GetStrAuthor_Xml(ShowAuthor, strAuthor)
                strUpdateTime = GetInfoList_GetStrUpdateTime_Xml(ShowDateType, strUpdateTime)
                strHits = GetInfoList_GetStrHits_Xml(ShowHits, strHits)
            End If

            iTitleLen = GetInfoList_GetTitleLen(TitleLen, ShowIncludePic, ShowCommentLink, rsInfoList("IncludePic"), rsInfoList("ShowCommentLink"))
            strTitle = GetInfoList_GetStrTitle(rsInfoList("Title"), iTitleLen, rsInfoList("TitleFontType"), rsInfoList("TitleFontColor"))

            strLink = ""
            If ShowClassName = True Then
                strLink = strLink & GetInfoList_GetStrClassLink(Character_Class, CssNameA, rsInfoList("ClassID"), rsInfoList("ClassName"), GetClassUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("ClassID"), rsInfoList("ClassPurview")))
            End If
            If ShowIncludePic = True Then
                strLink = strLink & GetInfoList_GetStrIncludePic(rsInfoList("IncludePic"))
            End If
            strLink = strLink & GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, CssNameA, strTitle, InfoUrl, rsInfoList("Title"), rsInfoList("Author"), rsInfoList("UpdateTime"))
            strContent = ""
            Select Case PE_CLng(ShowType)
            Case 1, 3, 5
                If ContentLen > 0 Then
                    strContent = strContent & "<div " & strList_Content_Div & ">"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("Content"), rsInfoList("Intro"))
                    strContent = strContent & "</div>"
                End If
            Case 2
                If ContentLen > 0 Then
                    strContent = "<tr><td colspan=""10"" class=""" & CssName & """>"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("Content"), rsInfoList("Intro"))
                    strContent = strContent & "</td></tr>"
                End If
            End Select

        ElseIf ShowType = 6 Then
            strTitle = GetInfoList_GetStrTitle(rsInfoList("Title"), TitleLen, rsInfoList("TitleFontType"), rsInfoList("TitleFontColor"))
            strTitle = ReplaceText(xml_nohtml(strTitle), 2)
            strLink = InfoUrl
            If ContentLen > 0 Then
                If Trim(rsInfoList("Intro") & "") = "" Then
                    strContent = Left(Replace(Replace(Replace(xml_nohtml(rsInfoList("Content")), "[NextPage]", ""), ">", "&gt;"), "<", "&lt;"), ContentLen)
                Else
                    strContent = Left(xml_nohtml(rsInfoList("Intro")), ContentLen)
                End If
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
            strInfoList = strInfoList & GetInfoList_GetStrAuthorDateHits(ShowAuthor, ShowDateType, ShowHits, strAuthor, strUpdateTime, strHits, rsInfoList("ChannelID"))
            strInfoList = strInfoList & strHotSign & strNewSign & strCommentLink & strContent & "<br />"

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
            strInfoList = strInfoList & "<td class=""" & CssName & """>" & strLink & strHotSign & strNewSign & strCommentLink & "</td>"
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
            strInfoList = strInfoList & strHotSign & strNewSign & strCommentLink & strContent
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
        Case 5 '���DIV
            strInfoList = strInfoList & "<div class=""" & CssName & """>"
            strInfoList = strInfoList & strProperty & "&nbsp;" & strLink
            strInfoList = strInfoList & strAuthor & strUpdateTime & strHits
            strInfoList = strInfoList & strHotSign & strNewSign & strCommentLink & strContent
            strInfoList = strInfoList & "</div>"

            iCount = iCount + 1
            If iCount Mod 2 = 0 Then
                CssName = CssName1
            Else
                CssName = CssName2
            End If
        Case 6 '���RSS
            strInfoList = strInfoList & GetInfoList_GetStrRSS(strTitle, strLink, strContent, strAuthor, strClassName, strUpdateTime)
            iCount = iCount + 1
        Case 4 '������ܶ���ʽ
            If TitleLen > 0 Then
                strTitle = ReplaceText(GetSubStr(rsInfoList("Title"), TitleLen, ShowSuspensionPoints), 2)
            Else
                strTitle = ReplaceText(rsInfoList("Title"), 2)
            End If
            iTitleLen = Charlong(strTitle)
            CurrentTitleLen = CurrentTitleLen + iTitleLen

            strLink = ""
            strLink = strLink & GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, CssNameA, strTitle, InfoUrl, rsInfoList("Title"), rsInfoList("Author"), rsInfoList("UpdateTime"))
             
            If ShowCommentLink = True And rsInfoList("ShowCommentLink") = True Then
                strLink = strLink & "&nbsp;<a href='" & ChannelUrl_ASPFile & "/Comment.asp?Action=ShowAll&ArticleID=" & rsInfoList("ArticleID") & "'>" & strComment & "</a>"
                CurrentTitleLen = CurrentTitleLen + 1 + Charlong(nohtml(strComment))
            End If
             
            If isfirst = True Then
                strInfoList = strInfoList & "<td valign='top' class='" & CssName & "'>" & strProperty & strLink
                rownum = rownum + 1
                If CurrentTitleLen > TitleLen + 1 Then
                    CurrentTitleLen = 0
                    If rownum > ArticleNum Then
                        strInfoList = strInfoList & "</td></tr>"
                        Exit Do
                    Else
                        strInfoList = strInfoList & "</td></tr><tr>"
                    End If
                    iCount = iCount + 1
                Else
                    isfirst = False
                    CurrentTitleLen = CurrentTitleLen + 1
                End If
                If iCount Mod 2 = 0 Then
                    CssName = CssName1
                Else
                    CssName = CssName2
                End If
            Else
                If CurrentTitleLen > TitleLen + 1 And outend = False Then
                    CurrentTitleLen = iTitleLen
                    If ShowCommentLink = True And rsInfoList("ShowCommentLink") = True Then
                        CurrentTitleLen = CurrentTitleLen + 1 + Charlong(nohtml(strComment))
                    End If
             
                    strInfoList = strInfoList & "</td></tr><tr>"
                    iCount = iCount + 1
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                    strInfoList = strInfoList & "<td valign='top' class='" & CssName & "'>" & strProperty & strLink
                    rownum = rownum + 1
                    If rownum > ArticleNum Then
                        If CurrentTitleLen >= TitleLen Then
                            strInfoList = strInfoList & "</td></tr>"
                            Exit Do
                        Else
                            outend = True
                        End If
                    End If
                Else
                    If CurrentTitleLen > TitleLen + 1 Then
                        strInfoList = strInfoList & "</td></tr>"
                        Exit Do
                    Else
                        strInfoList = strInfoList & "&nbsp;" & strLink
                        CurrentTitleLen = CurrentTitleLen + 1
                    End If
                End If
            End If
        End Select
        rsInfoList.MoveNext
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    If ShowType = 4 Then
        strInfoList = strInfoList & "</table>"	
    ElseIF ShowType = 2 Or (Cols > 1 and ShowType<>5) Then
        strInfoList = strInfoList & "</tr></table>"
    End If

    rsInfoList.Close
    Set rsInfoList = Nothing
    If ShowType = 6 And RssCodeType = False Then strInfoList = unicode(strInfoList)
    GetArticleList = strInfoList
End Function


Private Function GetInfoList_GetTitleLen(TitleLen, ShowIncludePic, ShowCommentLink, IncludePic, CommentLink)
    Dim iTitleLen
    If IncludePic > 0 And ShowIncludePic = True Then
        iTitleLen = TitleLen - 6
    Else
        iTitleLen = TitleLen
    End If
    If CommentLink = True And ShowCommentLink = True Then
        iTitleLen = iTitleLen - 4
    End If
    GetInfoList_GetTitleLen = iTitleLen
End Function


Private Function GetInfoList_GetStrIncludePic(IncludePic)
    Dim strIncludePic
    strIncludePic = ""
    Select Case IncludePic
    Case 1
        strIncludePic = strIncludePic & "<span class=""S_headline1"">" & ArticlePro1 & "</span>"
    Case 2
        strIncludePic = strIncludePic & "<span class=""S_headline2"">" & ArticlePro2 & "</span>"
    Case 3
        strIncludePic = strIncludePic & "<span class=""S_headline3"">" & ArticlePro3 & "</span>"
    Case 4
        strIncludePic = strIncludePic & "<span class=""S_headline4"">" & ArticlePro4 & "</span>"
    End Select
    GetInfoList_GetStrIncludePic = strIncludePic
End Function


'=================================================
'��������GetPicArticle
'��  �ã���ʾͼƬ����
'��  ����
'0        iChannelID ---- Ƶ��ID
'1        arrClassID ---- ��ĿID���飬0Ϊ������Ŀ
'2        IncludeChild ---- �Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'3        iSpecialID ---- ר��ID��0Ϊ�������£�����ר�����£������Ϊ����0����ֻ��ʾ��Ӧר�������
'4        ArticleNum ---- �����ʾ����ƪ����
'5        IsHot ---- �Ƿ�����������
'6        IsElite ---- �Ƿ����Ƽ�����
'7        DateNum ---- ���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�����
'8        OrderType ---- ����ʽ��1--������ID����2--������ID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'9        ShowType ---- ��ʾ��ʽ��1ΪͼƬ+����+���ݼ�飺�������У�2Ϊ��ͼƬ+���⣺�������У�+���ݼ�飺�������У�3ΪͼƬ+������+���ݼ�飺�������У����������У�4Ϊ���DIV��ʽ��5Ϊ���RSS��ʽ
'10       ImgWidth ---- ͼƬ���
'11       ImgHeight ---- ͼƬ�߶�
'12       TitleLen ---- ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ���⣻��Ϊ-1������ʾ��������
'13       ContentLen ---- ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ���ݼ��
'14       ShowTips ---- �Ƿ���ʾ���ߡ�����ʱ�䡢���������ʾ��Ϣ��TrueΪ��ʾ��FalseΪ����ʾ
'15       Cols ---- ÿ�е������������������ͻ��С�
'16       UrlType ---- ���ӵ�ַ���ͣ�0Ϊ���·����1Ϊ����ַ�ľ���·����
'=================================================
Public Function GetPicArticle(iChannelID, arrClassID, IncludeChild, iSpecialID, ArticleNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType)
    Dim sqlPic, rsPic, iCount, strPic, strLink, strAuthor, InfoUrl
    Dim strDefaultPicUrl, strLink_DefaultPicUrl, strTitle, strLink_Title, strContent, strLink_Content

    iCount = 0
    ArticleNum = PE_CLng(ArticleNum)
    ShowType = PE_CLng(ShowType)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)

    If ArticleNum < 0 Or ArticleNum >= 100 Then ArticleNum = 10
    If ShowType < 1 And ShowType > 5 Then ShowType = 2
    If ImgWidth < 0 Or ImgWidth > 1000 Then ImgWidth = 150
    If ImgHeight < 0 Or ImgHeight > 1000 Then ImgHeight = 150
    If ShowType = 5 Then UrlType = 1
    If Cols <= 0 Then Cols = 5

    FoundErr = False
    If (PE_Clng(iChannelID) <> 0 and Instr(iChannelID,",") = 0) and (PE_Clng(iChannelID)<>PrevChannelID Or ChannelID = 0) Then
        Call GetChannel(PE_Clng(iChannelID))
        PrevChannelID = iChannelID		 
    End If
    PrevChannelID = iChannelID
    If FoundErr = True Then
        GetPicArticle = ErrMsg
        Exit Function
    End If

    sqlPic = "select"
    If ArticleNum > 0 Then
        sqlPic = sqlPic & " top " & ArticleNum
    End If
    sqlPic = sqlPic & " A.ChannelID,A.ClassID,A.ArticleID,A.Title,A.TitleFontColor,A.TitleFontType,A.Author,A.UpdateTime,A.Hits,A.InfoPurview,A.InfoPoint,A.DefaultPicUrl"
    If ContentLen > 0 Then
        sqlPic = sqlPic & ",A.Intro,A.Content"
    End If
    sqlPic = sqlPic & ",C.ClassName,C.ClassDir,C.ParentDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    If ShowType < 4 Then strPic = "<table width='100%' cellpadding='0' cellspacing='5' border='0' align='center'><tr valign='top'>"
    If rsPic.BOF And rsPic.EOF Then
        If ArticleNum = 0 Then totalPut = 0
        If ShowType < 4 Then
            strPic = strPic & "<td align='center'><img class='pic1' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>" & R_XmlText_Class("PicArticle/NoFound", "û���κ�ͼƬ{$ChannelShortName}") & "</td></tr></table>"
        ElseIf ShowType = 4 Then
            strPic = strPic & "<div class=""pic_art""><img class=""pic1"" src=""" & strInstallDir & "images/nopic.gif"" width=""" & ImgWidth & """ height=""" & ImgHeight & """ border=""0""><br>" & R_XmlText_Class("PicArticle/NoFound", "û���κ�ͼƬ{$ChannelShortName}") & "</div>"
        End If
        
        rsPic.Close
        Set rsPic = Nothing
        GetPicArticle = strPic
        Exit Function
    End If
    If ArticleNum = 0 And ShowType < 5 Then
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
    PrevChannelID=0

    If ShowType = 5 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
    Do While Not rsPic.EOF
      '  If iChannelID = 0 Then
            If rsPic("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsPic("ChannelID"))
                PrevChannelID = rsPic("ChannelID")
            End If
        'End If

        ChannelUrl = UrlPrefix(UrlType, ChannelUrl) & ChannelUrl
        ChannelUrl_ASPFile = UrlPrefix(UrlType, ChannelUrl_ASPFile) & ChannelUrl_ASPFile
        If ShowType < 5 Then
            InfoUrl = GetArticleUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ArticleID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint"))
            strDefaultPicUrl = GetDefaultPicUrl(rsPic("DefaultPicUrl"), ImgWidth, ImgHeight)
            strLink_DefaultPicUrl = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strDefaultPicUrl, InfoUrl, rsPic("Title"), rsPic("Author"), rsPic("UpdateTime"))

            If ShowType = 4 Then
                strPic = strPic & "<div class=""pic_art"">" & vbCrLf
                strPic = strPic & "<div class=""pic_art_img"">" & strLink_DefaultPicUrl & "</div>" & vbCrLf
            Else
                strPic = strPic & "<td align='center'>"
                strPic = strPic & strLink_DefaultPicUrl
            End If

            If TitleLen <> 0 Then
                strTitle = GetInfoList_GetStrTitle(rsPic("Title"), TitleLen, rsPic("TitleFontType"), rsPic("TitleFontColor"))
                strLink_Title = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strTitle, InfoUrl, rsPic("Title"), rsPic("Author"), rsPic("UpdateTime"))
                Select Case PE_CLng(ShowType)
                Case 1, 2
                    strPic = strPic & "<br>" & strLink_Title
                Case 3
                    strPic = strPic & "</td><td valign='top' align='left'>" & strLink_Title
                Case 4
                    strPic = strPic & "<div class=""pic_art_title"">" & strLink_Title & "</div>" & vbCrLf
                End Select
            End If
            If ContentLen > 0 Then
                If Trim(rsPic("Intro") & "") = "" Then
                    strContent = Left(Replace(Replace(Replace(nohtml(rsPic("Content")), "[NextPage]", ""), ">", "&gt;"), "<", "&lt;"), ContentLen) & "����"
                Else
                    strContent = Left(rsPic("Intro"), ContentLen)
                End If
                strLink_Content = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strContent, InfoUrl, rsPic("Title"), rsPic("Author"), rsPic("UpdateTime"))
                Select Case PE_CLng(ShowType)
                Case 1, 3
                    strPic = strPic & "<br><div align='left'>" & strLink_Content & "</div>"
                Case 2
                    strPic = strPic & "</td><td valign='top' align='left'>" & strLink_Content
                Case 4
                    strPic = strPic & "<div class=""pic_art_content"">" & strLink_Content & "</div>" & vbCrLf
                End Select
            End If
            If ShowType = 4 Then
                strPic = strPic & "</div>" & vbCrLf
            Else
                strPic = strPic & "</td>"
            End If
        Else
            strTitle = GetInfoList_GetStrTitle(rsPic("Title"), TitleLen, 0, "")
            strLink = GetArticleUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ArticleID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint"))
            strAuthor = GetInfoList_GetStrAuthor_RSS(rsPic("Author"))
            If ContentLen > 0 Then
                If Trim(rsPic("Intro") & "") = "" Then
                    strContent = Left(Replace(Replace(Replace(xml_nohtml(rsPic("Content")), "[NextPage]", ""), ">", "&gt;"), "<", "&lt;"), ContentLen)
                Else
                    strContent = Left(xml_nohtml(rsPic("Intro")), ContentLen)
                End If
            End If
            strPic = strPic & GetInfoList_GetStrRSS(xml_nohtml(strTitle), strLink, strContent, strAuthor, xml_nohtml(rsPic("ClassName")), rsPic("UpdateTime"))
        End If
        rsPic.MoveNext
        iCount = iCount + 1
        If ArticleNum = 0 And iCount >= MaxPerPage Then Exit Do
        If ((iCount Mod Cols = 0) And (Not rsPic.EOF)) And ShowType < 4 Then strPic = strPic & "</tr><tr valign='top'>"
    Loop

    If ShowType < 4 Then strPic = strPic & "</tr></table>"
    rsPic.Close
    Set rsPic = Nothing
    If ShowType = 5 And RssCodeType = False Then strPic = unicode(strPic)
    GetPicArticle = strPic
End Function

'=================================================
'��������GetSlidePicArticle
'��  �ã��Իõ�ƬЧ����ʾͼƬ����
'��  ����
'0        iChannelID ---- Ƶ��ID
'1        arrClassID ---- ��ĿID���飬0Ϊ������Ŀ
'2        IncludeChild ---- �Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'3        iSpecialID ---- ר��ID��0Ϊ�������£�����ר�����£������Ϊ����0����ֻ��ʾ��Ӧר�������
'4        ArticleNum ---- �����ʾ����ƪ����
'5        IsHot ---- �Ƿ�����������
'6        IsElite ---- �Ƿ����Ƽ�����
'7        DateNum ---- ���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�����
'8        OrderType ---- ����ʽ��1--������ID����2--������ID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'9        ImgWidth ---- ͼƬ���
'10       ImgHeight ---- ͼƬ�߶�
'11       TitleLen ---- ���±����������ƣ�0Ϊ����ʾ��-1Ϊ��ʾ��������
'12       iTimeOut ---- Ч���任���ʱ�䣬�Ժ���Ϊ��λ
'13       effectID ---- ͼƬת��Ч����0��22ָ��ĳһ����Ч��23��ʾ���Ч��
'=================================================
Public Function GetSlidePicArticle(iChannelID, arrClassID, IncludeChild, iSpecialID, ArticleNum, IsHot, IsElite, DateNum, OrderType, ImgWidth, ImgHeight, TitleLen, iTimeOut, effectID)
    Dim sqlPic, rsPic, i, strPic
    Dim DefaultPicUrl, strTitle

    ArticleNum = PE_CLng(ArticleNum)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)

    If ArticleNum <= 0 Or ArticleNum > 100 Then ArticleNum = 10
    If ImgWidth < 0 Or ImgWidth > 1000 Then ImgWidth = 150
    If ImgHeight < 0 Or ImgHeight > 1000 Then ImgHeight = 150
    If iTimeOut < 1000 Or iTimeOut > 100000 Then iTimeOut = 5000
    If effectID < 0 Or effectID > 23 Then effectID = 23

    FoundErr = False
    If (PE_Clng(iChannelID) <> 0 and Instr(iChannelID,",") = 0) and (PE_Clng(iChannelID)<>PrevChannelID Or ChannelID = 0) Then
        Call GetChannel(PE_Clng(iChannelID))
        PrevChannelID = iChannelID		 
    End If	
    If FoundErr = True Then
        GetSlidePicArticle = ErrMsg
        Exit Function
    End If

    sqlPic = "select top " & ArticleNum & " A.ChannelID,A.ClassID,A.ArticleID,A.Title,A.UpdateTime,A.InfoPurview,A.InfoPoint,A.DefaultPicUrl"
    sqlPic = sqlPic & ",C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)


    Dim ranNum
    Randomize
    ranNum = Int(900 * Rnd) + 100
    strPic = "<script language=JavaScript>" & vbCrLf
    strPic = strPic & "<!--" & vbCrLf
    strPic = strPic & "var SlidePic_" & ranNum & " = new SlidePic_Article(""SlidePic_" & ranNum & """);" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Width    = " & ImgWidth & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Height   = " & ImgHeight & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".TimeOut  = " & iTimeOut & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Effect   = " & effectID & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".TitleLen = " & TitleLen & ";" & vbCrLf
    PrevChannelID=0

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    Do While Not rsPic.EOF
        'If iChannelID = 0 Then
            If rsPic("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsPic("ChannelID"))
                PrevChannelID = rsPic("ChannelID")
            End If
        'End If
        If Left(rsPic("DefaultPicUrl"), 1) <> "/" And InStr(rsPic("DefaultPicUrl"), "://") <= 0 Then
            DefaultPicUrl = ChannelUrl & "/" & UploadDir & "/" & rsPic("DefaultPicUrl")
        Else
            DefaultPicUrl = rsPic("DefaultPicUrl")
        End If
        If TitleLen = -1 Then
            strTitle = rsPic("Title")
        Else
            strTitle = GetSubStr(rsPic("Title"), TitleLen, ShowSuspensionPoints)
        End If
        
        strPic = strPic & "var oSP = new objSP_Article();" & vbCrLf
        strPic = strPic & "oSP.ImgUrl         = """ & DefaultPicUrl & """;" & vbCrLf
        strPic = strPic & "oSP.LinkUrl        = """ & GetArticleUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ArticleID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint")) & """;" & vbCrLf
        strPic = strPic & "oSP.Title         = """ & strTitle & """;" & vbCrLf
        strPic = strPic & "SlidePic_" & ranNum & ".Add(oSP);" & vbCrLf
        
        rsPic.MoveNext
    Loop
    strPic = strPic & "SlidePic_" & ranNum & ".Show();" & vbCrLf
    strPic = strPic & "//-->" & vbCrLf
    strPic = strPic & "</script>" & vbCrLf
    
    rsPic.Close
    Set rsPic = Nothing
    GetSlidePicArticle = strPic
End Function

Private Function JS_SlidePic()
    Dim strJS, LinkTarget
    LinkTarget = XmlText_Class("SlidePicArticle/LinkTarget", "_blank")
    strJS = strJS & "<script language=""JavaScript"">" & vbCrLf
    strJS = strJS & "<!--" & vbCrLf
    strJS = strJS & "var navigatorName = ""Microsoft Internet Explorer"";" & vbCrLf
    strJS = strJS & "var isIE = false; " & vbCrLf		
    strJS = strJS & "if(navigator.appName==navigatorName) isIE = true;" & vbCrLf	
    strJS = strJS & "function objSP_Article() {this.ImgUrl=""""; this.LinkUrl=""""; this.Title="""";}" & vbCrLf
    strJS = strJS & "function SlidePic_Article(_id) {this.ID=_id; this.Width=0;this.Height=0; this.TimeOut=5000; this.Effect=23; this.TitleLen=0; this.PicNum=-1; this.Img=null; this.Url=null; this.Title=null; this.AllPic=new Array(); this.Add=SlidePic_Article_Add; this.Show=SlidePic_Article_Show; this.LoopShow=SlidePic_Article_LoopShow;}" & vbCrLf
    strJS = strJS & "function SlidePic_Article_Add(_SP) {this.AllPic[this.AllPic.length] = _SP;}" & vbCrLf
    strJS = strJS & "function SlidePic_Article_Show() {" & vbCrLf
    strJS = strJS & "  if(this.AllPic[0] == null) return false;" & vbCrLf
    strJS = strJS & "  document.write(""<div align='center'><a id='Url_"" + this.ID + ""' href='' target='" & LinkTarget & "'><img id='Img_"" + this.ID + ""' style='width:"" + this.Width + ""px; height:"" + this.Height + ""px; filter: revealTrans(duration=2,transition=23);' src='javascript:null' border='0'></a>"");" & vbCrLf
    strJS = strJS & "  if(this.TitleLen != 0) {document.write(""<br><span id='Title_"" + this.ID + ""'></span></div>"");}" & vbCrLf
    strJS = strJS & "  else{document.write(""</div>"");}" & vbCrLf
    strJS = strJS & "  this.Img = document.getElementById(""Img_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Url = document.getElementById(""Url_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Title = document.getElementById(""Title_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.LoopShow();" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function SlidePic_Article_LoopShow() {" & vbCrLf
    strJS = strJS & "  if(this.PicNum<this.AllPic.length-1) this.PicNum++ ; " & vbCrLf
    strJS = strJS & "  else this.PicNum=0; " & vbCrLf
    strJS = strJS & "  if(isIE==true){" & vbCrLf	
    strJS = strJS & "  this.Img.filters.revealTrans.Transition=this.Effect; " & vbCrLf
    strJS = strJS & "  this.Img.filters.revealTrans.apply(); " & vbCrLf
    strJS = strJS & "  }" & vbCrLf		
    strJS = strJS & "  this.Img.src=this.AllPic[this.PicNum].ImgUrl;" & vbCrLf
    strJS = strJS & "  if(isIE==true){" & vbCrLf		
    strJS = strJS & "  this.Img.filters.revealTrans.play();" & vbCrLf
    strJS = strJS & "  }" & vbCrLf			
    strJS = strJS & "  this.Url.href=this.AllPic[this.PicNum].LinkUrl;" & vbCrLf
    strJS = strJS & "  if(this.Title) this.Title.innerHTML=""<a href=""+this.AllPic[this.PicNum].LinkUrl+"" target='" & LinkTarget & "'>""+this.AllPic[this.PicNum].Title+""</a>"";" & vbCrLf
    strJS = strJS & "  this.Img.timer=setTimeout(this.ID+"".LoopShow()"",this.TimeOut);" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "//-->" & vbCrLf
    strJS = strJS & "</script>" & vbCrLf
    JS_SlidePic = strJS
End Function

Private Function GetDefaultPicUrl(ByVal DefaultPicUrl, ByVal DefaultPicWidth, ByVal DefaultPicHeight)
    Dim strUrl, FileType, strPicUrl
    If DefaultPicUrl = "" Or IsNull(DefaultPicUrl) = True Then
        strUrl = strUrl & "<img src='" & strPicUrl & strInstallDir & "images/nopic.gif' "
        If DefaultPicWidth > 0 Then strUrl = strUrl & " width='" & DefaultPicWidth & "'"
        If DefaultPicHeight > 0 Then strUrl = strUrl & " height='" & DefaultPicHeight & "'"
        strUrl = strUrl & " border='0'>"
    Else
        FileType = LCase(Mid(DefaultPicUrl, InStrRev(DefaultPicUrl, ".") + 1))
        If Left(DefaultPicUrl, 1) <> "/" And InStr(DefaultPicUrl, "://") <= 0 Then
            strPicUrl = ChannelUrl & "/" & UploadDir & "/" & DefaultPicUrl
        Else
            strPicUrl = DefaultPicUrl
        End If
        Select Case FileType
        Case "swf"
            strUrl = strUrl & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' "
            If DefaultPicWidth > 0 Then strUrl = strUrl & " width='" & DefaultPicWidth & "'"
            If DefaultPicHeight > 0 Then strUrl = strUrl & " height='" & DefaultPicHeight & "'"
            strUrl = strUrl & "><param name='movie' value='" & strPicUrl & "'><param name='quality' value='high'><embed src='" & strPicUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' "
            If DefaultPicWidth > 0 Then strUrl = strUrl & " width='" & DefaultPicWidth & "'"
            If DefaultPicHeight > 0 Then strUrl = strUrl & " height='" & DefaultPicHeight & "'"
            strUrl = strUrl & "></embed></object>"
        Case "gif", "jpg", "jpeg", "jpe", "bmp", "png"
            strUrl = strUrl & "<img class='pic1' src='" & strPicUrl & "' "
            If DefaultPicWidth > 0 Then strUrl = strUrl & " width='" & DefaultPicWidth & "'"
            If DefaultPicHeight > 0 Then strUrl = strUrl & " height='" & DefaultPicHeight & "'"
            strUrl = strUrl & " border='0'>"
        Case Else
            strUrl = strUrl & "<img class='pic1' src='" & strInstallDir & "images/nopic.gif' "
            If DefaultPicWidth > 0 Then strUrl = strUrl & " width='" & DefaultPicWidth & "'"
            If DefaultPicHeight > 0 Then strUrl = strUrl & " height='" & DefaultPicHeight & "'"
            strUrl = strUrl & " border='0'>"
        End Select
    End If
    GetDefaultPicUrl = strUrl
End Function


Private Function GetSearchResultIDArr(iChannelID)
    Dim sqlSearch, rsSearch
    Dim rsField
    Dim ArticleNum, arrArticleID

    If PE_CLng(SearchResultNum) > 0 Then
        sqlSearch = "select top " & PE_CLng(SearchResultNum) & " ArticleID "
    Else
        sqlSearch = "select ArticleID "
    End If
    sqlSearch = sqlSearch & " from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0"
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
        sqlSearch = sqlSearch & " and ArticleID in (select ItemID from PE_InfoS where SpecialID=" & SpecialID & ")"
    End If
    If strField <> "" Then  '��ͨ����
        Select Case strField
            Case "Title"
                sqlSearch = sqlSearch & SetSearchString("Title")
            Case "ArticleID"
                sqlSearch = sqlSearch & " and ArticleID = "&PE_Clng(Keyword)			
            Case "Content"
                sqlSearch = sqlSearch & SetSearchString("Content")
            Case "Author"
                sqlSearch = sqlSearch & SetSearchString("Author")
            Case "Inputer"
                sqlSearch = sqlSearch & SetSearchString("Inputer")
            Case "Editor"
                sqlSearch = sqlSearch & SetSearchString("Editor")
            Case "Keywords"
                sqlSearch = sqlSearch & SetSearchString("Keyword")
            Case Else  '�Զ����ֶ�
                Set rsField = Conn.Execute("select Title from PE_Field where (ChannelID=-1 or ChannelID=" & iChannelID & ") and FieldName='" & ReplaceBadChar(strField) & "'")
                If rsField.BOF And rsField.EOF Then
                    sqlSearch = sqlSearch & SetSearchString("Title")
                Else
                    sqlSearch = sqlSearch & SetSearchString(ReplaceBadChar(strField))
                End If
                rsField.Close
                Set rsField = Nothing
        End Select
    Else   '�߼�����
        '����߼���������
        Dim Title, Intro, Content, Author, CopyFrom, Keyword2, LowInfoPoint, HighInfoPoint, BeginDate, EndDate, Inputer, ArticleID
        Title = Trim(Request("Title"))
        ArticleID = PE_Clng(Trim(Request("ArticleID")))		
        Content = Trim(Request("Content"))
        Intro = Trim(Request("Intro"))
        Author = Trim(Request("Author"))
        CopyFrom = Trim(Request("CopyFrom"))
        Keyword2 = Trim(Request("Keywords"))
        LowInfoPoint = PE_CLng(Request("LowInfoPoint"))
        HighInfoPoint = PE_CLng(Request("HighInfoPoint"))
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        Inputer = Trim(Request("Inputer"))
        strFileName = "Search.asp?ModuleName=Article&ClassID=" & ClassID & "&SpecialID=" & SpecialID
        If Title <> "" Then
            Title = ReplaceBadChar(Title)
            strFileName = strFileName & "&Title=" & Title
            sqlSearch = sqlSearch & " and Title like '%" & Title & "%' "
        End If
        If Content <> "" Then
            Content = ReplaceBadChar(Content)
            strFileName = strFileName & "&Content=" & Content
            sqlSearch = sqlSearch & " and Content like '%" & Content & "%'"
        End If
        If ArticleID <> 0 Then
            ArticleID = ReplaceBadChar(ArticleID)
            strFileName = strFileName & "&ArticleID=" & ArticleID
            sqlSearch = sqlSearch & " and ArticleID =" & ArticleID
        End If		
        If Intro <> "" Then
            Intro = ReplaceBadChar(Intro)
            strFileName = strFileName & "&Intro=" & Intro
            sqlSearch = sqlSearch & " and Intro like '%" & Intro & "%'"
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

        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            If Trim(Request(rsField("FieldName"))) <> "" Then
                strFileName = strFileName & "&" & Trim(rsField("FieldName")) & "=" & ReplaceBadChar(Trim(Request(rsField("FieldName"))))
                sqlSearch = sqlSearch & " and " & Trim(rsField("FieldName")) & " like '%" & ReplaceBadChar(Trim(Request(rsField("FieldName")))) & "%' "
            End If
            rsField.MoveNext
        Loop
        Set rsField = Nothing
        
    End If
    sqlSearch = sqlSearch & " order by ArticleID desc"

    arrArticleID = ""
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
        ArticleNum = 0
        Do While Not rsSearch.EOF
            If arrArticleID = "" Then
                arrArticleID = rsSearch(0)
            Else
                arrArticleID = arrArticleID & "," & rsSearch(0)
            End If
            ArticleNum = ArticleNum + 1
            If ArticleNum >= MaxPerPage Then Exit Do
            rsSearch.MoveNext
        Loop
    End If
    rsSearch.Close
    Set rsSearch = Nothing
    GetSearchResultIDArr = arrArticleID
End Function

'=================================================
'��������GetSearchResult
'��  �ã���ҳ��ʾ�������
'��  ������
'=================================================
Private Function GetSearchResult(iChannelID)
    Dim sqlSearch, Intro, rsSearch, iCount, ArticleNum, arrArticleID, strSearchResult, Content
    strSearchResult = ""
    arrArticleID = GetSearchResultIDArr(iChannelID)
    
    
    If arrArticleID = "" Then
        GetSearchResult = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "û�л�û���ҵ��κ�{$ChannelShortName}") & "<br><br></p>"
        Exit Function
    End If

    ArticleNum = 1
    sqlSearch = "select A.ChannelID,A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,A.Intro,A.InfoPurview,A.InfoPoint,A.Content,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where ArticleID in (" & arrArticleID & ") order by ArticleID desc"
    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1
    Do While Not rsSearch.EOF
        If iChannelID = 0 Then
            If rsSearch("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsSearch("ChannelID"))
                PrevChannelID = rsSearch("ChannelID")
            End If
        End If
        
        strSearchResult = strSearchResult & "<b>" & CStr(MaxPerPage * (CurrentPage - 1) + ArticleNum) & ".</b> "
        
        strSearchResult = strSearchResult & "[<a class='LinkSearchResult' href='" & GetClassUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("ClassID"), rsSearch("ClassPurview")) & "' target='_blank'>" & rsSearch("ClassName") & "</a>] "
        
        strSearchResult = strSearchResult & "<a class='LinkSearchResult' href='" & GetArticleUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("UpdateTime"), rsSearch("ArticleID"), rsSearch("ClassPurview"), rsSearch("InfoPurview"), rsSearch("InfoPoint")) & "' target='_blank'>"
        
        If strField = "Title" Then
            strSearchResult = strSearchResult & "<b>" & Replace(ReplaceText(rsSearch("Title"), 2) & "", "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "</b>"
        Else
            strSearchResult = strSearchResult & "<b>" & ReplaceText(rsSearch("Title"), 2) & "</b>"
        End If
        strSearchResult = strSearchResult & "</a>"
        If strField = "Author" Then
            strSearchResult = strSearchResult & "&nbsp;[" & Replace(rsSearch("Author") & "", "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "]"
        Else
            strSearchResult = strSearchResult & "&nbsp;[" & rsSearch("Author") & "]"
        End If
        strSearchResult = strSearchResult & "[" & FormatDateTime(rsSearch("UpdateTime"), 1) & "][" & rsSearch("Hits") & "]"
        strSearchResult = strSearchResult & "<br>"
        
        If rsSearch("Intro") <> "" Then
            Intro = "��飺" & Replace(Replace(ReplaceText(nohtml(rsSearch("Intro")), 1), ">", "&gt;"), "<", "&lt;") & "<br>"
        Else
            Intro = "��飺��<br>"
        End If
        If rsSearch("ClassPurview") > 0 Or rsSearch("InfoPoint") > 0 Then
            strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & SearchResult_Content_NoPurview & "</div>"
        Else
            Content = Intro & "���ݣ�<br>" & Left(Replace(Replace(ReplaceText(nohtml(rsSearch("content")), 1), ">", "&gt;"), "<", "&lt;"), SearchResult_ContentLenth)
            If strField = "Content" Then
                strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & Replace(Content, "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "����</div>"
            Else
                strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & Content & "����</div>"
            End If
        End If
        strSearchResult = strSearchResult & "<br>"
        ArticleNum = ArticleNum + 1
        rsSearch.MoveNext
    Loop
    rsSearch.Close
    Set rsSearch = Nothing
    GetSearchResult = strSearchResult
End Function


Public Function GetSearchResult2(iChannelID, strValue)   '�õ��Զ����б�İ�����Ƶ�HTML����
    Dim strCustom, strParameter
    strCustom = strValue
    regEx.Pattern = "��SearchResultList\((.*?)\)��([\s\S]*?)��\/SearchResultList��"
    Set Matches = regEx.Execute(strCustom)
    For Each Match In Matches
        strParameter = Replace(Match.SubMatches(0), Chr(34), " ")
        strCustom = PE_Replace(strCustom, Match.Value, GetSearchResultLabel(strParameter, Match.SubMatches(1), iChannelID))
    Next
    GetSearchResult2 = strCustom
End Function


'�����Զ����ǩ
Private Function GetSearchResultLabel(strTemp, strList, iChannelID)
    Dim sqlSearch, rsSearch, rsCustom, iCount, arrArticleID
    Dim arrTemp, strCustomList
    Dim strArticlePic, strPicTemp, arrPicTemp, UsePage
    Dim IncludeChild, iSpecialID, IsHot, IsElite, DateNum, OrderType, TitleLen, ContentLen
    Dim iCols, iColsHtml, iRows, iRowsHtml, iNumber
    Dim rsField, ArrField, iField

    If strTemp = "" Or strList = "" Then GetSearchResultLabel = "": Exit Function
    iCols = 1: iRows = 1: iColsHtml = "": iRowsHtml = ""
        
    regEx.Pattern = "��(Cols|Rows)=(\d{1,2})\s*(?:\||��)(.+?)��"
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
        GetSearchResultLabel = "�Զ����б��ǩ����SearchResultList(�����б�)���б����ݡ�/SearchResultList���Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If
    
    TitleLen = arrTemp(0)
    UsePage = arrTemp(1)
    ContentLen = arrTemp(2)
    
    arrArticleID = GetSearchResultIDArr(iChannelID)

    If arrArticleID = "" Then
        GetSearchResultLabel = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "û�л�û���ҵ��κ�{$ChannelShortName}") & "<br><br></p>"
        Exit Function
    End If

    Set rsField = Conn.Execute("select FieldName,LabelName from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
    If Not (rsField.BOF And rsField.EOF) Then
        ArrField = rsField.getrows(-1)
    End If
    Set rsField = Nothing
    
    sqlSearch = "select A.ChannelID,A.ArticleID,A.Title,A.Subheading,"
    If IsArray(ArrField) Then
     For iField = 0 To UBound(ArrField, 2)
         sqlSearch = sqlSearch & "A." & ArrField(0, iField) & ","
     Next
    End If
    
    iCount = 0
    strCustomList = ""
    sqlSearch = sqlSearch & "A.Author,A.Keyword,A.CopyFrom,A.DefaultPicUrl,A.InfoPoint,A.Editor,A.OnTop,A.UpdateTime,A.Hits,A.Elite,A.Intro,A.Inputer,A.InfoPurview,A.Content,A.Stars,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,C.ReadMe from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where ArticleID in (" & arrArticleID & ") order by ArticleID desc"
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

        strTemp = PE_Replace(strTemp, "{$ArticleID}", rsCustom("ArticleID"))
        If InStr(strTemp, "{$ArticleUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ArticleUrl}", GetArticleUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("ArticleID"), rsCustom("ClassPurview"), rsCustom("InfoPurview"), rsCustom("InfoPoint")))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        strTemp = PE_Replace(strTemp, "{$Author}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$CopyFrom}", rsCustom("CopyFrom"))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$Inputer}", rsCustom("Inputer"))
        strTemp = PE_Replace(strTemp, "{$Editor}", rsCustom("Editor"))
        If InStr(strTemp, "{$InfoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$InfoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$ReadPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$ReadPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$Keyword}") > 0 Then strTemp = PE_Replace(strTemp, "{$Keyword}", GetKeywords(",", rsCustom("Keyword")))
        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "OnTop")
        ElseIf rsCustom("Elite") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Elite")
        ElseIf rsCustom("Hits") > HitsOfHot Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Hot")
        Else
            strTemp = PE_Replace(strTemp, "{$Property}", "Common")
        End If

        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Top}", strTop2)
        Else
            strTemp = PE_Replace(strTemp, "{$Top}", "")
        End If
        If rsCustom("Elite") = True Then
            strTemp = PE_Replace(strTemp, "{$Elite}", strElite2)
        Else
            strTemp = PE_Replace(strTemp, "{$Elite}", "")
        End If
        If rsCustom("Hits") > HitsOfHot Then
            strTemp = PE_Replace(strTemp, "{$Hot}", strHot2)
        Else
            strTemp = PE_Replace(strTemp, "{$Hot}", "")
        End If
        
        If TitleLen > 0 Then
            strTemp = PE_Replace(strTemp, "{$Title}", GetSubStr(rsCustom("Title"), TitleLen, ShowSuspensionPoints))
        Else
            strTemp = PE_Replace(strTemp, "{$Title}", rsCustom("Title"))
        End If
        strTemp = PE_Replace(strTemp, "{$TitleOriginal}", rsCustom("Title"))

        If ContentLen > 0 Then
            If InStr(strTemp, "{$Content}") > 0 Then strTemp = PE_Replace(strTemp, "{$Content}", Left(nohtml(rsCustom("Content")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$Content}", "")
        End If
        strTemp = PE_Replace(strTemp, "{$Subheading}", rsCustom("Subheading"))
        strTemp = PE_Replace(strTemp, "{$Intro}", rsCustom("Intro"))
        
        '�滻��ҳͼƬ
        regEx.Pattern = "\{\$ArticlePic\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strArticlePic = GetDefaultPicUrl(Trim(rsCustom("DefaultPicUrl")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)))
            strTemp = Replace(strTemp, Match.Value, strArticlePic)
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
'��������GetCorrelative
'��  �ã���ʾ�������
'��  ����ArticleNum  ----�����ʾ����ƪ����
'        TitleLen   ----��������ַ�����һ������=����Ӣ���ַ�
'        OrderType ---- ����ʽ��1--������ID����2--������ID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'        OpenType ---- ���´򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�
'        Cols ---- ÿ�е������������������ͻ��С�
'=================================================
Private Function GetCorrelative(iChannelID, arrClassID, MaxNum, ArticleNum, TitleLen, OrderType, OpenType, Cols, ShowClassName)
    Dim rsCorrelative, sqlCorrelative, strCorrelative, iCols, iTemp
    Dim strKey, arrKey, i
    iChannelID = Replace(iChannelID,"|",",")
    Select Case iChannelID
    Case "ChannelID"
        iChannelID = ChannelID
    Case else
        If IsValidID(iChannelID) = False Then
            iChannelID = 0
        End If  
    End Select	
    arrClassID = Replace(arrClassID,"|",",")	
    If IsValidID(arrClassID) = False Then
        arrClassID = 0
    End If  
    MaxNum = PE_Clng(MaxNum)
    iTemp = 1
    If PE_CLng(Cols) <> 0 Then
        iCols = PE_CLng(Cols)
    Else
        iCols = 1
    End If

    If ArticleNum > 0 And ArticleNum <= 100 Then
        sqlCorrelative = "select top " & ArticleNum
    Else
        sqlCorrelative = "Select Top 5 "
    End If
    strKey = Mid(rsArticle("Keyword"), 2, Len(rsArticle("Keyword")) - 2)
    If InStr(strKey, "|") > 1 Then
        arrKey = Split(strKey, "|")
        If  MaxNum > UBound(arrKey) Then
            MaxNum = UBound(arrKey)   
        End IF	
        If MaxNum > 4 Then MaxNum = 4
        strKey = "((A.Keyword like '%|" & Replace(Replace(arrKey(0), "��", ""), "��", "") & "|%')"
        For i = 1 To MaxNum
            strKey = strKey & " or (A.Keyword like '%|" & Replace(Replace(arrKey(i), "��", ""), "��", "") & "|%')"
        Next
        strKey = strKey & ")"
    Else
        strKey = "(A.Keyword like '%|" & strKey & "|%')"
    End If
    sqlCorrelative = sqlCorrelative & " A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,A.InfoPurview,A.InfoPoint,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where 1=1"
    If InStr(iChannelID, ",") > 0 Then
        sqlCorrelative = sqlCorrelative & " and A.ChannelID in (" & FilterArrNull(iChannelID, ",") & ")"
    Else
        If PE_CLng(iChannelID) > 0 Then sqlCorrelative = sqlCorrelative & " and A.ChannelID=" & PE_CLng(iChannelID)
    End If	
    If arrClassID <> "0" Then
        If InStr(arrClassID, ",") > 0 Then
            sqlCorrelative = sqlCorrelative & " and A.ClassID in (" & FilterArrNull(arrClassID, ",") & ")"
        Else
            If PE_CLng(arrClassID) > 0 Then sqlCorrelative = sqlCorrelative & " and A.ClassID=" & PE_CLng(arrClassID)
        End If
    End If
    sqlCorrelative = sqlCorrelative & " and A.Deleted=" & PE_False & " and A.Status=3 and A.ReceiveType=0"

    sqlCorrelative = sqlCorrelative & " and " & strKey & " and A.ArticleID<>" & ArticleID & " Order by "
    Select Case PE_CLng(OrderType)
    Case 1
        sqlCorrelative = sqlCorrelative & "A.ArticleID desc"
    Case 2
        sqlCorrelative = sqlCorrelative & "A.ArticleID asc"
    Case 3
        sqlCorrelative = sqlCorrelative & "A.UpdateTime desc"
    Case 4
        sqlCorrelative = sqlCorrelative & "A.UpdateTime asc"
    Case 5
        sqlCorrelative = sqlCorrelative & "A.Hits desc"
    Case 6
        sqlCorrelative = sqlCorrelative & "A.Hits asc"
    Case 7
        sqlCorrelative = sqlCorrelative & "A.CommentCount desc"
    Case 8
        sqlCorrelative = sqlCorrelative & "A.CommentCount asc"
    Case Else
        sqlCorrelative = sqlCorrelative & "A.ArticleID desc"
    End Select
    Set rsCorrelative = Conn.Execute(sqlCorrelative)
    If TitleLen < 0 Or TitleLen > 255 Then TitleLen = 50
    If rsCorrelative.BOF And rsCorrelative.EOF Then
        strCorrelative = R_XmlText_Class("ShowArticle/NoCorrelative", "û�����{$ChannelShortName}")
    Else
        Do While Not rsCorrelative.EOF
            If PE_CBool(ShowClassName) = True Then 
                strCorrelative = strCorrelative & GetInfoList_GetStrClassLink(Character_Class,"", rsCorrelative("ClassID"), rsCorrelative("ClassName"), GetClassUrl(rsCorrelative("ParentDir"), rsCorrelative("ClassDir"), rsCorrelative("ClassID"), rsCorrelative("ClassPurview")))
            End If
            strCorrelative = strCorrelative & "<a class='LinkArticleCorrelative' href='" & GetArticleUrl(rsCorrelative("ParentDir"), rsCorrelative("ClassDir"), rsCorrelative("UpdateTime"), rsCorrelative("ArticleID"), rsCorrelative("ClassPurview"), rsCorrelative("InfoPurview"), rsCorrelative("InfoPoint")) & "'"
            strCorrelative = strCorrelative & " title='" & Replace(Replace(Replace(Replace(strList_Title, "{$Title}", rsCorrelative("Title")), "{$Author}", rsCorrelative("Author")), "{$UpdateTime}", rsCorrelative("UpdateTime")), "{$br}", vbCrLf)
            If OpenType = 0 Then
                strCorrelative = strCorrelative & "' target=""_self"">"
            Else
                strCorrelative = strCorrelative & "' target=""_blank"">"
            End If
            strCorrelative = strCorrelative & GetSubStr(rsCorrelative("Title"), TitleLen, ShowSuspensionPoints) & "</a>"
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

'=================================================
'��������GetPrevArticle
'��  �ã���ʾ��һƪ����
'��  ����TitleLen   ----��������ַ�����һ������=����Ӣ���ַ�
'=================================================
Private Function GetPrevArticle(TitleLen)
    Dim rsPrev, sqlPrev, strPrev
    strPrev = Replace(XmlText_Class("ShowArticle/PrevArticle_Link", "<li>��һ{$ItemUnit}�� "), "{$ItemUnit}", ChannelItemUnit & ChannelShortName)
    sqlPrev = "Select Top 1 ArticleID,Title,Author,UpdateTime,Hits,InfoPurview,InfoPoint from PE_Article Where ChannelID=" & ChannelID & " and Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and ClassID=" & rsArticle("ClassID") & " and ArticleID<" & rsArticle("ArticleID") & " order by ArticleID DESC"
    Set rsPrev = Conn.Execute(sqlPrev)
    If TitleLen < 0 Or TitleLen > 255 Then TitleLen = 50
    If rsPrev.EOF Then
        strPrev = strPrev & XmlText_Class("ShowArticle/NoPrevArticle", "û����")
    Else
        strPrev = strPrev & "<a class='LinkPrevArticle' href='" & GetArticleUrl(ParentDir, ClassDir, rsPrev("UpdateTime"), rsPrev("ArticleID"), ClassPurview, rsPrev("InfoPurview"), rsPrev("InfoPoint")) & "'"
        strPrev = strPrev & " title='" & Replace(Replace(Replace(Replace(strList_Title, "{$Title}", rsPrev("Title")), "{$Author}", rsPrev("Author")), "{$UpdateTime}", rsPrev("UpdateTime")), "{$br}", vbCrLf) & "'>" & GetSubStr(rsPrev("Title"), TitleLen, ShowSuspensionPoints) & "</a>"
    End If
    rsPrev.Close
    Set rsPrev = Nothing
    strPrev = strPrev & "</li>"
    GetPrevArticle = strPrev
End Function


'=================================================
'��������GetCurArticleUrl
'��  �ã���õ�ǰ���µ����ӵ�ַ
'=================================================
Private Function GetCurArticleUrl()
    Dim  strCur
    strCur = GetArticleUrl(ParentDir, ClassDir, rsArticle("UpdateTime"), rsArticle("ArticleID"), ClassPurview, rsArticle("InfoPurview"), rsArticle("InfoPoint")) 
    GetCurArticleUrl = strCur
End Function

'=================================================
'��������GetNextArticle
'��  �ã���ʾ��һƪ����
'��  ����TitleLen   ----��������ַ�����һ������=����Ӣ���ַ�
'=================================================
Private Function GetNextArticle(TitleLen)
    Dim rsNext, sqlNext, strNext
    strNext = Replace(XmlText_Class("ShowArticle/NextArticle_Link", "<li>��һ{$ItemUnit}�� "), "{$ItemUnit}", ChannelItemUnit & ChannelShortName)
    sqlNext = "Select Top 1 ArticleID,Title,Author,UpdateTime,Hits,InfoPurview,InfoPoint from PE_Article Where ChannelID=" & ChannelID & " and Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and ClassID=" & rsArticle("ClassID") & " and ArticleID>" & rsArticle("ArticleID") & " order by ArticleID ASC"
    Set rsNext = Conn.Execute(sqlNext)
    If TitleLen < 0 Or TitleLen > 255 Then TitleLen = 50
    If rsNext.EOF Then
        strNext = strNext & XmlText_Class("ShowArticle/NoNextArticle", "û����")
    Else
        strNext = strNext & "<a class='LinkNextArticle' href='" & GetArticleUrl(ParentDir, ClassDir, rsNext("UpdateTime"), rsNext("ArticleID"), ClassPurview, rsNext("InfoPurview"), rsNext("InfoPoint")) & "'"
        strNext = strNext & " title='" & Replace(Replace(Replace(Replace(strList_Title, "{$Title}", rsNext("Title")), "{$Author}", rsNext("Author")), "{$UpdateTime}", rsNext("UpdateTime")), "{$br}", vbCrLf) & "'>" & GetSubStr(rsNext("Title"), TitleLen, ShowSuspensionPoints) & "</a>"
    End If
    rsNext.Close
    Set rsNext = Nothing
    strNext = strNext & "</li>"
    GetNextArticle = strNext
End Function


'=================================================
'��������GetArticleContent
'��  �ã���ʾ���¾�������ݣ����Է�ҳ��ʾ
'��  ����PreviewContentLength   ----Ԥ�����ݵĳ���
'=================================================
Public Function GetArticleContent(PreviewContentLength)
    Dim FoundErr, ErrMsg, PurviewChecked, ConsumePoint
    FoundErr = False
    ErrMsg = ""

    If ClassPurview > 0 Or rsArticle("InfoPurview") > 0 Or rsArticle("InfoPoint") > 0 Then
        Dim ErrMsg_NoLogin, ErrMsg_PurviewCheckedErr, ErrMsg_PurviewCheckedErr2, ErrMsg_NoMail, ErrMsg_NoCheck, ErrMsg_NeedPoint, ErrMsg_UsePoint, ErrMsg_OutTime, ErrMsg_Overflow_Total, ErrMsg_Overflow_Today
        ErrMsg_NoLogin = Replace(Replace(Replace(R_XmlText_Class("ArticleContent/Nologin", "<br>&nbsp;&nbsp;&nbsp;&nbsp;�㻹ûע�᣿����û�е�¼����{$ItemUnit}Ҫ�������Ǳ�վ��ע���Ա�����Ķ���<br><br>&nbsp;&nbsp;&nbsp;&nbsp;����㻹ûע�ᣬ��Ͻ�<a href='{$InstallDir}Reg/User_Reg.asp'><font color=red>���ע��</font></a>�ɣ�<br><br>&nbsp;&nbsp;&nbsp;&nbsp;������Ѿ�ע�ᵫ��û��¼����Ͻ�<a href='{$InstallDir}User/User_Login.asp'><font color=red>��˵�¼</font></a>�ɣ�<br><br>"), "{$ItemUnit}", ChannelItemUnit & ChannelShortName), "{$ChannelItemUnit}", ChannelItemUnit), "{$InstallDir}", strInstallDir)
        If UserLogined <> True Then
            FoundErr = True
            ErrMsg = ErrMsg & ErrMsg_NoLogin
        Else
            Call GetUser(UserName)
            ErrMsg_PurviewCheckedErr = XmlText("BaseText", "PurviewCheckedErr", "<li>�Բ�����û�в鿴����Ŀ���ݵ�Ȩ�ޣ�</li>")
            ErrMsg_PurviewCheckedErr2 = XmlText("BaseText", "PurviewCheckedErr2", "<li>�Բ�����û�в鿴����Ϣ��Ȩ�ޣ�</li>")
            ErrMsg_NoMail = "<li>" & R_XmlText_Class("ArticleContent/NoMail", "�Բ�������δͨ���ʼ���֤�����ܲ鿴��{$ChannelShortName}") & "</li>"
            ErrMsg_NoCheck = "<li>" & R_XmlText_Class("ArticleContent/NoCheck", "�Բ�������δͨ������Ա��ˣ����ܲ鿴�շ�{$ChannelShortName}") & "</li>"
            ErrMsg_NeedPoint = Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ArticleContent/NeedPoint", "<p align='center'><br><br>�Բ����Ķ�������Ҫ���� <b><font color=red>{$NeedPoint}</font></b> {$PointUnit}{$PointName}������Ŀǰֻ�� <b><font color=blue>{$NowPoint}</font></b> {$PointUnit}{$PointName}���á�{$PointName}�����㣬�޷��Ķ����ġ�����������ϵ���г�ֵ��</p>"), "{$InfoPoint}", rsArticle("InfoPoint")), "{$NeedPoint}", rsArticle("InfoPoint")), "{$NowPoint}", UserPoint), "{$PointName}", PointName), "{$PointUnit}", PointUnit)
            ErrMsg_UsePoint = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ArticleContent/UsePoint", "<p align='center'><br><br>�Ķ�������Ҫ���� <b><font color=red>{$InfoPoint}</font></b> {$PointUnit}{$PointName}����Ŀǰ���� <b><font color=blue>{$NowPoint}</font></b> {$PointUnit}{$PointName}���á��Ķ����ĺ��㽫ʣ�� <b><font color=green>{$FinalPoint}</font></b> {$PointUnit}{$PointName}<br><br>��ȷʵԸ�⻨�� <b><font color=red>{$InfoPoint}</font></b> {$PointUnit}{$PointName}���Ķ�������<br><br><a href='{$FileName}?Pay=yes&ArticleID={$ArticleID}'>��Ը��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='{$InstallDir}index.asp'>�Ҳ�Ը��</a></p>"), "{$InfoPoint}", rsArticle("InfoPoint")), "{$NowPoint}", UserPoint), "{$FinalPoint}", UserPoint - rsArticle("InfoPoint")), "{$FileName}", strFileName), "{$ArticleID}", ArticleID), "{$InstallDir}", strInstallDir), "{$PointName}", PointName), "{$PointUnit}", PointUnit)
            ErrMsg_OutTime = R_XmlText_Class("ArticleContent/OutTime", "<p align='center'><br><br><font color=red>�Բ��𣬱���Ϊ�շ����ݣ���������Ч���Ѿ����ڣ������޷��Ķ����ġ�����������ϵ���г�ֵ��</font></p>")
            ErrMsg_Overflow_Total = "<li>" & R_XmlText_Class("ArticleContent/Overflow_Total", "���Ѿ��ﵽ�򳬹���Ч�������ܲ鿴����Ϣ������") & "</li>"
            ErrMsg_Overflow_Today = "<li>" & R_XmlText_Class("ArticleContent/Overflow_Today", "���Ѿ��ﵽ�򳬹��������ܲ鿴����Ϣ������") & "</li>"
            Select Case rsArticle("InfoPurview")
            Case 0
                If ClassPurview > 0 Then
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
                PurviewChecked = FoundInArr(rsArticle("arrGroupID"), GroupID, ",")
                If PurviewChecked = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & ErrMsg_PurviewCheckedErr2
                End If
            End Select
            If PurviewChecked = True Then
                If rsArticle("InfoPoint") > 0 And rsArticle("InfoPoint") < 9999 Then
                    If GroupType < 1 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_NoMail
                    ElseIf GroupType = 1 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_NoCheck
                    Else
                        Dim trs, ValidConsumeLogID, DividePoint

                        If UserChargeType = 0 Then   '��������
                            ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, ArticleID, rsArticle("ChargeType"), rsArticle("PitchTime"), rsArticle("ReadTimes"))
                            If ValidConsumeLogID = 0 Then   '���û���ҵ���¼���ѣ���Ҫ��ʼ�Ʒ�
                                If UserPoint < rsArticle("InfoPoint") Then  '����û��ĵ���С��Ҫ�۵ĵ���
                                    FoundErr = True
                                    ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                Else
                                    If LCase(Trim(Request("Pay"))) = "yes" Then  '����û�ȷ��Ҫ�۵�
                                        Conn.Execute "update PE_User set UserPoint=UserPoint-" & rsArticle("InfoPoint") & " where UserName='" & UserName & "'"
                                        Call AddConsumeLog("System", ModuleType, UserName, ArticleID, rsArticle("InfoPoint"), 2, "���ڲ鿴�շ�" & ChannelShortName & "��" & rsArticle("Title"))
                                        If rsArticle("DividePercent") <= 0 Then
                                            DividePoint = 0
                                        ElseIf rsArticle("DividePercent") > 0 And rsArticle("DividePercent") < 100 Then
                                            DividePoint = PE_CLng(rsArticle("InfoPoint") * rsArticle("DividePercent") / 100)
                                        Else
                                            DividePoint = rsArticle("InfoPoint")
                                        End If
                                        If DividePoint > 0 Then
                                            Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsArticle("Inputer") & "'"
                                            Call AddConsumeLog("System", ModuleType, rsArticle("Inputer"), 0, DividePoint, 1, "�ӡ�" & rsArticle("Title") & "�����շ��зֳ�")
                                        End If
                                    Else    '���û�û��ȷ��ǰ���Ƚ��п۷���ʾ
                                        FoundErr = True
                                        ErrMsg = ErrMsg & ErrMsg_UsePoint
                                    End If
                                End If
                            Else   '����ҵ������Ѽ�¼��ֱ�Ӹ������Ѽ�¼�����Ѵ���
                                Conn.Execute ("update PE_ConsumeLog set Times=Times+1,IP='" & UserTrueIP & "' where LogID=" & ValidConsumeLogID & "")
                            End If
                        Else
                            If ValidDays <= 0 Then  '����
                                If UserChargeType = 1 Or UserChargeType = 2 Then '��Ч�����ȣ�������ͬʱ�жϵ�ȯ����Ч�ڣ���ȯ�������Ч�ڵ��ں󣬾Ͳ��ɲ鿴�շ�����
                                    FoundErr = True
                                    ErrMsg = ErrMsg & ErrMsg_OutTime
                                Else
                                    '���ں��յ�ȯ��������
                                    ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, ArticleID, rsArticle("ChargeType"), rsArticle("PitchTime"), rsArticle("ReadTimes"))
                                    If ValidConsumeLogID = 0 Then   '���û���ҵ���¼���ѣ���Ҫ��ʼ�Ʒ�
                                        If UserPoint < rsArticle("InfoPoint") Then  '����û��ĵ���С��Ҫ�۵ĵ���
                                            FoundErr = True
                                            ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                        Else
                                            If LCase(Trim(Request("Pay"))) = "yes" Then  '����û�ȷ��Ҫ�۵�
                                                Conn.Execute "update PE_User set UserPoint=UserPoint-" & rsArticle("InfoPoint") & " where UserName='" & UserName & "'"
                                                Call AddConsumeLog("System", ModuleType, UserName, ArticleID, rsArticle("InfoPoint"), 2, "���ڲ鿴�շ�" & ChannelShortName & "��" & rsArticle("Title"))
                                                If rsArticle("DividePercent") <= 0 Then
                                                    DividePoint = 0
                                                ElseIf rsArticle("DividePercent") > 0 And rsArticle("DividePercent") < 100 Then
                                                    DividePoint = PE_CLng(rsArticle("InfoPoint") * rsArticle("DividePercent") / 100)
                                                Else
                                                    DividePoint = rsArticle("InfoPoint")
                                                End If
                                                If DividePoint > 0 Then
                                                    Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsArticle("Inputer") & "'"
                                                    Call AddConsumeLog("System", ModuleType, rsArticle("Inputer"), 0, DividePoint, 1, "�ӡ�" & rsArticle("Title") & "�����շ��зֳ�")
                                                End If
                                            Else    '���û�û��ȷ��ǰ���Ƚ��п۷���ʾ
                                                FoundErr = True
                                                ErrMsg = ErrMsg & ErrMsg_UsePoint
                                            End If
                                        End If
                                    Else   '����ҵ������Ѽ�¼��ֱ�Ӹ������Ѽ�¼�����Ѵ���
                                        Conn.Execute ("update PE_ConsumeLog set Times=Times+1,IP='" & UserTrueIP & "' where LogID=" & ValidConsumeLogID & "")
                                    End If
                                End If
                            Else   '��Ч����
                                '�������Ч���ڵĿ۷ѷ�ʽ����
                                If PE_CLng(UserSetting(15)) > 0 Then   'PE_CLng(UserSetting(15))����Ч���ڣ��鿴�շ������Ƿ�۵�ͼ�¼��0Ϊ���۵㣬1Ϊ���۵㣬������¼��2Ϊ�۵�
                                    '�������Ѽ�¼
                                    ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, ArticleID, rsArticle("ChargeType"), rsArticle("PitchTime"), rsArticle("ReadTimes"))
                                    If ValidConsumeLogID = 0 Then    'δ�ҵ����Ѽ�¼
                                        If PE_CLng(UserSetting(16)) > 0 Then   '��Ч�����ܹ����Բ鿴��������Ϣ
                                            Set trs = Conn.Execute("select count(0) from PE_ConsumeLog where UserName='" & UserName & "' and Income_Payout=2 and InfoID>0")
                                            If PE_CLng(trs(0)) >= PE_CLng(UserSetting(16)) Then
                                                FoundErr = True
                                                ErrMsg = ErrMsg & ErrMsg_Overflow_Total
                                            End If
                                            Set trs = Nothing
                                        End If
                                        If PE_CLng(UserSetting(17)) > 0 Then    '��Ч����ÿ����Բ鿴��������Ϣ
                                            Set trs = Conn.Execute("select count(0) from PE_ConsumeLog where UserName='" & UserName & "' and Income_Payout=2 and InfoID>0 and DateDiff(" & PE_DatePart_D & ",LogTime," & PE_Now & ")<1")
                                            If PE_CLng(trs(0)) >= PE_CLng(UserSetting(17)) Then
                                                FoundErr = True
                                                ErrMsg = ErrMsg & ErrMsg_Overflow_Today
                                            End If
                                            Set trs = Nothing
                                        End If
                                        If FoundErr = False Then
                                            If PE_CLng(UserSetting(15)) = 1 Then  '���۵㣬������¼
                                                Call AddConsumeLog("System", ModuleType, UserName, ArticleID, 0, 2, "��Ч���ڲ鿴�շ�" & ChannelShortName & "��" & rsArticle("Title") & "��Ӧ�۵�����" & rsArticle("InfoPoint") & "")
                                            Else  '�۵�
                                                If UserPoint >= rsArticle("InfoPoint") Then   '��������㹻
                                                    '�����Ŀ۷���ʾ
                                                    If LCase(Trim(Request("Pay"))) = "yes" Then  '����û�ȷ��Ҫ�۵�
                                                        Conn.Execute "update PE_User set UserPoint=UserPoint-" & rsArticle("InfoPoint") & " where UserName='" & UserName & "'"
                                                        Call AddConsumeLog("System", ModuleType, UserName, ArticleID, rsArticle("InfoPoint"), 2, "��Ч���ڲ鿴�շ�" & ChannelShortName & "��" & rsArticle("Title"))
                                                        If rsArticle("DividePercent") <= 0 Then
                                                            DividePoint = 0
                                                        ElseIf rsArticle("DividePercent") > 0 And rsArticle("DividePercent") < 100 Then
                                                            DividePoint = PE_CLng(rsArticle("InfoPoint") * rsArticle("DividePercent") / 100)
                                                        Else
                                                            DividePoint = rsArticle("InfoPoint")
                                                        End If
                                                        If DividePoint > 0 Then
                                                            Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsArticle("Inputer") & "'"
                                                            Call AddConsumeLog("System", ModuleType, rsArticle("Inputer"), 0, DividePoint, 1, "�ӡ�" & rsArticle("Title") & "�����շ��зֳ�")
                                                        End If
                                                    Else    '���û�û��ȷ��ǰ���Ƚ��п۷���ʾ
                                                        FoundErr = True
                                                        ErrMsg = ErrMsg & ErrMsg_UsePoint
                                                    End If
                                                Else   '����������ʱ
                                                    If UserChargeType = 2 Then '�����������Ч�ڵ��ں󣬾Ͳ��ɲ鿴�շ����ݡ�
                                                        FoundErr = True
                                                        ErrMsg = ErrMsg_NeedPoint
                                                    Else   '��Ч�����Ȼ���Ч�ڹ��ں͵������꣬�Ų��ɲ鿴�շ����ݣ���ʱֻ���¼
                                                        Call AddConsumeLog("System", ModuleType, UserName, ArticleID, 0, 2, "��Ч���ڲ鿴�շ�" & ChannelShortName & "��" & rsArticle("Title") & "��Ӧ�۵�����" & rsArticle("InfoPoint") & "")
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else     '�ҵ����Ѽ�¼��ֻ���������Ѵ���
                                        Conn.Execute ("update PE_ConsumeLog set Times=Times+1,IP='" & UserTrueIP & "' where LogID=" & ValidConsumeLogID & "")
                                    End If
                                Else   '��Ч���ڣ��鿴�շ����ݲ��۵�����Ҳ������¼��
                                    '�����κδ���
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If FoundErr = True Then
        If PreviewContentLength > 0 Then
            If Trim(rsArticle("Intro")) & "" <> "" Then
                ErrMsg = "<p align=left><b>" & XmlText_Class("ArticleContent/Intro", "���ݼ�飺") & "</b><br><br>" & ReplaceText(rsArticle("Intro"), 1) & "����</p>" & ErrMsg
            Else
                ErrMsg = "<p align=left><b>" & XmlText_Class("ArticleContent/Content", "����Ԥ����") & "</b><br><br>" & Left(ReplaceText(nohtml(rsArticle("Content")), 1), PreviewContentLength) & "����</p>" & ErrMsg
            End If
        End If
        GetArticleContent = ErrMsg
        Exit Function
    End If
    GetArticleContent = GetArticleContent_CurrentPage
End Function

Public Function GetArticleContent_CurrentPage()
    '�ж��Ƿ�Ϊת�����ӣ�����ǣ����� GetArticleConten�ķ�������Ϊת�����ӵ�ַ
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then
        GetArticleContent = "<script language='javascript'>window.location.href='" & rsArticle("LinkUrl") & "';</script>"
    Else
        Dim PaginationType
        PaginationType = rsArticle("PaginationType")
        strTempContent = ReplaceKeyLink(ReplaceText(Replace(Replace(rsArticle("Content") & "", "[InstallDir_ChannelDir]", ChannelUrl & "/"), "{$UploadDir}", UploadDir), 1))
        Select Case PaginationType
            Case 0    '����ҳ��ʾ
                GetArticleContent_CurrentPage = strTempContent
            Case 1    '�Զ���ҳ��ʾ
                GetArticleContent_CurrentPage = AutoPagination()
            Case 2    '�ֶ���ҳ��ʾ
                GetArticleContent_CurrentPage = ManualPagination()
        End Select
   End If
End Function



'=================================================
'��������ManualPagination
'��  �ã������ֶ���ҳ��ʽ��ʾ���¾��������
'��  ������
'=================================================
Private Function ManualPagination()
    Dim strTemp
    Dim i
    Dim HasPageTitle
    i = 1
    strTemp = ""
    HasPageTitle = False
    regEx.MultiLine = False
    regEx.Pattern = "\[NextPage(.*?)\]"
    Set Matches = regEx.Execute(strTempContent)
    For Each Match In Matches
        If i = 1 Then
            If Match.SubMatches(0) <> "" Then
                If nohtml(Left(strTempContent, InStr(strTempContent, Match.SubMatches(0)))) <> "" Then
                    strTempContent = PE_Replace(strTempContent, Match.Value, "") 'ȥ����һ��ƥ���[NextPage ��ҳ����]
                    strTemp = "�� 1 ҳ��" & Trim(Match.SubMatches(0))
                Else
                    strTemp = "�� 1 ҳ��$$$�� 2 ҳ��" & Trim(Match.SubMatches(0))
                    i = i + 1
                End If
                HasPageTitle = True
            Else
                strTemp = "�� 1 ҳ��$$$�� 2 ҳ��"
                i = i + 1
            End If
        Else
            If Match.SubMatches(0) <> "" Then
                strTemp = strTemp & "$$$�� " & i & " ҳ��" & Trim(Match.SubMatches(0))
                HasPageTitle = True
            Else
                strTemp = strTemp & "$$$�� " & i & " ҳ��"
            End If
        End If
        i = i + 1
    Next
    regEx.MultiLine = True
    strTempContent = regEx.Replace(strTempContent, "[NextPage]")  '��[NextPage ��ҳ����]�滻��[NextPage]
    If HasPageTitle = True Then
        strContentPageTitleArr = strTemp
    Else
        strContentPageTitleArr = ""
    End If
    
    Dim arrContent
    If InStr(strTempContent, "[NextPage]") <= 0 Then
        ManualPagination = strTempContent & "</p>"
        Exit Function
    End If
    arrContent = Split(strTempContent, "[NextPage]")
    totalPage = UBound(arrContent) + 1
    If CurrentPage < 1 Then CurrentPage = 1
    If CurrentPage > totalPage Then CurrentPage = totalPage
    If InStr(strHtml, "{$ShowPageContent}") > 0 Then 
        ManualPagination = arrContent(CurrentPage - 1) & "</p>" & GetContentPagesDiv(totalPage, CurrentPage)
    Else
        ManualPagination = arrContent(CurrentPage - 1) & "</p>" & GetContentPages(totalPage, CurrentPage)
    End If	
End Function
'****************************************************************
'�� �� ����GetSpecialPosition(ByVal strContent,ByVal BeginPoint,ByVal MaxCharperPage)
'��    �ã�����ѷ�ҳ��ŵ���<img ,<p , <li��֮��
'��    ����strContent �������� BeginPoint ��ҳ��ʼ�� MaxCharperPage ÿҳ����
'****************************************************************
Private Function GetSpecialPosition(ByVal strContent, ByVal BeginPoint, ByVal MaxCharPerPage)
    Dim strTemp
    regEx.Pattern = "^(([^<]*>)[^<]{0,100})(?:<p|<img|<br|<li)*"
    strTemp = Right(strContent, Len(strContent) - BeginPoint)
    Set Matches = regEx.Execute(strTemp)
    If Matches.Count > 0 Then
        GetSpecialPosition = BeginPoint + Len(Matches(0).SubMatches(1)) + 1
    Else
        GetSpecialPosition = BeginPoint
    End If
End Function
'=================================================
'�� �� ����AutoPagination
'��    �ã������Զ���ҳ��ʽ��ʾ���¾��������
'��    ������
'=================================================
Private Function AutoPagination()
    Dim lContent, MaxCharPerPage
    Dim ContentLen, MaxPerPage, lngBound
    Dim BeginPoint, EndPoint
    ContentLen = Len(strTempContent)
    MaxCharPerPage = rsArticle("MaxCharPerPage")
    If MaxCharPerPage <= 100 Or ContentLen <= MaxCharPerPage Or InStr(LCase(strTempContent), "<table") > 0 Or InStr(LCase(strTempContent), "</table>") > 0 Then
        AutoPagination = strTempContent & "</p>"
       Exit Function
    End If

    lContent = LCase(strTempContent)
    If InStr(lContent, "<table") > 0 And InStr(lContent, "</table>") > 0 Then
        AutoPagination = strTempContent & "</p>"
        Exit Function
    End If
    
    totalPage = ContentLen \ MaxCharPerPage
    If MaxCharPerPage * totalPage < ContentLen Then
        totalPage = totalPage + 1
    End If
    If CurrentPage < 1 Then CurrentPage = 1
    If CurrentPage > totalPage Then CurrentPage = totalPage
  
    If CurrentPage = 1 Then
        BeginPoint = 1
    Else
        BeginPoint = MaxCharPerPage * (CurrentPage - 1)
        BeginPoint = GetSpecialPosition(strTempContent, BeginPoint, MaxCharPerPage) '��ֹ�Ѳ��Html���룬�����µ�λ��
    End If
    If CurrentPage = totalPage Then
        EndPoint = ContentLen
    Else
        EndPoint = MaxCharPerPage * CurrentPage
        If EndPoint >= ContentLen Then
            EndPoint = ContentLen
        Else
            EndPoint = GetSpecialPosition(strTempContent, EndPoint, MaxCharPerPage)
        End If
    End If
    If EndPoint < BeginPoint Then EndPoint = BeginPoint
    If InStr(strHtml, "{$ShowPageContent}") > 0 Then 
	AutoPagination = Mid(strTempContent, BeginPoint, EndPoint - BeginPoint) & "</p>" & GetContentPagesDiv(totalPage, CurrentPage)
    Else
        AutoPagination = Mid(strTempContent, BeginPoint, EndPoint - BeginPoint) & "</p>" & GetContentPages(totalPage, CurrentPage)
    End If
	
End Function

Private Function GetContentPagesDiv(totalPage, CurrentPage)
    Dim strPages
    Dim tmpArticleUrl
    Dim i
    Dim ShowPageNum 'ÿҳ��ʾ��ҳ��
    Dim startFlag, endFlag '��ҳ�Ŀ�ʼ�������
    ShowPageNum = 10
    If totalPage > ShowPageNum Then
        If (CurrentPage Mod ShowPageNum) <> 0 Then
            startFlag = (CurrentPage \ ShowPageNum) * ShowPageNum + 1
            endFlag = startFlag + ShowPageNum - 1
        Else
            startFlag = CurrentPage - ShowPageNum + 1
            endFlag = CurrentPage
        End If
    Else
        startFlag = 1
        endFlag = totalPage
    End If
    If endFlag > totalPage Then
       endFlag = totalPage
    End If
    tmpArticleUrl = Left(ArticleUrl, InStrRev(ArticleUrl, ".", -1, vbBinaryCompare) - 1)
    strPages = strPages & "<div align='left'><b>"
    If CurrentPage > 1 Then
        If UseCreateHTML > 0 And ClassPurview = 0 And rsArticle("InfoPoint") = 0 And rsArticle("InfoPurview") = 0 And strFileName <> "Print.asp" Then
            If CurrentPage > 2 Then
                If CurrentPage > ShowPageNum And CurrentPage <= totalPage And startFlag > 1 Then
                    strPages = strPages & "<div id=pages style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><a href='" & tmpArticleUrl & "_" & startFlag - 1 & FileExt_Item & "'><<</a></div>"
                End If
                strPages = strPages & "<div id=nextpage style='float:left;border:1px solid #39f;text-align:center;WIDTH: 55px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;'><a href='" & tmpArticleUrl & "_" & CurrentPage - 1 & FileExt_Item & "'>" & XmlText_Class("ContentPages/PrevPage", "��һҳ") & "</a></div>"
            Else
                strPages = strPages & "<div id=nextpage style='float:left;border:1px solid #39f;text-align:center;WIDTH: 55px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;'><a href='" & tmpArticleUrl & FileExt_Item & "'>" & XmlText_Class("ContentPages/PrevPage", "��һҳ") & "</a></div> "
            End If
        Else
            If CurrentPage > ShowPageNum And CurrentPage <= totalPage Then
                strPages = strPages & "<div id=pages style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & startFlag - 1 & "' ><<</a></div>"
            End If
            strPages = strPages & "<div id=nextpage style='float:left;border:1px solid #39f;text-align:center;WIDTH: 55px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;'><a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & CurrentPage - 1 & "'>" & XmlText_Class("ContentPages/PrevPage", "��һҳ") & "</a></div>"
        End If
    End If

    For i = startFlag To endFlag
        If i = CurrentPage Then
            strPages = strPages & "<div id=pages style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><font color='red'>" & CStr(i) & "</font></div>"
        Else
            If UseCreateHTML > 0 And ClassPurview = 0 And rsArticle("InfoPoint") = 0 And rsArticle("InfoPurview") = 0 And strFileName <> "Print.asp" Then
                If i > 1 Then
                    strPages = strPages & "<div id=pages style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;' ><a href='" & tmpArticleUrl & "_" & i & FileExt_Item & "'>" & i & "</a></div>"
                Else
                    strPages = strPages & "<div id=pages style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><a href='" & tmpArticleUrl & FileExt_Item & "'>" & i & "</a></div>"
                End If
            Else
                strPages = strPages & "<div id=pages style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & i & "'>" & i & "</a></div>"
            End If
        End If
    Next
    If endFlag < totalPage Then
        strPages = strPages & " <div id=pages  style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'>...</div>"
    End If
    If CurrentPage < totalPage Then
        If UseCreateHTML > 0 And ClassPurview = 0 And rsArticle("InfoPoint") = 0 And rsArticle("InfoPurview") = 0 And strFileName <> "Print.asp" Then
            strPages = strPages & "<div id=nextpage style='float:left;border:1px solid #39f;text-align:center;WIDTH: 55px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;'><a href='" & tmpArticleUrl & "_" & CurrentPage + 1 & FileExt_Item & "'>" & XmlText_Class("ContentPages/NextPage", "��һҳ") & "</a></div> "
            If totalPage > ShowPageNum And endFlag < totalPage Then
                strPages = strPages & "<div id=pages  style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><a href='" & tmpArticleUrl & "_" & endFlag + 1 & FileExt_Item & "'>>></a></div>"
            End If
        Else
            strPages = strPages & "<div id=nextpage style='float:left;border:1px solid #39f;text-align:center;WIDTH: 55px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;'><a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & CurrentPage + 1 & "'>" & XmlText_Class("ContentPages/NextPage", "��һҳ") & "</a></div>"
            If totalPage > endFlag Then
                strPages = strPages & "<div id=pages  style='float:left;border:1px solid #39f;text-align:center;WIDTH: 30px;HEIGHT: 20px;background-color:#eee;padding:2px;margin:5px;margin-right:0;'><a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & endFlag + 1 & "'>>></a></div>"
            End If
        End If
    End If
    strPages = strPages & "</b></div>"
    GetContentPagesDiv = strPages
End Function

Private Function GetContentPages(totalPage, CurrentPage)
    Dim strPages
    Dim tmpArticleUrl
    Dim i
    Dim ShowPageNum 'ÿҳ��ʾ��ҳ��
    Dim startFlag, endFlag '��ҳ�Ŀ�ʼ�������
    ShowPageNum = 10
    If totalPage > ShowPageNum Then
        If (CurrentPage Mod ShowPageNum) <> 0 Then
            startFlag = (CurrentPage \ ShowPageNum) * ShowPageNum + 1
            endFlag = startFlag + ShowPageNum - 1
        Else
            startFlag = CurrentPage - ShowPageNum + 1
            endFlag = CurrentPage
        End If
    Else
        startFlag = 1
        endFlag = totalPage
    End If
    If endFlag > totalPage Then
       endFlag = totalPage
    End If
    tmpArticleUrl = Left(ArticleUrl, InStrRev(ArticleUrl, ".", -1, vbBinaryCompare) - 1)
    strPages = strPages & "<p align='center'><b>"
    If CurrentPage > 1 Then
        If UseCreateHTML > 0 And ClassPurview = 0 And rsArticle("InfoPoint") = 0 And rsArticle("InfoPurview") = 0 And strFileName <> "Print.asp" Then
            If CurrentPage > 2 Then
                If CurrentPage > ShowPageNum And CurrentPage <= totalPage And startFlag > 1 Then
                    strPages = strPages & "<a href='" & tmpArticleUrl & "_" & startFlag - 1 & FileExt_Item & "'>&nbsp;&lt;&lt;&nbsp;</a>"
                End If
                strPages = strPages & "<a href='" & tmpArticleUrl & "_" & CurrentPage - 1 & FileExt_Item & "'>" & XmlText_Class("ContentPages/PrevPage", "��һҳ") & "</a>&nbsp;&nbsp;"
            Else
                strPages = strPages & "<a href='" & tmpArticleUrl & FileExt_Item & "'>" & XmlText_Class("ContentPages/PrevPage", "��һҳ") & "</a>&nbsp;&nbsp;"
            End If
        Else
            If CurrentPage > ShowPageNum And CurrentPage <= totalPage Then
                strPages = strPages & "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & startFlag - 1 & "' >&nbsp;&lt;&lt;&nbsp;</a>"
            End If
            strPages = strPages & "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & CurrentPage - 1 & "'>" & XmlText_Class("ContentPages/PrevPage", "��һҳ") & "</a>&nbsp;&nbsp;"
        End If
    End If

    For i = startFlag To endFlag
        If i = CurrentPage Then
            strPages = strPages & "<font color='red'>[" & CStr(i) & "]</font>&nbsp;"
        Else
            If UseCreateHTML > 0 And ClassPurview = 0 And rsArticle("InfoPoint") = 0 And rsArticle("InfoPurview") = 0 And strFileName <> "Print.asp" Then
                If i > 1 Then
                    strPages = strPages & "<a href='" & tmpArticleUrl & "_" & i & FileExt_Item & "'>[" & i & "]</a>&nbsp;"
                Else
                    strPages = strPages & "<a href='" & tmpArticleUrl & FileExt_Item & "'>[" & i & "]</a>&nbsp;"
                End If
            Else
                strPages = strPages & "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & i & "'>[" & i & "]</a>&nbsp;"
            End If
        End If
    Next
    If endFlag < totalPage Then
        strPages = strPages & " ... "
    End If
    If CurrentPage < totalPage Then
        If UseCreateHTML > 0 And ClassPurview = 0 And rsArticle("InfoPoint") = 0 And rsArticle("InfoPurview") = 0 And strFileName <> "Print.asp" Then
            strPages = strPages & "<a href='" & tmpArticleUrl & "_" & CurrentPage + 1 & FileExt_Item & "'>" & XmlText_Class("ContentPages/NextPage", "��һҳ") & "</a> "
            If totalPage > ShowPageNum And endFlag < totalPage Then
                strPages = strPages & "<a href='" & tmpArticleUrl & "_" & endFlag + 1 & FileExt_Item & "'>&nbsp;&gt;&gt;&nbsp;</a>"
            End If
        Else
            strPages = strPages & "&nbsp;<a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & CurrentPage + 1 & "'>" & XmlText_Class("ContentPages/NextPage", "��һҳ") & "</a>"
            If totalPage > endFlag Then
                strPages = strPages & "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & endFlag + 1 & "'>&nbsp;&gt;&gt;&nbsp;</a>"
            End If
        End If
    End If
    strPages = strPages & "</b></p>"
    GetContentPages = strPages
End Function

Private Function GetArticleTitle()
    Dim strTitle
    If Trim(rsArticle("TitleIntact")) <> "" Then
        strTitle = Trim(rsArticle("TitleIntact"))
    Else
        strTitle = Trim(rsArticle("Title"))
    End If
    GetArticleTitle = strTitle
End Function

Private Function GetSubheading()
    Dim strSubheading
    If Trim(rsArticle("Subheading")) <> "" Then
        strSubheading = Trim(rsArticle("Subheading"))
    End If
    GetSubheading = strSubheading
End Function

Private Function GetArticleInfo()
    Dim ArticleHits
    If UseCreateHTML > 0 And Not (ClassPurview > 0 Or rsArticle("InfoPoint") > 0) Then
        ArticleHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?ArticleID=" & ArticleID & "'></script>"
    Else
        ArticleHits = rsArticle("Hits")
    End If
    GetArticleInfo = Replace(Replace(Replace(Replace(R_XmlText_Class("ArticleInfo", "���ߣ�{$Author}&nbsp;&nbsp;&nbsp;&nbsp;{$ChannelShortName}��Դ��{$CopyFrom}&nbsp;&nbsp;&nbsp;&nbsp;�������{$Hits}&nbsp;&nbsp;&nbsp;&nbsp;����ʱ�䣺{$Time}"), "{$Author}", GetAuthorInfo(rsArticle("Author"), ChannelID)), "{$CopyFrom}", GetCopyFromInfo(rsArticle("CopyFrom"), ChannelID)), "{$Time}", FormatDateTime(rsArticle("UpdateTime"), 2)), "{$Hits}", ArticleHits)
End Function

Private Function GetArticleHits()
    If UseCreateHTML > 0 And Not (ClassPurview > 0 Or rsArticle("InfoPoint") > 0) Then
        GetArticleHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?ArticleID=" & ArticleID & "'></script>"
    Else
        GetArticleHits = rsArticle("Hits")
    End If
End Function

Private Function GetArticleEditor()
    GetArticleEditor = Replace(Replace(R_XmlText_Class("ArticleEditor", "{$ChannelShortName}¼�룺{$Inputer}&nbsp;&nbsp;&nbsp;&nbsp;���α༭��{$Editor}&nbsp;"), "{$Inputer}", replacebadchar(rsArticle("Inputer"))), "{$Editor}", rsArticle("Editor"))
End Function

Private Function GetArticleAction()
    Dim strAction
    strAction = Replace(Replace(Replace(Replace(XmlText_Class("ArticleAction", "��<a href='{$ChannelUrl}/Comment.asp?ArticleID={$ArticleID}' target='_blank'>��������</a>����<a href='{$InstallDir}User/User_Favorite.asp?Action=Add&ChannelID={$ChannelID}&InfoID={$ArticleID}' target='_blank'>�����ղ�</a>����<a href='{$ChannelUrl}/SendMail.asp?ArticleID={$ArticleID}' target='_blank'>���ߺ���</a>����<a href='{$ChannelUrl}/Print.asp?ArticleID={$ArticleID}' target='_blank'>��ӡ����</a>����<a href='javascript:window.close();'>�رմ���</a>��"), "{$ChannelUrl}", ChannelUrl_ASPFile), "{$ArticleID}", rsArticle("ArticleID")), "{$InstallDir}", strInstallDir), "{$ChannelID}", ChannelID)
    GetArticleAction = strAction
End Function

Private Function GetArticleProtect()
    If EnableProtect = True Then
        GetArticleProtect = " oncontextmenu='return false' ondragstart='return false' onselectstart ='return false' onselect='document.selection.empty()' oncopy='document.selection.empty()' onbeforecopy='return false' onmouseup='document.selection.empty()'"
    Else
        GetArticleProtect = ""
    End If
End Function


Private Function GetArticleProperty()
    Dim strProperty
    If rsArticle("OnTop") = True Then
        strProperty = strProperty & XmlText_Class("ShowArticle/OnTop", "<font color=blue>��</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Hits") >= HitsOfHot Then
        strProperty = strProperty & XmlText_Class("ShowArticle/Hot", "<font color=red>��</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Elite") = True Then
        strProperty = strProperty & XmlText_Class("ShowArticle/Elite", "<font color=green>��</font>")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;"
    End If
    strProperty = strProperty & "&nbsp;&nbsp;" & GetStars(rsArticle("Stars"))
    GetArticleProperty = strProperty
End Function

Private Function GetStars(Stars)
    GetStars = "<font color='" & XmlText_Class("ShowArticle/Star_Color", "#009900") & "'>" & String(Stars, XmlText_Class("ShowArticle/Star", "��")) & "</font>"
End Function

Public Function GetCustomFromTemplate(strValue)   '�õ��Զ����б�İ�����Ƶ�HTML����
    Dim strCustom, strParameter
    strCustom = strValue
    regEx.Pattern = "��ArticleList\((.*?)\)��([\s\S]*?)��\/ArticleList��"
    Set Matches = regEx.Execute(strCustom)
    For Each Match In Matches
        strParameter = Replace(Match.SubMatches(0), Chr(34), " ")
        strCustom = PE_Replace(strCustom, Match.Value, GetCustomFromLabel(strParameter, Match.SubMatches(1)))
    Next
    GetCustomFromTemplate = strCustom
End Function

Public Function GetListFromTemplate(ByVal strValue)
    Dim strList
    strList = strValue
    regEx.Pattern = "\{\$GetArticleList\((.*?)\)\}"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        strList = PE_Replace(strList, Match.Value, GetListFromLabel(Match.SubMatches(0)))
    Next
    GetListFromTemplate = strList
End Function

Public Function GetPicFromTemplate(ByVal strValue)
    Dim strPicList
    strPicList = strValue
    regEx.Pattern = "\{\$GetPicArticle\((.*?)\)\}"
    Set Matches = regEx.Execute(strPicList)
    For Each Match In Matches
        strPicList = PE_Replace(strPicList, Match.Value, GetPicFromLabel(Match.SubMatches(0)))
    Next
    GetPicFromTemplate = strPicList
End Function

Public Function GetSlidePicFromTemplate(ByVal strValue)
    Dim strSlidePic, InitSlideJS
    InitSlideJS = False
    strSlidePic = strValue
    regEx.Pattern = "\{\$GetSlidePicArticle\((.*?)\)\}"
    Set Matches = regEx.Execute(strSlidePic)
    For Each Match In Matches
        If InitSlideJS = False Then
            strSlidePic = PE_Replace(strSlidePic, Match.Value, JS_SlidePic & GetSlidePicFromLabel(Match.SubMatches(0)))
            InitSlideJS = True
        Else
            strSlidePic = PE_Replace(strSlidePic, Match.Value, GetSlidePicFromLabel(Match.SubMatches(0)))
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
        tChannelID = arrTemp(0)
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
    tChannelID = Replace(Trim(tChannelID), "|", ",")
	
    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(3))
    End Select
    
    Select Case UBound(arrTemp)
    Case 12
        GetSlidePicFromLabel = GetSlidePicArticle(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), -1)
    Case 13
        GetSlidePicFromLabel = GetSlidePicArticle(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)))
    Case Else
        GetSlidePicFromLabel = "����ʽ��ǩ��{$GetSlidePicArticle(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
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
        GetPicFromLabel = "����ʽ��ǩ��{$GetPicArticle(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If
    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        tChannelID = ChannelID
    Case Else
        tChannelID = arrTemp(0)
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
    tChannelID = Replace(Trim(tChannelID), "|", ",")
    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(3))
    End Select

    GetPicFromLabel = GetPicArticle(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CBool(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CLng(arrTemp(16)))
End Function

Private Function GetListFromLabel(ByVal strSource)
    Dim arrTemp
    Dim tChannelID, ArticleNum, arrClassID, tSpecialID, AuthorName, OrderType, OpenType
    If strSource = "" Then
        GetListFromLabel = ""
        Exit Function
    End If
    
    strSource = Replace(strSource, Chr(34), "")
    strSource = FillInArrStr(strSource, "1,listA,listbg,listbg2", 30)
    arrTemp = Split(strSource, ",")
    If UBound(arrTemp) + 1 < 30 Then
        GetListFromLabel = "����ʽ��ǩ��{$GetArticleList(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If

    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        tChannelID = ChannelID
    Case Else
        tChannelID = arrTemp(0)
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
    tChannelID = Replace(Trim(tChannelID), "|", ",")
    Select Case Trim(arrTemp(3))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(3))
    End Select
    
    Select Case Trim(arrTemp(5))
    Case "rsClass_TopNumber"
        ArticleNum = 8
    Case "TopNumber"
        ArticleNum = 8
    Case Else
        ArticleNum = PE_CLng(arrTemp(5))
    End Select
    
    AuthorName = Replace(Replace(Trim(arrTemp(8)), "?", ""), "&quot;", "")

    Select Case Trim(arrTemp(10))
    Case "rsClass_ItemListOrderType"
        OrderType = rsClass("ItemListOrderType")
    Case "ItemListOrderType"
        OrderType = ItemListOrderType
    Case Else
        OrderType = PE_CLng(arrTemp(10))
    End Select

    Select Case Trim(arrTemp(25))
    Case "rsClass_ItemOpenType"
        OpenType = rsClass("ItemOpenType")
    Case "ItemOpenType"
        OpenType = ItemOpenType
    Case Else
        OpenType = PE_CLng(arrTemp(25))
    End Select
    
    GetListFromLabel = GetArticleList(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), ArticleNum, PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), AuthorName, PE_CLng(arrTemp(9)), OrderType, PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CBool(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CBool(arrTemp(17)), PE_CLng(arrTemp(18)), PE_CBool(arrTemp(19)), PE_CBool(arrTemp(20)), PE_CBool(arrTemp(21)), PE_CBool(arrTemp(22)), PE_CBool(arrTemp(23)), PE_CBool(arrTemp(24)), OpenType, PE_CLng(arrTemp(26)), Trim(arrTemp(27)), Trim(arrTemp(28)), Trim(arrTemp(29)))

End Function

Private Function GetCustomFromLabel(strTemp, strList)
    Dim arrTemp
    Dim strArticlePic, strPicTemp, arrPicTemp
    Dim iChannelID, arrClassID, IncludeChild, iSpecialID, ItemNum, IsHot, IsElite, Author, DateNum, OrderType, UsePage, TitleLen, ContentLen
    Dim iCols, iColsHtml, iRows, iRowsHtml, iNumber
    Dim IncludePic
    If strTemp = "" Or strList = "" Then GetCustomFromLabel = "": Exit Function

    iCols = 1: iRows = 1: iColsHtml = "": iRowsHtml = ""
    regEx.Pattern = "��(Cols|Rows)=(\d{1,2})\s*(?:\||��)(.+?)��"
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
    If UBound(arrTemp) <> 13 and UBound(arrTemp) <> 12 Then
        GetCustomFromLabel = "�Զ����б��ǩ����ArticleList(�����б�)���б����ݡ�/ArticleList���Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If


    Select Case Trim(arrTemp(0))
    Case "ChannelID"
        iChannelID = ChannelID
    Case Else
        iChannelID = arrTemp(0)
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
    iChannelID = Replace(Trim(iChannelID), "|", ",")	
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
    If (PE_Clng(iChannelID) <> 0 and Instr(iChannelID,",")=0) and (PE_Clng(iChannelID)<>PrevChannelID Or ChannelID = 0) Then
        Call GetChannel(PE_Clng(iChannelID))
        PrevChannelID = iChannelID		 
    End If
    If FoundErr = True Then
        GetCustomFromLabel = ErrMsg
        Exit Function
    End If

    Dim rsField, ArrField, iField
    Set rsField = Conn.Execute("select FieldName,LabelName,FieldType from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
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
        sqlCustom = sqlCustom & "A.Content,"
    End If
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlCustom = sqlCustom & "A." & ArrField(0, iField) & ","
        Next
    End If
    sqlCustom = sqlCustom & "A.ArticleID,A.ChannelID,A.ClassID,A.Title,A.Subheading,A.Keyword,A.Intro,A.DefaultPicUrl"
    sqlCustom = sqlCustom & ",A.Author,A.CopyFrom,A.Inputer,A.Editor,A.UpdateTime,A.Stars,A.Hits,A.OnTop,A.Elite,A.InfoPoint,A.InfoPurview"
    sqlCustom = sqlCustom & ",C.ClassName,C.ParentDir,C.ClassDir,C.Readme,C.ClassPurview"
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
    PrevChannelID = 0
    Do While Not rsCustom.EOF
        'If iChannelID = 0 Then
            If rsCustom("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsCustom("ChannelID"))
                PrevChannelID = rsCustom("ChannelID")
            End If
       ' End If
                
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

        strTemp = PE_Replace(strTemp, "{$ArticleID}", rsCustom("ArticleID"))
        If InStr(strTemp, "{$ArticleUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ArticleUrl}", GetArticleUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("ArticleID"), rsCustom("ClassPurview"), rsCustom("InfoPurview"), rsCustom("InfoPoint")))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        strTemp = PE_Replace(strTemp, "{$Author}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$CopyFrom}", rsCustom("CopyFrom"))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$Inputer}", rsCustom("Inputer"))
        strTemp = PE_Replace(strTemp, "{$Editor}", rsCustom("Editor"))
        If InStr(strTemp, "{$InfoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$InfoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$ReadPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$ReadPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$Keyword}") > 0 Then strTemp = PE_Replace(strTemp, "{$Keyword}", GetKeywords(",", rsCustom("Keyword")))
        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "OnTop")
        ElseIf rsCustom("Elite") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Elite")
        ElseIf rsCustom("Hits") > HitsOfHot Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Hot")
        Else
            strTemp = PE_Replace(strTemp, "{$Property}", "Common")
        End If
        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Top}", strTop2)
        Else
            strTemp = PE_Replace(strTemp, "{$Top}", "")
        End If
        If rsCustom("Elite") = True Then
            strTemp = PE_Replace(strTemp, "{$Elite}", strElite2)
        Else
            strTemp = PE_Replace(strTemp, "{$Elite}", "")
        End If
        If rsCustom("Hits") > HitsOfHot Then
            strTemp = PE_Replace(strTemp, "{$Hot}", strHot2)
        Else
            strTemp = PE_Replace(strTemp, "{$Hot}", "")
        End If
        
        If TitleLen > 0 Then
            strTemp = PE_Replace(strTemp, "{$Title}", GetSubStr(rsCustom("Title"), TitleLen, ShowSuspensionPoints))
        Else
            strTemp = PE_Replace(strTemp, "{$Title}", rsCustom("Title"))
        End If
        strTemp = PE_Replace(strTemp, "{$TitleOriginal}", rsCustom("Title"))

        If ContentLen > 0 Then
            If InStr(strTemp, "{$Content}") > 0 Then strTemp = PE_Replace(strTemp, "{$Content}", Left(nohtml(rsCustom("Content")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$Content}", "")
        End If
        strTemp = PE_Replace(strTemp, "{$Subheading}", rsCustom("Subheading"))
        strTemp = PE_Replace(strTemp, "{$Intro}", rsCustom("Intro"))
        
        '�滻��ҳͼƬ
        regEx.Pattern = "\{\$ArticlePic\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strArticlePic = GetDefaultPicUrl(Trim(rsCustom("DefaultPicUrl")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)))
            strTemp = Replace(strTemp, Match.Value, strArticlePic)
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
                        strTemp = PE_Replace(strTemp, ArrField(1, iField), "<img class='fieldImg' src='" &PE_HTMLDecode(rsCustom(Trim(ArrField(0, iField))))&"' border=0>")	
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

Public Function GetLinkUrlContent(iLinkUrl, iArticleID)
    Dim strLinkUrlContent
    strLinkUrlContent = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?ShowType=0&ArticleID=" & iArticleID & "'></script>" & vbCrLf
    strLinkUrlContent = strLinkUrlContent & "<script language='javascript'>window.location.href='" & iLinkUrl & "';</script>"
    GetLinkUrlContent = strLinkUrlContent
End Function

'**************************************************
'��������GetInputerInfo
'��  �ã���ȡ����ҳ¼������ϸ��Ϣ
'��  ����InputerField
'����ֵ�����ݲ��������Ӧ�ֶε�ֵ
'**************************************************

Public Function GetInputerInfo(InputerField)
    Dim str, RsInputer, tempInputerInfo
    str = "TrueName,Title,Sex,Company,Income,UserType,Eduction,Department,Position,CompanyAddress,Operation,Email,Homepage,QQ,ICQ,MSN,Yahoo,UC,Aim,Address,OfficePhone,HomePhone,PHS,Fax,Mobile,IDCard ,Country,Province,City,ZipCode,Marriage,NativePlace,Nation,Birthday,GraduateFrom,Family,Income,InterestsOfLife,InterestsOfCulture,InterestsOfAmusement,InterestsOfSport,InterestsOfOther,Owner"
    If FoundInArr(LCase(str), LCase(InputerField), ",") = False Then
        GetInputerInfo = "�����Ƿ�"
        Exit Function
    End If
    Set RsInputer = Conn.Execute("select C.*,U.* from PE_Contacter C inner join PE_User U ON U.ContacterID = C.ContacterID  Where U.UserName = '" & rsArticle("Inputer") & "'")
    If RsInputer.EOF And RsInputer.BOF Then
        GetInputerInfo = ""
    Else
        Select Case LCase(InputerField)
        
            Case "sex"
                If RsInputer("Sex") = 0 Then tempInputerInfo = "����"
                If RsInputer("Sex") = 1 Then tempInputerInfo = "��"
                If RsInputer("Sex") = 2 Then tempInputerInfo = "Ů"
            Case "marriage"
                If RsInputer("Marriage") = 0 Then tempInputerInfo = "����"
                If RsInputer("Marriage") = 1 Then tempInputerInfo = "δ��"
                If RsInputer("Marriage") = 2 Then tempInputerInfo = "�ѻ�"
                If RsInputer("Marriage") = 3 Then tempInputerInfo = "����"
            Case "usertype"
                If RsInputer("UserType") = 1 Then tempInputerInfo = "���˻�Ա"
                If RsInputer("UserType") = 2 Then tempInputerInfo = "��ҵ��Ա"
            Case "education"
                If RsInputer("Education") = 0 Then tempInputerInfo = "Сѧ"
                If RsInputer("Education") = 1 Then tempInputerInfo = "����"
                If RsInputer("Education") = 2 Then tempInputerInfo = "����"
                If RsInputer("Education") = 3 Then tempInputerInfo = "��ר"
                If RsInputer("Education") = 4 Then tempInputerInfo = "��ר"
                If RsInputer("Education") = 5 Then tempInputerInfo = "����"
                If RsInputer("Education") = 6 Then tempInputerInfo = "˶ʿ"
                If RsInputer("Education") = 7 Then tempInputerInfo = "��ʿ"
                If RsInputer("Education") = 8 Then tempInputerInfo = "��ʿ��"
                If RsInputer("Education") = 9 Then tempInputerInfo = "����"
            Case "income"
                If RsInputer("Income") = 0 Then tempInputerInfo = "1000Ԫ����"
                If RsInputer("Income") = 1 Then tempInputerInfo = "1000--3000Ԫ"
                If RsInputer("Income") = 2 Then tempInputerInfo = "3000--6000Ԫ"
                If RsInputer("Income") = 3 Then tempInputerInfo = "6000--10000Ԫ"
                If RsInputer("Income") = 3 Then tempInputerInfo = "610000Ԫ����"
            Case Else
                tempInputerInfo = RsInputer(InputerField)
        End Select
    End If
    RsInputer.Close
    Set RsInputer = Nothing
    GetInputerInfo = tempInputerInfo
End Function


'**************************************************
'��������ReplaceContentLabel
'��  �ã��滻�������ݱ�ǩ
'��  ����strArticleHTML ----�滻��(�ֹ���ҳ�����˵�)����
'����ֵ�������õ���������
'**************************************************
Public Function ReplaceContentLabel(strArticleHTML)
    Dim tmpHtml, strTemp, arrTemp
    tmpHtml = Replace(strArticleHTML, "{$ArticleContent}", GetArticleContent(0))
    If InStr(tmpHtml, "{$ShowPageContent}") > 0 Then tmpHtml = Replace(tmpHtml, "{$ShowPageContent}","")
    If InStr(tmpHtml, "{$PageNum}") > 0 Then tmpHtml = Replace(tmpHtml, "{$PageNum}",CurrentPage)	
    regEx.Pattern = "\{\$ArticleContent\((.*?)\)\}"
    Set Matches = regEx.Execute(strArticleHTML)
    For Each Match In Matches
        tmpHtml = Replace(tmpHtml, Match.Value, GetArticleContent(PE_CLng(Match.SubMatches(0))))
    Next
    regEx.Pattern = "\{\$GetSubTitleHtml\((.*?)\)\}"
    Set Matches = regEx.Execute(tmpHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        Select Case PE_CLng(arrTemp(0))
        Case 0
            strTemp = GetContentPageGuide_List()
        Case 1
            strTemp = GetContentPageGuide_Table(PE_CLng(arrTemp(1)))
        Case Else
            strTemp = ""
        End Select
        tmpHtml = Replace(tmpHtml, Match.Value, strTemp)
    Next
    ReplaceContentLabel = tmpHtml
End Function

'*******************************************************
'��������GetContentPageGuide_List()
'��������
'���ã�����strContentPageTitleArr��������·�ҳ��ǩ���б�ʽ����Ч��
'*******************************************************
Private Function GetContentPageGuide_List()
    If rsArticle("PaginationType") <> 2 Then
        GetContentPageGuide_List = ""
        Exit Function
    End If

    Dim arrTitle, i, str
    Dim tempUrl
    If strContentPageTitleArr = "" Then
        GetContentPageGuide_List = ""
        Exit Function
    End If
    arrTitle = Split(strContentPageTitleArr, "$$$")
    If Action = "CreateArticle" Or Action = "CreateArticle2" Then '����htmlҳʱ�����ɱ�ǩ�ķ�ʽ
        str = "<option value=" & ArticleUrl
        If CurrentPage = 1 Then str = str & " selected "
        str = str & ">" & arrTitle(i) & "</option>" & vbCrLf
        
        For i = 1 To UBound(arrTitle)
            tempUrl = Left(ArticleUrl, InStrRev(ArticleUrl, ".", -1, vbBinaryCompare) - 1)
            str = str & "<option value=" & tempUrl & "_" & (i + 1) & FileExt_Item
            If CurrentPage = (i + 1) Then str = str & " selected "
            str = str & ">" & arrTitle(i) & "</option>" & vbCrLf
        Next
        GetContentPageGuide_List = "<Select Name='PageSelect' id='PageSelect' onchange=javascript:window.location=(this.options[this.selectedIndex].value)>" & str & "</Select>"
    Else
        For i = 0 To UBound(arrTitle)
            str = str & "<option value=" & i + 1
            If CurrentPage = i + 1 Then str = str & " selected "
            str = str & ">" & arrTitle(i) & "</option>"
        Next
        GetContentPageGuide_List = "<Select Name='PageSelect' id='PageSelect' onchange=""if(this.options[this.selectedIndex].value!=''){location='" & strFileName & "?ArticleID=" & ArticleID & "&Page=" & "'+this.options[this.selectedIndex].value;}"">" & str & "</Select>"
    End If
End Function

'*******************************************
'��������GetContentPageGuide_Table(iCols)
'������iCols --- ���ɵ�����
'���ã����ɱ��ʽ����
'********************************************
Private Function GetContentPageGuide_Table(iCols)
    If rsArticle("PaginationType") <> 2 Then
        GetContentPageGuide_Table = ""
        Exit Function
    End If
    If strContentPageTitleArr = "" Then
        GetContentPageGuide_Table = ""
        Exit Function
    End If

    Dim arrTitle, i, str, strTemp, strTable, m
    Dim tempUrl, TempTitle
    If iCols < 1 Then iCols = 1
    arrTitle = Split(strContentPageTitleArr, "$$$")
    strTable = "<table class='ContnetPageGuide' cellSpacing='1' cellPadding='2' border='0'><tr align='left'>"
    If Action = "CreateArticle" Or Action = "CreateArticle2" Then '����htmlҳʱ�����ɱ�ǩ�ķ�ʽ
        If CurrentPage = 1 Then
            TempTitle = "<font Color=""red"">" & Trim(arrTitle(i)) & "</font>"
        Else
            TempTitle = Trim(arrTitle(i))
        End If
        strTable = strTable & "<td class='ContnetPageGuideTD'><a href='" & ArticleUrl & "'>" & TempTitle & "</a></td>"
        If iCols = 1 Then
            strTable = strTable & "</tr><tr align='left'>"
        End If

        tempUrl = Left(ArticleUrl, InStrRev(ArticleUrl, ".", -1, vbBinaryCompare) - 1)
        For i = 1 To UBound(arrTitle)
            If CurrentPage = i + 1 Then
                TempTitle = "<font Color=""red"">" & Trim(arrTitle(i)) & "</font>"
            Else
                TempTitle = Trim(arrTitle(i))
            End If
            strTable = strTable & "<td class='ContnetPageGuideTD'><a href='" & tempUrl & "_" & (i + 1) & FileExt_Item & "'>" & TempTitle & "</a></td>"

            If ((i + 1) Mod iCols) = 0 Then
                strTable = strTable & "</tr><tr align='left'>"
            End If
        Next
    Else
        For i = 0 To UBound(arrTitle)
            If CurrentPage = i + 1 Then
                TempTitle = "<font Color=""red"">" & Trim(arrTitle(i)) & "</font>"
            Else
                TempTitle = Trim(arrTitle(i))
            End If
            strTable = strTable & "<td class='ContnetPageGuideTD'><a href='" & strFileName & "?ArticleID=" & ArticleID & "&page=" & i + 1 & "'>" & TempTitle & "</a></td>"
            If ((i + 1) Mod iCols) = 0 Then
                strTable = strTable & "</tr><tr align='left'>"
            End If
        Next
    End If
    strTable = strTable & "</tr></table>" & vbCrLf
    GetContentPageGuide_Table = strTable
End Function

Public Sub GetHTML_Index()
    Dim strTemp, arrTemp, iCols, iClassID
    Dim ArticleList_ChildClass, ArticleList_ChildClass2

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

    '�õ�����Ŀ�б�İ�����Ƶ�HTML����
    regEx.Pattern = "��ArticleList_ChildClass��([\s\S]*?)��\/ArticleList_ChildClass��"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        ArticleList_ChildClass = Match.SubMatches(0)
        strHtml = regEx.Replace(strHtml, "{$ArticleList_ChildClass}")
        
        '�õ�ÿ����ʾ������
        iCols = 1
        regEx.Pattern = "��Cols=(\d{1,2})��"
        Set Matches2 = regEx.Execute(ArticleList_ChildClass)
        ArticleList_ChildClass = regEx.Replace(ArticleList_ChildClass, "")
        For Each Match2 In Matches2
            If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
        Next
     
        '��ʼѭ�����õ���������Ŀ�б��HTML����
        ArticleList_ChildClass2 = ""
        iClassID = 0
        Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=0 and ShowOnIndex=" & PE_True & " order by RootID")
        Do While Not rsClass.EOF
            strTemp = ArticleList_ChildClass
            
            strTemp = GetCustomFromTemplate(strTemp)
            strTemp = GetListFromTemplate(strTemp)
            strTemp = GetPicFromTemplate(strTemp)
            
            strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), rsClass("ClassPurview")))
            strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", rsClass("Readme"))
            strTemp = PE_Replace(strTemp, "{$rsClass_Tips}", rsClass("Tips"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", rsClass("ClassID"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", rsClass("ClassName"))
            strTemp = Replace(strTemp, "{$ShowClassAD}", "")
            strTemp = CustomContent("Class",rsClass("Custom_Content"),strTemp)
            rsClass.MoveNext
            iClassID = iClassID + 1
            If iClassID Mod iCols = 0 And Not rsClass.EOF Then
                ArticleList_ChildClass2 = ArticleList_ChildClass2 & strTemp
                If iCols > 1 Then ArticleList_ChildClass2 = ArticleList_ChildClass2 & "</tr><tr>"
            Else
                ArticleList_ChildClass2 = ArticleList_ChildClass2 & strTemp
                If iCols > 1 Then ArticleList_ChildClass2 = ArticleList_ChildClass2 & "<td width='1'></td>"
            End If
        Loop
        rsClass.Close
        Set rsClass = Nothing
        strHtml = Replace(strHtml, "{$ArticleList_ChildClass}", ArticleList_ChildClass2)
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
    '��������
    strHtml = CustomContent("Class", Custom_Content_Class, strHtml)
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & strInstallDir & "Rss.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
    If Child > 0 Then    '�����ǰ��Ŀ������Ŀ
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(0, 0, 3, 3, 0, True))
    Else
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(ParentID, 0, 3, 3, 0, True))
    End If
    
    Dim ArticleList_CurrentClass, ArticleList_CurrentClass2, ArticleList_ChildClass, ArticleList_ChildClass2
    If Child > 0 And ClassShowType <> 2 Then    '�����ǰ��Ŀ������Ŀ
        ItemCount = PE_CLng(Conn.Execute("select Count(*) from PE_Article where ClassID=" & ClassID & "")(0))
        If ItemCount <= 0 Then     '�����ǰ��Ŀû������
            regEx.Pattern = "��ArticleList_CurrentClass��([\s\S]*?)��\/ArticleList_CurrentClass��"
            strHtml = regEx.Replace(strHtml, "") '��ȥ����ʾ��ǰ��Ŀ��ֻ���ڱ���Ŀ�������б�
        Else      '�����ǰ��Ŀ������Ŀ���ҵ�ǰ��Ŀ�����ݣ�����Ҫ��ʾ������
            regEx.Pattern = "��ArticleList_CurrentClass��([\s\S]*?)��\/ArticleList_CurrentClass��"
            Set Matches = regEx.Execute(strHtml)
            For Each Match In Matches
                ArticleList_CurrentClass = Match.SubMatches(0)
                strHtml = regEx.Replace(strHtml, "{$ArticleList_CurrentClass}")
                
                strTemp = ArticleList_CurrentClass
                strTemp = GetCustomFromTemplate(strTemp)
                strTemp = GetListFromTemplate(strTemp)
                strTemp = GetPicFromTemplate(strTemp)
                
                strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
                strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", ReadMe)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", ClassName)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", ClassID)
                
                strHtml = Replace(strHtml, "{$ArticleList_CurrentClass}", strTemp)
            Next
        End If
        
        '�õ�����Ŀ�б�İ�����Ƶ�HTML����
        regEx.Pattern = "��ArticleList_ChildClass��([\s\S]*?)��\/ArticleList_ChildClass��"
        Set Matches = regEx.Execute(strHtml)
        For Each Match In Matches
            ArticleList_ChildClass = Match.SubMatches(0)
            strHtml = regEx.Replace(strHtml, "{$ArticleList_ChildClass}")
            
            '�õ�ÿ����ʾ������
            iCols = 1
            regEx.Pattern = "��Cols=(\d{1,2})��"
            Set Matches2 = regEx.Execute(ArticleList_ChildClass)
            ArticleList_ChildClass = regEx.Replace(ArticleList_ChildClass, "")
            For Each Match2 In Matches2
                If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
            Next
            
            '��ʼѭ�����õ���������Ŀ�б��HTML����
            iClassID = 0
            Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=" & ClassID & " and IsElite=" & PE_True & " and ClassType=1 order by RootID,OrderID")
            Do While Not rsClass.EOF
                strTemp = ArticleList_ChildClass
                
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
                    ArticleList_ChildClass2 = ArticleList_ChildClass2 & strTemp
                    If iCols > 1 Then ArticleList_ChildClass2 = ArticleList_ChildClass2 & "</tr><tr>"
                Else
                    ArticleList_ChildClass2 = ArticleList_ChildClass2 & strTemp
                    If iCols > 1 Then ArticleList_ChildClass2 = ArticleList_ChildClass2 & "<td width='1'></td>"
                End If
            Loop
            rsClass.Close
            Set rsClass = Nothing

            strHtml = Replace(strHtml, "{$ArticleList_ChildClass}", ArticleList_ChildClass2)
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

Public Sub GetHtml_Article()
    totalPage = 1
    strHtml = GetCustomFromTemplate(strHtml)  '�����Ƚ����Զ����б��ǩ
    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    strHtml = PE_Replace(strHtml, "{$ArticleID}", ArticleID)
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommonLabel   '����ͨ�ñ�ǩ�������Զ����ǩ

    strHtml = GetCustomFromTemplate(strHtml)  '�����Ƚ����Զ����б��ǩ

    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    strHtml = PE_Replace(strHtml, "{$ArticleID}", ArticleID)
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = Replace(strHtml, "{$PageTitle}", ReplaceText(ArticleTitle, 2))
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())

    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    
    If InStr(strHtml, "{$MY_") > 0 Then
        Dim rsField
        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            If rsField("FieldType") = 8 Or rsField("FieldType") = 9 Then
                strHtml = PE_Replace(strHtml, rsField("LabelName"), PE_HTMLDecode(rsArticle(Trim(rsField("FieldName")))))
            Else
                strHtml = PE_Replace(strHtml, rsField("LabelName"), PE_HTMLEncode(rsArticle(Trim(rsField("FieldName")))))		
            End If	
            rsField.MoveNext
        Loop
        Set rsField = Nothing
    End If

    '�滻{$GetInputerInfo(InputerField)}��ǩ
    Dim strInputerInfo
    regEx.Pattern = "\{\$GetInputerInfo\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 0 Then
            strInputerInfo= "����ʽ��ǩ��{$GetInputerInfo(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Else
            strInputerInfo = GetInputerInfo(arrTemp(0))
        End If
        strHtml = Replace(strHtml, Match.Value, strInputerInfo)
    Next
	
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    If InStr(strHtml, "{$ClassUrl}") > 0 Then strHtml = PE_Replace(strHtml, "{$ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
    strHtml = CustomContent("Class", Custom_Content_Class, strHtml)
    If InStr(strHtml, "{$Author}") > 0 Then strHtml = PE_Replace(strHtml, "{$Author}", GetAuthorInfo(rsArticle("Author"), ChannelID))
    If InStr(strHtml, "{$CopyFrom}") > 0 Then strHtml = PE_Replace(strHtml, "{$CopyFrom}", GetCopyFromInfo(rsArticle("CopyFrom"), ChannelID))
    If InStr(strHtml, "{$Hits}") > 0 Then strHtml = PE_Replace(strHtml, "{$Hits}", GetArticleHits())
    If InStr(strHtml, "{$UpdateDate}") > 0 Then strHtml = PE_Replace(strHtml, "{$UpdateDate}", FormatDateTime(rsArticle("UpdateTime"), 2))
    strHtml = PE_Replace(strHtml, "{$UpdateTime}", rsArticle("UpdateTime"))
    strHtml = PE_Replace(strHtml, "{$Inputer}", rsArticle("Inputer"))
    strHtml = PE_Replace(strHtml, "{$Editor}", rsArticle("Editor"))
    If InStr(strHtml, "{$Stars}") > 0 Then strHtml = PE_Replace(strHtml, "{$Stars}", GetStars(rsArticle("Stars")))
    If InStr(strHtml, "{$ArticleProperty}") > 0 Then strHtml = PE_Replace(strHtml, "{$ArticleProperty}", GetArticleProperty())
    strHtml = PE_Replace(strHtml, "{$Rss}", "")
    If InStr(strHtml, "{$Keyword}") > 0 Then strHtml = PE_Replace(strHtml, "{$Keyword}", GetKeywords(",", rsArticle("Keyword")))
    If InStr(strHtml, "{$InfoPoint}") > 0 Then strHtml = PE_Replace(strHtml, "{$InfoPoint}", GetInfoPoint(rsArticle("InfoPoint")))
    If InStr(strHtml, "{$ReadPoint}") > 0 Then strHtml = PE_Replace(strHtml, "{$ReadPoint}", GetInfoPoint(rsArticle("InfoPoint")))
    If InStr(strHtml, "{$ArticleIntro}") > 0 Then strHtml = PE_Replace(strHtml, "{$ArticleIntro}", ReplaceKeyLink(ReplaceText(rsArticle("Intro"), 1)))
    '�滻{$ArticleIntro(Type,InfoLength)}��ǩ
    Dim strArticleIntro
    regEx.Pattern = "\{\$ArticleIntro\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strArticleIntro= "����ʽ��ǩ��{$ArticleIntro(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"

        Else
            Select Case PE_Clng(arrTemp(0))
            Case 1
                strArticleIntro = ReplaceKeyLink(ReplaceText(rsArticle("Intro"), 1))
            Case 2
                If PE_Clng(arrTemp(1))>0 then
                    strArticleIntro = GetSubStr(nohtml(rsArticle("Intro")),PE_Clng(arrTemp(1)),False)
                Else
                    strArticleIntro = nohtml(rsArticle("Intro"))
                End IF
            End Select
        End If
        strHtml = Replace(strHtml, Match.Value, strArticleIntro)
	Next
	
    If InStr(strHtml, "{$ArticleProtect}") > 0 Then strHtml = Replace(strHtml, "{$ArticleProtect}", GetArticleProtect())
    If InStr(strHtml, "{$ArticleTitle2}") > 0 Then strHtml = Replace(strHtml, "{$ArticleTitle2}", ReplaceText(GetInfoList_GetStrIncludePic(rsArticle("IncludePic")) & rsArticle("Title"), 2))
    If InStr(strHtml, "{$ArticleTitle}") > 0 Then strHtml = Replace(strHtml, "{$ArticleTitle}", ReplaceText(GetArticleTitle(), 2))
    If InStr(strHtml, "{$ArticleSubheading}") > 0 Then strHtml = Replace(strHtml, "{$ArticleSubheading}", GetSubheading())
    If InStr(strHtml, "{$ArticleInfo}") > 0 Then strHtml = Replace(strHtml, "{$ArticleInfo}", GetArticleInfo())
    If InStr(strHtml, "{$ArticleEditor}") > 0 Then strHtml = Replace(strHtml, "{$ArticleEditor}", GetArticleEditor())
    If InStr(strHtml, "{$PrevArticle}") > 0 Then strHtml = Replace(strHtml, "{$PrevArticle}", GetPrevArticle(200))
    If InStr(strHtml, "{$NextArticle}") > 0 Then strHtml = Replace(strHtml, "{$NextArticle}", GetNextArticle(200))
    If InStr(strHtml, "{$ArticleUrl}") > 0 Then strHtml = Replace(strHtml, "{$ArticleUrl}", GetCurArticleUrl())	
    If InStr(strHtml, "{$ArticleSign}") > 0 Then strHtml = Replace(strHtml, "{$ArticleSign}", "<iframe width='1' height='1' src='{$InstallDir}User/User_ArticleReceive.asp?ArticleID="&ArticleID&"'></iframe>")		
    If InStr(strHtml, "{$ArticleAction}") > 0 Then strHtml = Replace(strHtml, "{$ArticleAction}", GetArticleAction())
    If InStr(strHtml, "{$Vote}") > 0 Then strHtml = Replace(strHtml, "{$Vote}", GetVoteOfContent(ArticleID)) 'ͶƱ��ǩ
    If InStr(strHtml, "{$CorrelativeArticle}") > 0 Then strHtml = Replace(strHtml, "{$CorrelativeArticle}", GetCorrelative("ChannelID",0, 2, 10, 26, 1, 0, 1, False))

    Dim arrTemp
    Dim strCorrelativeArticle
    regEx.Pattern = "\{\$CorrelativeArticle\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        Select Case UBound(arrTemp)
        Case 1
            strCorrelativeArticle = GetCorrelative("ChannelID",0, 2, PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), 1, 0, 1, False)
        Case 4
            strCorrelativeArticle = GetCorrelative("ChannelID",0, 2, PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), False)
        Case 8
            strCorrelativeArticle = GetCorrelative(arrTemp(0), arrTemp(1), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CBool(arrTemp(8)))			
        Case Else
            strCorrelativeArticle = "����ʽ��ǩ��{$CorrelativeArticle(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        End Select
        strHtml = Replace(strHtml, Match.Value, strCorrelativeArticle)
    Next
End Sub

Public Sub GetHtml_Special()
    MaxPerPage = MaxPerPage_Special
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
    strHtml = PE_Replace(strHtml, "{$GetAllSpecial}", GetAllSpecial())
    
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��ר��", False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��ר��", False))
End Sub

Public Sub GetHtml_List()
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)  
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
End Sub

Public Sub GetHTML_Search()

    Dim SearchChannelID
    SearchChannelID = ChannelID
    If ChannelID > 0 Then
        strHtml = GetTemplate(ChannelID, 5, 0)
    Else
        strHtml = GetTemplate(ChannelID, 3, 0)
        ChannelID = PE_CLng(Conn.Execute("select min(ChannelID) from PE_Channel where ModuleType=1 and Disabled=" & PE_False & "")(0))
        Call GetChannel(ChannelID)
    End If
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())

    If strField <> "" Then
        regEx.Pattern = "��SearchForm��([\s\S]*?)��\/SearchForm��"
        Set Matches = regEx.Execute(strHtml)
        strHtml = regEx.Replace(strHtml, "")
    Else
        If Trim(Request.ServerVariables("QUERY_STRING")) <> "" Then
            regEx.Pattern = "��SearchForm��([\s\S]*?)��\/SearchForm��"
            Set Matches = regEx.Execute(strHtml)
            strHtml = regEx.Replace(strHtml, "")
        Else
            regEx.Pattern = "��ShowResult��([\s\S]*?)��\/ShowResult��"
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

    strHtml = Replace(strHtml, "��ShowResult��", "")
    strHtml = Replace(strHtml, "��/ShowResult��", "")
    strHtml = Replace(strHtml, "��SearchForm��", "")
    strHtml = Replace(strHtml, "��/SearchForm��", "")
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

Private Function GetInfoList_GetStrCommentLink(ShowCommentLink, ShowCommentLink_ListItem, ArticleID_ListItem)
    If ShowCommentLink = True And ShowCommentLink_ListItem = True Then
        GetInfoList_GetStrCommentLink = "&nbsp;&nbsp;<a href=""" & ChannelUrl_ASPFile & "/Comment.asp?Action=ShowAll&ArticleID=" & ArticleID_ListItem & """ target='_blank'>" & strComment & "</a>"
    Else
        GetInfoList_GetStrCommentLink = ""
    End If
End Function

Public Sub ShowFavorite()
    Response.Write "<table width='100%' cellpadding='2' cellspacing='1' border='0' class='border'>"
    Response.Write "  <tr class='title' align='center'><td width='30'>ѡ��</td><td>" & ChannelShortName & "����</td><td width='100'>����</td><td width='80'>����ʱ��</td><td width='80'>����</td></tr>"
    
    Dim sqlFavorite, rsFavorite, iCount, strLink
    iCount = 0
    
    sqlFavorite = "select A.ChannelID,A.ArticleID,A.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,A.InfoPurview,A.Title,A.Author,A.UpdateTime,A.IncludePic,A.InfoPoint from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where A.Deleted=" & PE_False & " and A.Status=3 and A.ReceiveType=0 "
    sqlFavorite = sqlFavorite & " and ArticleID in (select InfoID from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & ")"
    sqlFavorite = sqlFavorite & " order by A.ArticleID desc"

    Set rsFavorite = Server.CreateObject("ADODB.Recordset")
    rsFavorite.Open sqlFavorite, Conn, 1, 1
    If rsFavorite.BOF And rsFavorite.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td height='50' colspan='20' align='center'>û���ղ��κ�" & ChannelShortName & "</td></tr>"
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
            strLink = strLink & "<a href='" & GetArticleUrl(rsFavorite("ParentDir"), rsFavorite("ClassDir"), rsFavorite("UpdateTime"), rsFavorite("ArticleID"), rsFavorite("ClassPurview"), rsFavorite("InfoPurview"), rsFavorite("InfoPoint")) & "' target='_blank'>" & rsFavorite("Title") & "</a>"
            
            Response.Write "<tr class='tdbg'>"
            Response.Write "<td align='center' width='30'><input type='checkbox' name='InfoID' value='" & rsFavorite("ArticleID") & "'></td>"
            Response.Write "<td align='left'>" & strLink & "</td>"
            Response.Write "<td width='100' align='center'>" & rsFavorite("Author") & "</td>"
            Response.Write "<td width='80' align='right'>" & Year(rsFavorite("UpdateTime")) & "-" & Right("0" & Month(rsFavorite("UpdateTime")), 2) & "-" & Right("0" & Day(rsFavorite("UpdateTime")), 2) & "</td>"
            Response.Write "<td width='80' align='center'><a href='User_Favorite.asp?Action=Remove&ChannelID=" & ChannelID & "&InfoID=" & rsFavorite("ArticleID") & "' onclick=""return confirm('ȷʵ�����ղش�" & ChannelShortName & "��');"">ȡ���ղ�</a></td>"
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
    XmlText_Class = XmlText("Article", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Article", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>
