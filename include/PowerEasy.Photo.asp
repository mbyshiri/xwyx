<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim PhotoID, rsPhoto, PhotoName, PhotoUrl

Class Photo

'��������ȫ�ֵı���
Private rsClass

'��ʼ����Ҫ�õ���һЩ����
Public Sub Init()
    FoundErr = False
    ErrMsg = ""
    PrevChannelID = ChannelID
    ChannelShortName = "ͼƬ"
    
    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
         
    '*****************************
    '��ȡ���԰��е��ַ�����
    ChannelShortName = XmlText_Class("ChannelShortName", "ͼƬ")
    strListStr_Font = XmlText_Class("PhotoList/UpdateTimeColor_New", "color=""red""")
    strTop = XmlText_Class("PhotoList/t4", "�̶�")
    strElite = XmlText_Class("PhotoList/t3", "�Ƽ�")
    strCommon = XmlText_Class("PhotoList/t5", "��ͨ")
    strHot = XmlText_Class("PhotoList/t7", "�ȵ�")
    strNew = XmlText_Class("PhotoList/t6", "����")
    strTop2 = XmlText_Class("PhotoList/Top", " ��")
    strElite2 = XmlText_Class("PhotoList/Elite", " ��")
    strHot2 = XmlText_Class("PhotoList/Hot", " ��")
    Character_Author = XmlText("Photo", "Include/Author", "[{$Text}]")
    Character_Date = XmlText("Photo", "Include/Date", "[{$Text}]")
    Character_Hits = XmlText("Photo", "Include/Hits", "[{$Text}]")
    Character_Class = XmlText("Photo", "Include/ClassChar", "[{$Text}]")
    SearchResult_Content_NoPurview = XmlText("BaseText", "SearchPurviewContent", "��������Ҫ��ָ��Ȩ�޲ſ���Ԥ��")
    SearchResult_ContentLenth = PE_CLng(XmlText_Class("ShowSearch/Content_Lenght", "200"))
    strList_Content_Div = XmlText_Class("PhotoList/Content_DIV", "style=""padding:0px 20px""")
    strList_Title = R_XmlText_Class("PhotoList/Title", "{$ChannelShortName}���ƣ�{$PhotoName}{$br}��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�{$Author}{$br}����ʱ�䣺{$UpdateTime}")
    strComment = XmlText_Class("PhotoList/CommentLink", "<font color=""red"">����</font>")
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
'response.write "ChannelID =" & ChannelID
'response.write "ChannelUrl =" & ChannelUrl
End Sub


'=================================================
'��������ShowChannelCount
'��  �ã���ʾƵ��ͳ����Ϣ
'��  ������
'=================================================
Private Function GetChannelCount()
    GetChannelCount = Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ChannelCount", "{$ChannelShortName}������ {$ItemChecked_Channel} {$ChannelItemUnit}<br>����{$ChannelShortName}�� {$UnItemChecked} {$ChannelItemUnit}<br>���������� {$CommentCount_Channel} ��<br>ר�������� {$SpecialCount_Channel} ��<br>{$ChannelShortName}�鿴�� {$HitsCount_Channel} �˴�<br>"), "{$ItemChecked_Channel}", ItemChecked_Channel), "{$ChannelItemUnit}", ChannelItemUnit), "{$UnItemChecked}", treatAuditing("Photo", ChannelID)), "{$CommentCount_Channel}", CommentCount_Channel), "{$SpecialCount_Channel}", SpecialCount_Channel), "{$HitsCount_Channel}", "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?Action=Count'></script>")
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
        strSql = strSql & " from PE_InfoS I inner join (PE_Photo P left join PE_Class C on P.ClassID=C.ClassID) on I.ItemID=P.PhotoID"
    Else
        strSql = strSql & " from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID"
    End If
    strSql = strSql & " where P.Deleted=" & PE_False & " and P.Status=3"
    If iChannelID > 0 Then
        strSql = strSql & " and P.ChannelID=" & iChannelID
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
            strSql = strSql & " and P.ClassID in (" & FilterArrNull(arrClassID, ",") & ")"
        Else
            If PE_CLng(arrClassID) > 0 Then strSql = strSql & " and P.ClassID=" & PE_CLng(arrClassID)
        End If
    End If
    If iSpecialID > 0 Then
        strSql = strSql & " and I.ModuleType=3 and I.SpecialID=" & iSpecialID
    End If
    If IsHot = True Then
        strSql = strSql & " and P.Hits>=" & HitsOfHot
    End If
    If IsElite = True Then
        strSql = strSql & " and P.Elite=" & PE_True
    End If
    If Trim(Author) <> "" Then
        strSql = strSql & " and P.Author='" &  ReplaceBadChar(Author) & "'"
    End If
    If DateNum > 0 Then
        strSql = strSql & " and DateDiff(" & PE_DatePart_D & ",P.UpdateTime," & PE_Now & ")<" & DateNum
    End If

    If IsPicUrl = True Then
        strSql = strSql & " and P.PhotoThumb<>'' "
    End If

    strSql = strSql & " order by P.OnTop " & PE_OrderType & ","
    Select Case PE_CLng(OrderType)
    Case 1, 2

    Case 3
        strSql = strSql & "P.UpdateTime desc,"
    Case 4
        strSql = strSql & "P.UpdateTime asc,"
    Case 5
        strSql = strSql & "P.Hits desc,"
    Case 6
        strSql = strSql & "P.Hits asc,"
    Case 7
        strSql = strSql & "P.CommentCount desc,"
    Case 8
        strSql = strSql & "P.CommentCount asc,"
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
        strSql = strSql & "P.PhotoID " & IDOrder
    End If
    GetSqlStr = strSql
End Function

'=================================================
'��������GetPhotoList
'��  �ã���ʾͼƬ���Ƶ���Ϣ
'��  ����
'0        iChannelID ---- Ƶ��ID
'1        arrClassID ---- ��ĿID���飬0Ϊ������Ŀ
'2        IncludeChild ---- �Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'3        iSpecialID ---- ר��ID��0Ϊ����ͼƬ������ר��ͼƬ�������Ϊ����0����ֻ��ʾ��Ӧר���ͼƬ
'4        UrlType ---- ���ӵ�ַ���ͣ�0Ϊ���·����1Ϊ����ַ�ľ���·���������⹫����4.03ʱΪShowAllPhoto
'5        PhotoNum ---- ͼƬ����������0����ֻ��ѯǰ����ͼƬ
'6        IsHot ---- �Ƿ�������ͼƬ��TrueΪֻ��ʾ����ͼƬ��FalseΪ��ʾ����ͼƬ
'7        IsElite ---- �Ƿ����Ƽ�ͼƬ��TrueΪֻ��ʾ�Ƽ�ͼƬ��FalseΪ��ʾ����ͼƬ
'8        Author ---- ���������������Ϊ�գ���ֻ��ʾָ�����ߵ�ͼƬ����������ͼƬ��
'9        DateNum ---- ���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�ͼƬ
'10       OrderType ---- ����ʽ��1--��ͼƬID����2--��ͼƬID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'11       ShowType ---- ��ʾ��ʽ��1Ϊ��ͨ��ʽ��2Ϊ���ʽ��3Ϊ�������ʽ��4Ϊ���DIV��ʽ��5Ϊ���RSS��ʽ
'12       TitleLen ---- ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ��������
'13       ContentLen ---- ͼƬ�������ַ�����һ������=����Ӣ���ַ���Ϊ0ʱ����ʾ��
'14       ShowClassName ---- �Ƿ���ʾ������Ŀ���ƣ�TrueΪ��ʾ��FalseΪ����ʾ
'15       ShowPropertyType ---- ��ʾͼƬ���ԣ��̶�/�Ƽ�/��ͨ���ķ�ʽ��0Ϊ����ʾ��1ΪСͼƬ��2Ϊ����
'16       ShowAuthor ---- �Ƿ���ʾͼƬ���ߣ�TrueΪ��ʾ��FalseΪ����ʾ
'17       ShowDateType ---- ��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���գ�3Ϊ�ԡ���-�ա���ʽ��ʾ���ա�
'18       ShowHits ---- �Ƿ���ʾͼƬ�������TrueΪ��ʾ��FalseΪ����ʾ
'19       ShowHotSign ---- �Ƿ���ʾ����ͼƬ��־��TrueΪ��ʾ��FalseΪ����ʾ
'20       ShowNewSign ---- �Ƿ���ʾ��ͼƬ��־��TrueΪ��ʾ��FalseΪ����ʾ
'21       ShowTips ---- �Ƿ���ʾ���ߡ��������ڡ�������ȸ�����ʾ��Ϣ��TrueΪ��ʾ��FalseΪ����ʾ
'22       UsePage ---- �Ƿ��ҳ��ʾ��TrueΪ��ҳ��ʾ��FalseΪ����ҳ��ʾ��ÿҳ��ʾ��ͼƬ������MaxPerPageָ��
'23       OpenType ---- ͼƬ�򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�
'24       Cols ---- ÿ�е������������������ͻ��С�
'25       CssNameA ---- �б����������ӵ��õ�CSS����
'26       CssName1 ---- �б��������е�CSSЧ��������
'27       CssName2 ---- �б���ż���е�CSSЧ��������
'=================================================
Public Function GetPhotoList(iChannelID, arrClassID, IncludeChild, iSpecialID, UrlType, PhotoNum, IsHot, IsElite, Author, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, UsePage, OpenType, Cols, CssNameA, CssName1, CssName2)
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
        GetPhotoList = ErrMsg
        Exit Function
    End If

    sqlInfo = "select"
    If PhotoNum > 0 Then
        sqlInfo = sqlInfo & " top " & PhotoNum
    End If
    sqlInfo = sqlInfo & " P.ChannelID,P.ClassID,P.PhotoID,P.PhotoName,P.Author,P.UpdateTime,P.Hits,P.OnTop,P.Elite,P.InfoPurview,P.InfoPoint"
    If ContentLen > 0 Then
        sqlInfo = sqlInfo & ",P.PhotoIntro"
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
        GetPhotoList = strInfoList
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
        InfoUrl = GetPhotoUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("UpdateTime"), rsInfoList("PhotoID"), rsInfoList("ClassPurview"), rsInfoList("InfoPurview"), rsInfoList("InfoPoint"))
        strTitle = GetInfoList_GetStrTitle(rsInfoList("PhotoName"), TitleLen, 0, "")
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
            strLink = strLink & GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, CssNameA, strTitle, InfoUrl, rsInfoList("PhotoName"), rsInfoList("Author"), rsInfoList("UpdateTime"))

            strContent = ""
            Select Case PE_CLng(ShowType)
            Case 1, 3, 4
                If ContentLen > 0 Then
                    strContent = strContent & "<div " & strList_Content_Div & ">"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("PhotoIntro"), "")
                    strContent = strContent & "</div>"
                End If
            Case 2
                If ContentLen > 0 Then
                    strContent = strContent & "<tr><td colspan=""10"" class=""" & CssName & """>"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("PhotoIntro"), "")
                    strContent = strContent & "</td></tr>"
                End If
            End Select

        ElseIf ShowType = 5 Then

            strTitle = xml_nohtml(strTitle)
            strLink = InfoUrl
            If ContentLen > 0 Then
                strContent = Left(Replace(Replace(xml_nohtml(rsInfoList("PhotoIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen)
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
        Case 4 '���DIV
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
        Case 5 '���RSS
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
    GetPhotoList = strInfoList
End Function

Private Function GetInfoList_GetArrClassID(strClassID, IncludeChild)
    Dim trs, arrClassID, SingleClassID
    If InStr(strClassID, ",") > 0 Then
        arrClassID = strClassID
    Else
        SingleClassID = PE_CLng(strClassID)
        If IncludeChild = True Then
            Set trs = Conn.Execute("select arrChildID from PE_Class where ClassID=" & SingleClassID & "")
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
        Else
            arrClassID = SingleClassID
        End If
    End If
    GetInfoList_GetArrClassID = arrClassID
End Function

'=================================================
'��������GetPicPhoto
'��  �ã���ʾͼƬͼƬ
'��  ����
'0        iChannelID ---- Ƶ��ID
'1        arrClassID ---- ��ĿID���飬0Ϊ������Ŀ
'2        IncludeChild ---- �Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'3        iSpecialID ---- ר��ID��0Ϊ����ͼƬ������ר��ͼƬ�������Ϊ����0����ֻ��ʾ��Ӧר���ͼƬ
'4        PhotoNum ---- �����ʾ���ٸ�ͼƬ
'5        IsHot ---- �Ƿ�������ͼƬ
'6        IsElite ---- �Ƿ����Ƽ�ͼƬ
'7        DateNum ---- ���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�ͼƬ
'8        OrderType ---- ����ʽ��1--��ͼƬID����2--��ͼƬID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'9        ShowType ---- ��ʾ��ʽ��1ΪͼƬ+����+���ݼ�飺�������У�2Ϊ��ͼƬ+���⣺�������У�+���ݼ�飺�������У�3ΪͼƬ+������+���ݼ�飺�������У����������У�4Ϊ���DIV��ʽ��5Ϊ���RSS��ʽ
'10       ImgWidth ---- ͼƬ���
'11       ImgHeight ---- ͼƬ�߶�
'12       TitleLen ---- ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ���⣻��Ϊ-1������ʾ��������
'13       ContentLen ---- ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ���ݼ��
'14       ShowTips ---- �Ƿ���ʾ���ߡ�����ʱ�䡢���������ʾ��Ϣ��TrueΪ��ʾ��FalseΪ����ʾ
'15       Cols ---- ÿ�е������������������ͻ��С�
'16       UrlType ---- ���ӵ�ַ���ͣ�0Ϊ���·����1Ϊ����ַ�ľ���·����
'=================================================
Public Function GetPicPhoto(iChannelID, arrClassID, IncludeChild, iSpecialID, PhotoNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType)
    Dim sqlPic, rsPic, iCount, strPic, strLink, strAuthor, InfoUrl
    Dim strPhotoThumb, strLink_PhotoThumb, strTitle, strLink_Title, strContent, strLink_Content

    iCount = 0
    PhotoNum = PE_CLng(PhotoNum)
    ShowType = PE_CLng(ShowType)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)

    If PhotoNum < 0 Or PhotoNum >= 100 Then PhotoNum = 10
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
        GetPicPhoto = ErrMsg
        Exit Function
    End If

    sqlPic = "select"
    If PhotoNum > 0 Then
        sqlPic = sqlPic & " top " & PhotoNum
    End If
    sqlPic = sqlPic & " P.ChannelID,P.ClassID,P.PhotoID,P.PhotoName,P.Author,P.UpdateTime,P.Hits,P.InfoPurview,P.InfoPoint,P.PhotoThumb"
    If ContentLen > 0 Then
        sqlPic = sqlPic & ",P.PhotoIntro"
    End If
    sqlPic = sqlPic & ",C.ClassName,C.ClassDir,C.ParentDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    If ShowType < 4 Then strPic = "<table width='100%' cellpadding='0' cellspacing='5' border='0' align='center'><tr valign='top'>"
    If rsPic.BOF And rsPic.EOF Then
        If PhotoNum = 0 Then totalPut = 0
        If ShowType < 4 Then
            strPic = strPic & "<td align='center'><img class='pic3' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>" & R_XmlText_Class("PicPhoto/NoFound", "û���κ�{$ChannelShortName}") & "</td></tr></table>"
        ElseIf ShowType = 4 Then
            strPic = strPic & "<div class=""pic_photo""><img class='pic3' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>" & R_XmlText_Class("PicPhoto/NoFound", "û���κ�{$ChannelShortName}") & "</div>"
        End If
        rsPic.Close
        Set rsPic = Nothing
        GetPicPhoto = strPic
        Exit Function
    End If

    If PhotoNum = 0 And ShowType < 5 Then
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
            InfoUrl = GetPhotoUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("PhotoID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint"))
            strPhotoThumb = GetPhotoThumb(rsPic("PhotoThumb"), ImgWidth, ImgHeight)
            strLink_PhotoThumb = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strPhotoThumb, InfoUrl, rsPic("PhotoName"), rsPic("Author"), rsPic("UpdateTime"))

            If ShowType = 4 Then
                strPic = strPic & "<div class=""pic_photo"">" & vbCrLf
                strPic = strPic & "<div class=""pic_photo_img"">" & strLink_PhotoThumb & "</div>" & vbCrLf
            Else
                strPic = strPic & "<td align='center'>"
                strPic = strPic & strLink_PhotoThumb
            End If

            If TitleLen <> 0 Then
                strTitle = GetInfoList_GetStrTitle(rsPic("PhotoName"), TitleLen, 0, "")
                strLink_Title = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strTitle, InfoUrl, rsPic("PhotoName"), rsPic("Author"), rsPic("UpdateTime"))
                Select Case PE_CLng(ShowType)
                Case 1, 2
                    strPic = strPic & "<br>" & strLink_Title
                Case 3
                    strPic = strPic & "</td><td valign='top' align='left'>" & strLink_Title
                Case 4
                    strPic = strPic & "<div class=""pic_photo_title"">" & strLink_Title & "</div>" & vbCrLf
                End Select
            End If
            If ContentLen > 0 Then
                strContent = Left(Replace(Replace(nohtml(rsPic("PhotoIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen) & "����"
                strLink_Content = GetInfoList_GetStrInfoLink(strList_Title, ShowTips, 1, "", strContent, InfoUrl, rsPic("PhotoName"), rsPic("Author"), rsPic("UpdateTime"))
                Select Case PE_CLng(ShowType)
                Case 1, 3
                    strPic = strPic & "<br><div align='left'>" & strLink_Content & "</div>"
                Case 2
                    strPic = strPic & "</td><td valign='top' align='left'>" & strLink_Content
                Case 4
                    strPic = strPic & "<div class=""pic_photo_content"">" & strLink_Content & "</div>" & vbCrLf
                End Select
            End If
            If ShowType = 4 Then
                strPic = strPic & "</div>" & vbCrLf
            Else
                strPic = strPic & "</td>"
            End If
        Else
            strTitle = GetInfoList_GetStrTitle(rsPic("PhotoName"), TitleLen, 0, "")
            strLink = GetPhotoUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("PhotoID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint"))
            strAuthor = GetInfoList_GetStrAuthor_RSS(rsPic("Author"))
            If ContentLen > 0 Then
                strContent = Left(Replace(Replace(xml_nohtml(rsPic("PhotoIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen)
            End If
            strPic = strPic & GetInfoList_GetStrRSS(xml_nohtml(strTitle), strLink, strContent, strAuthor, xml_nohtml(rsPic("ClassName")), rsPic("UpdateTime"))
        End If
        rsPic.MoveNext
        iCount = iCount + 1
        If PhotoNum = 0 And iCount >= MaxPerPage Then Exit Do
        If ((iCount Mod Cols = 0) And (Not rsPic.EOF)) And ShowType < 4 Then strPic = strPic & "</tr><tr valign='top'>"
    Loop

    If ShowType < 4 Then strPic = strPic & "</tr></table>"
    rsPic.Close
    Set rsPic = Nothing
    If ShowType = 5 And RssCodeType = False Then strPic = unicode(strPic)
    GetPicPhoto = strPic
End Function

'=================================================
'��������GetSlidePicPhoto
'��  �ã��Իõ�ƬЧ����ʾͼƬ
'��  ����
'0        iChannelID ---- Ƶ��ID
'1        arrClassID ---- ��ĿID���飬0Ϊ������Ŀ
'2        IncludeChild ---- �Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'3        iSpecialID ---- ר��ID��0Ϊ����ͼƬ������ר��ͼƬ�������Ϊ����0����ֻ��ʾ��Ӧר���ͼƬ
'4        PhotoNum ---- �����ʾ���ٸ�ͼƬ
'5        IsHot ---- �Ƿ�������ͼƬ
'6        IsElite ---- �Ƿ����Ƽ�ͼƬ
'7        DateNum ---- ���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�ͼƬ
'8        OrderType ---- ����ʽ��1--��ͼƬID����2--��ͼƬID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'9        ImgWidth ---- ͼƬ���
'10       ImgHeight ---- ͼƬ�߶�
'11       TitleLen ---- ͼƬ�����������ƣ�0Ϊ����ʾ��-1Ϊ��ʾ��������
'12       iTimeOut ---- Ч���任���ʱ�䣬�Ժ���Ϊ��λ
'13       effectID ---- ͼƬת��Ч����0��22ָ��ĳһ����Ч��23��ʾ���Ч��
'=================================================
Public Function GetSlidePicPhoto(iChannelID, arrClassID, IncludeChild, iSpecialID, PhotoNum, IsHot, IsElite, DateNum, OrderType, ImgWidth, ImgHeight, TitleLen, iTimeOut, effectID)
    Dim sqlPic, rsPic, i, strPic
    Dim PhotoThumb, strTitle

    PhotoNum = PE_CLng(PhotoNum)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)

    If PhotoNum <= 0 Or PhotoNum > 100 Then PhotoNum = 10
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
        GetSlidePicPhoto = ErrMsg
        Exit Function
    End If

    sqlPic = "select top " & PhotoNum & " P.ChannelID,P.ClassID,P.PhotoID,P.PhotoName,P.UpdateTime,P.InfoPurview,P.InfoPoint,P.PhotoThumb"
    sqlPic = sqlPic & ",C.ClassName,C.ClassDir,C.ParentDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)

    Dim ranNum
    Randomize
    ranNum = Int(900 * Rnd) + 100
    strPic = "<script language=JavaScript>" & vbCrLf
    strPic = strPic & "<!--" & vbCrLf
    strPic = strPic & "var SlidePic_" & ranNum & " = new SlidePic_Photo(""SlidePic_" & ranNum & """);" & vbCrLf
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
        If Left(rsPic("PhotoThumb"), 1) <> "/" And InStr(rsPic("PhotoThumb"), "://") <= 0 Then
            PhotoThumb = ChannelUrl & "/" & UploadDir & "/" & rsPic("PhotoThumb")
        Else
            PhotoThumb = rsPic("PhotoThumb")
        End If
        If TitleLen = -1 Then
            strTitle = rsPic("PhotoName")
        Else
            strTitle = GetSubStr(rsPic("PhotoName"), TitleLen, ShowSuspensionPoints)
        End If
        
        strPic = strPic & "var oSP = new objSP_Photo();" & vbCrLf
        strPic = strPic & "oSP.ImgUrl         = """ & PhotoThumb & """;" & vbCrLf
        strPic = strPic & "oSP.LinkUrl        = """ & GetPhotoUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("PhotoID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint")) & """;" & vbCrLf
        strPic = strPic & "oSP.Title         = """ & strTitle & """;" & vbCrLf
        strPic = strPic & "SlidePic_" & ranNum & ".Add(oSP);" & vbCrLf
        
        rsPic.MoveNext
    Loop
    strPic = strPic & "SlidePic_" & ranNum & ".Show();" & vbCrLf
    strPic = strPic & "//-->" & vbCrLf
    strPic = strPic & "</script>" & vbCrLf
    
    rsPic.Close
    Set rsPic = Nothing
    GetSlidePicPhoto = strPic
End Function

Private Function JS_SlidePic()
    Dim strJS, LinkTarget
    LinkTarget = XmlText_Class("SlidePicPhoto/LinkTarget", "_blank")
    strJS = strJS & "<script language=""JavaScript"">" & vbCrLf
    strJS = strJS & "<!--" & vbCrLf
    strJS = strJS & "function objSP_Photo() {this.ImgUrl=""""; this.LinkUrl=""""; this.Title="""";}" & vbCrLf
    strJS = strJS & "function SlidePic_Photo(_id) {this.ID=_id; this.Width=0;this.Height=0; this.TimeOut=5000; this.Effect=23; this.TitleLen=0; this.PicNum=-1; this.Img=null; this.Url=null; this.Title=null; this.AllPic=new Array(); this.Add=SlidePic_Photo_Add; this.Show=SlidePic_Photo_Show; this.LoopShow=SlidePic_Photo_LoopShow;}" & vbCrLf
    strJS = strJS & "function SlidePic_Photo_Add(_SP) {this.AllPic[this.AllPic.length] = _SP;}" & vbCrLf
    strJS = strJS & "function SlidePic_Photo_Show() {" & vbCrLf
    strJS = strJS & "  if(this.AllPic[0] == null) return false;" & vbCrLf
    strJS = strJS & "  document.write(""<div align='center'><a id='Url_"" + this.ID + ""' href='' target='" & LinkTarget & "'><img id='Img_"" + this.ID + ""' style='width:"" + this.Width + ""px; height:"" + this.Height + ""px; filter: revealTrans(duration=2,transition=23);' src='javascript:null' border='0'></a>"");" & vbCrLf
    strJS = strJS & "  if(this.TitleLen != 0) {document.write(""<br><span id='Title_"" + this.ID + ""'></span></div>"");}" & vbCrLf
    strJS = strJS & "  else{document.write(""</div>"");}" & vbCrLf
    strJS = strJS & "  this.Img = document.getElementById(""Img_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Url = document.getElementById(""Url_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Title = document.getElementById(""Title_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.LoopShow();" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function SlidePic_Photo_LoopShow() {" & vbCrLf
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

Private Function GetPhotoThumb(PhotoThumb, PhotoThumbWidth, PhotoThumbHeight)
    Dim strPhotoThumb, FileType, strPicUrl

    If PhotoThumb = "" Then
        strPhotoThumb = strPhotoThumb & "<img src='" & strPicUrl & strInstallDir & "images/nopic.gif' "
        If PhotoThumbWidth > 0 Then strPhotoThumb = strPhotoThumb & " width='" & PhotoThumbWidth & "'"
        If PhotoThumbHeight > 0 Then strPhotoThumb = strPhotoThumb & " height='" & PhotoThumbHeight & "'"
        strPhotoThumb = strPhotoThumb & " border='0'>"
    Else
        FileType = LCase(Mid(PhotoThumb, InStrRev(PhotoThumb, ".") + 1))
        If Left(PhotoThumb, 1) <> "/" And InStr(PhotoThumb, "://") <= 0 Then
            strPicUrl = ChannelUrl & "/" & UploadDir & "/" & PhotoThumb
        Else
            strPicUrl = PhotoThumb
        End If
        If FileType = "swf" Then
            strPhotoThumb = strPhotoThumb & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' "
            If PhotoThumbWidth > 0 Then strPhotoThumb = strPhotoThumb & " width='" & PhotoThumbWidth & "'"
            If PhotoThumbHeight > 0 Then strPhotoThumb = strPhotoThumb & " height='" & PhotoThumbHeight & "'"
            strPhotoThumb = strPhotoThumb & "><param name='movie' value='" & strPicUrl & "'><param name='quality' value='high'><embed src='" & strPicUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' "
            If PhotoThumbWidth > 0 Then strPhotoThumb = strPhotoThumb & " width='" & PhotoThumbWidth & "'"
            If PhotoThumbHeight > 0 Then strPhotoThumb = strPhotoThumb & " height='" & PhotoThumbHeight & "'"
            strPhotoThumb = strPhotoThumb & "></embed></object>"
        ElseIf FileType = "gif" Or FileType = "jpg" Or FileType = "jpeg" Or FileType = "jpe" Or FileType = "bmp" Or FileType = "png" Then
            strPhotoThumb = strPhotoThumb & "<img class='pic3' src='" & strPicUrl & "' "
            If PhotoThumbWidth > 0 Then strPhotoThumb = strPhotoThumb & " width='" & PhotoThumbWidth & "'"
            If PhotoThumbHeight > 0 Then strPhotoThumb = strPhotoThumb & " height='" & PhotoThumbHeight & "'"
            strPhotoThumb = strPhotoThumb & " border='0'>"
        Else
            strPhotoThumb = strPhotoThumb & "<img class='pic3' src='" & strInstallDir & "images/nopic.gif' "
            If PhotoThumbWidth > 0 Then strPhotoThumb = strPhotoThumb & " width='" & PhotoThumbWidth & "'"
            If PhotoThumbHeight > 0 Then strPhotoThumb = strPhotoThumb & " height='" & PhotoThumbHeight & "'"
            strPhotoThumb = strPhotoThumb & " border='0'>"
        End If
    End If
    GetPhotoThumb = strPhotoThumb
End Function

Private Function GetSearchResultIDArr(iChannelID)
    Dim sqlSearch, rsSearch
    Dim rsField
    Dim PhotoNum, arrPhotoID

    If PE_CLng(SearchResultNum) > 0 Then
        sqlSearch = "select top " & PE_CLng(SearchResultNum) & " PhotoID "
    Else
        sqlSearch = "select PhotoID "
    End If
    sqlSearch = sqlSearch & " from PE_Photo where Deleted=" & PE_False & " and Status=3"
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
        sqlSearch = sqlSearch & " and PhotoID in (select ItemID from PE_InfoS where SpecialID=" & SpecialID & ")"
    End If
    If strField <> "" Then  '��ͨ����
        Select Case strField
            Case "Title", "PhotoName"
                sqlSearch = sqlSearch & SetSearchString("PhotoName")
            Case "Content", "PhotoIntro"
                sqlSearch = sqlSearch & SetSearchString("PhotoIntro")
            Case "Author"
                sqlSearch = sqlSearch & SetSearchString("Author")
            Case "Inputer"
                sqlSearch = sqlSearch & SetSearchString("Inputer")
            Case "Keywords"
                sqlSearch = sqlSearch & SetSearchString("Keyword")
            Case Else  '�Զ����ֶ�
                Set rsField = Conn.Execute("select Title from PE_Field where (ChannelID=-3 or ChannelID=" & iChannelID & ") and FieldName='" & ReplaceBadChar(strField) & "'")
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
        Dim PhotoName, PhotoIntro, Author, CopyFrom, Keyword2, LowInfoPoint, HighInfoPoint, BeginDate, EndDate, Inputer
        PhotoName = Trim(Request("PhotoName"))
        PhotoIntro = Trim(Request("PhotoIntro"))
        Author = Trim(Request("Author"))
        CopyFrom = Trim(Request("CopyFrom"))
        Keyword2 = Trim(Request("Keywords"))
        LowInfoPoint = PE_CLng(Request("LowInfoPoint"))
        HighInfoPoint = PE_CLng(Request("HighInfoPoint"))
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        Inputer = Trim(Request("Inputer"))
        strFileName = "Search.asp?ModuleName=Photo&ClassID=" & ClassID & "&SpecialID=" & SpecialID
        If PhotoName <> "" Then
            PhotoName = ReplaceBadChar(PhotoName)
            strFileName = strFileName & "&PhotoName=" & PhotoName
            sqlSearch = sqlSearch & " and PhotoName like '%" & PhotoName & "%' "
        End If
        If PhotoIntro <> "" Then
            PhotoIntro = ReplaceBadChar(PhotoIntro)
            strFileName = strFileName & "&PhotoIntro=" & PhotoIntro
            sqlSearch = sqlSearch & " and PhotoIntro like '%" & PhotoIntro & "%'"
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

        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-3 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            If Trim(Request(rsField("FieldName"))) <> "" Then
                strFileName = strFileName & "&" & Trim(rsField("FieldName")) & "=" & ReplaceBadChar(Trim(Request(rsField("FieldName"))))
                sqlSearch = sqlSearch & " and " & Trim(rsField("FieldName")) & " like '%" & ReplaceBadChar(Trim(Request(rsField("FieldName")))) & "%' "
            End If
            rsField.MoveNext
        Loop
        Set rsField = Nothing
        
    End If
    sqlSearch = sqlSearch & " order by PhotoID desc"
    arrPhotoID = ""
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
        PhotoNum = 0
        Do While Not rsSearch.EOF
            If arrPhotoID = "" Then
                arrPhotoID = rsSearch(0)
            Else
                arrPhotoID = arrPhotoID & "," & rsSearch(0)
            End If
            PhotoNum = PhotoNum + 1
            If PhotoNum >= MaxPerPage Then Exit Do
            rsSearch.MoveNext
        Loop
    End If
    rsSearch.Close
    Set rsSearch = Nothing

    GetSearchResultIDArr = arrPhotoID
End Function


'=================================================
'��������GetSearchResult
'��  �ã���ҳ��ʾ�������
'��  ������
'=================================================
Private Function GetSearchResult(iChannelID)
    Dim sqlSearch, rsSearch, iCount, PhotoNum, arrPhotoID, strSearchResult, Content
    strSearchResult = ""
    arrPhotoID = GetSearchResultIDArr(iChannelID)
    If arrPhotoID = "" Then
        GetSearchResult = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "û�л�û���ҵ��κ�{$ChannelShortName}") & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If

    PhotoNum = 1
    sqlSearch = "select P.ChannelID,P.PhotoID,P.PhotoName,P.Author,P.UpdateTime,P.Hits,P.InfoPurview,P.InfoPoint,P.PhotoIntro,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID where PhotoID in (" & arrPhotoID & ") order by PhotoID desc"
    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1
    Do While Not rsSearch.EOF
        If iChannelID = 0 Then
            If rsSearch("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsSearch("ChannelID"))
                PrevChannelID = rsSearch("ChannelID")
            End If
        End If
        
        strSearchResult = strSearchResult & "<b>" & CStr(MaxPerPage * (CurrentPage - 1) + PhotoNum) & ".</b> "
        
        strSearchResult = strSearchResult & "[<a class='LinkSearchResult' href='" & GetClassUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("ClassID"), rsSearch("ClassPurview")) & "' target='_blank'>" & rsSearch("ClassName") & "</a>] "
        
        strSearchResult = strSearchResult & "<a class='LinkSearchResult' href='" & GetPhotoUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("UpdateTime"), rsSearch("PhotoID"), rsSearch("ClassPurview"), rsSearch("InfoPurview"), rsSearch("InfoPoint")) & "' target='_blank'>"
        
        If strField = "PhotoName" Then
            strSearchResult = strSearchResult & "<b>" & Replace(ReplaceText(rsSearch("PhotoName"), 2) & "", "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "</b>"
        Else
            strSearchResult = strSearchResult & "<b>" & ReplaceText(rsSearch("PhotoName"), 2) & "</b>"
        End If
        strSearchResult = strSearchResult & "</a>"
        If strField = "Author" Then
            strSearchResult = strSearchResult & "&nbsp;[" & Replace(rsSearch("Author") & "", "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "]"
        Else
            strSearchResult = strSearchResult & "&nbsp;[" & rsSearch("Author") & "]"
        End If
        strSearchResult = strSearchResult & "[" & FormatDateTime(rsSearch("UpdateTime"), 1) & "][" & rsSearch("Hits") & "]"
        strSearchResult = strSearchResult & "<br>"
        
        Content = Left(Replace(Replace(ReplaceText(nohtml(rsSearch("PhotoIntro")), 1), ">", "&gt;"), "<", "&lt;"), SearchResult_ContentLenth)
        If strField = "Content" Then
            strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & Replace(Content, "" & Keyword & "", "<font color=red>" & Keyword & "</font>") & "����</div>"
        Else
            strSearchResult = strSearchResult & "<div style='padding:10px 20px'>" & Content & "����</div>"
        End If
        strSearchResult = strSearchResult & "<br>"
        PhotoNum = PhotoNum + 1
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
        strCustom = PE_Replace(strCustom, Match.value, GetSearchResultLabel(strParameter, Match.SubMatches(1), iChannelID))
    Next
    GetSearchResult2 = strCustom
End Function

Private Function GetSearchResultLabel(strTemp, strList, iChannelID)
    Dim sqlSearch, rsSearch, iCount, PhotoNum, arrPhotoID, Content
    Dim arrTemp
    Dim strPhotoPic, strPicTemp, arrPicTemp
    Dim arrClassID, IncludeChild, iSpecialID, ItemNum, IsHot, IsElite, Author, DateNum, OrderType, UsePage, TitleLen, ContentLen
    Dim iCols, iColsHtml, iRows, iRowsHtml, iNumber
    Dim rsField, ArrField, iField
    Dim rsCustom, strCustomList

    iCount = 0
    strCustomList = ""
    
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
    
    arrPhotoID = GetSearchResultIDArr(iChannelID)
    
    If arrPhotoID = "" Then
        GetSearchResultLabel = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "û�л�û���ҵ��κ�{$ChannelShortName}") & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If
       
    Set rsField = Conn.Execute("select FieldName,LabelName from PE_Field where ChannelID=-3 or ChannelID=" & ChannelID & "")
    If Not (rsField.BOF And rsField.EOF) Then
        ArrField = rsField.getrows(-1)
    End If
    Set rsField = Nothing
    
    sqlSearch = "select P.ChannelID,P.PhotoID,P.PhotoName,P.Author,P.UpdateTime,P.Hits,"
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlSearch = sqlSearch & "P." & ArrField(0, iField) & ","
        Next
    End If
    sqlSearch = sqlSearch & "P.InfoPurview,P.Keyword,P.InfoPoint,P.DayHits,P.WeekHits,P.MonthHits,P.PhotoThumb,P.OnTop,P.Elite,P.PhotoIntro,P.Editor,P.Inputer,P.CopyFrom,P.ChannelID,P.Stars,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,C.ReadMe from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID where PhotoID in (" & arrPhotoID & ") order by PhotoID desc"
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

        strTemp = PE_Replace(strTemp, "{$PhotoID}", rsCustom("PhotoID"))
        If InStr(strTemp, "{$PhotoUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoUrl}", GetPhotoUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("PhotoID"), rsCustom("ClassPurview"), rsCustom("InfoPurview"), rsCustom("InfoPoint")))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        strTemp = PE_Replace(strTemp, "{$Author}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$CopyFrom}", rsCustom("CopyFrom"))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$Inputer}", rsCustom("Inputer"))
        strTemp = PE_Replace(strTemp, "{$Editor}", rsCustom("Editor"))
        If InStr(strTemp, "{$InfoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$InfoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$PhotoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
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
            strTemp = PE_Replace(strTemp, "{$PhotoName}", GetSubStr(rsCustom("PhotoName"), TitleLen, ShowSuspensionPoints))
        Else
            strTemp = PE_Replace(strTemp, "{$PhotoName}", rsCustom("PhotoName"))
        End If
        strTemp = PE_Replace(strTemp, "{$PhotoNameOriginal}", rsCustom("PhotoName"))
        If ContentLen > 0 Then
            If InStr(strTemp, "{$PhotoIntro}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoIntro}", Left(nohtml(rsCustom("PhotoIntro")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$PhotoIntro}", "")
        End If
        If InStr(strTemp, "{$PhotoThumb}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoThumb}", GetPhotoThumb(rsCustom("PhotoThumb"), 130, 0))
        If InStr(strTemp, "{$DayHits}") > 0 Then strTemp = PE_Replace(strTemp, "{$DayHits}", GetHits(rsCustom("InfoPoint"), rsCustom("DayHits"), 1))
        If InStr(strTemp, "{$WeekHits}") > 0 Then strTemp = PE_Replace(strTemp, "{$WeekHits}", GetHits(rsCustom("InfoPoint"), rsCustom("WeekHits"), 2))
        If InStr(strTemp, "{$MonthHits}") > 0 Then strTemp = PE_Replace(strTemp, "{$MonthHits}", GetHits(rsCustom("InfoPoint"), rsCustom("MonthHits"), 3))
        
        '�滻ͼƬ����ͼ
        regEx.Pattern = "\{\$PhotoThumb\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strPhotoPic = GetPhotoThumb(Trim(rsCustom("PhotoThumb")), PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
            strTemp = Replace(strTemp, Match.value, strPhotoPic)
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
'��  �ã���ʾ���ͼƬ
'��  ����PhotoNum  ----�����ʾ���ٸ�ͼƬ
'        TitleLen   ----��������ַ�����һ������=����Ӣ���ַ�
'=================================================
Private Function GetCorrelative(PhotoNum, TitleLen)
    Dim rsCorrelative, sqlCorrelative, strCorrelative
    Dim strKey, arrKey, i, MaxNum
    If PhotoNum > 0 And PhotoNum <= 100 Then
        sqlCorrelative = "select top " & PhotoNum
    Else
        sqlCorrelative = "Select Top 5 "
    End If
    strKey = Mid(rsPhoto("Keyword"), 2, Len(rsPhoto("Keyword")) - 2)
    If InStr(strKey, "|") > 1 Then
        arrKey = Split(strKey, "|")
        MaxNum = UBound(arrKey)
        If MaxNum > 2 Then MaxNum = 2
        strKey = "((P.Keyword like '%|" & arrKey(0) & "|%')"
        For i = 1 To MaxNum
            strKey = strKey & " or (P.Keyword like '%|" & arrKey(i) & "|%')"
        Next
        strKey = strKey & ")"
    Else
        strKey = "(P.Keyword like '%|" & strKey & "|%')"
    End If
    sqlCorrelative = sqlCorrelative & " P.PhotoID,P.PhotoName,P.Author,P.UpdateTime,P.Hits,P.InfoPurview,P.InfoPoint,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Photo P inner join PE_Class C on P.ClassID=C.ClassID where P.ChannelID=" & ChannelID & " and P.Deleted=" & PE_False & " and P.Status=3"

    sqlCorrelative = sqlCorrelative & " and " & strKey & " and P.PhotoID<>" & PhotoID & " Order by P.PhotoID desc"
    Set rsCorrelative = Conn.Execute(sqlCorrelative)
    If TitleLen < 0 Or TitleLen > 255 Then TitleLen = 50
    If rsCorrelative.BOF And rsCorrelative.EOF Then
        strCorrelative = R_XmlText_Class("ShowPhoto/NoCorrelative", "û�����{$ChannelShortName}")
    Else
        Do While Not rsCorrelative.EOF
            strCorrelative = strCorrelative & "<li><a class='LinkPhotoCorrelative' href='" & GetPhotoUrl(rsCorrelative("ParentDir"), rsCorrelative("ClassDir"), rsCorrelative("UpdateTime"), rsCorrelative("PhotoID"), rsCorrelative("ClassPurview"), rsCorrelative("InfoPurview"), rsCorrelative("InfoPoint")) & "'"
            strCorrelative = strCorrelative & " title='" & Replace(Replace(Replace(Replace(strList_Title, "{$PhotoName}", rsCorrelative("PhotoName")), "{$Author}", rsCorrelative("Author")), "{$UpdateTime}", rsCorrelative("UpdateTime")), "{$br}", vbCrLf) & "'>" & GetSubStr(rsCorrelative("PhotoName"), TitleLen, ShowSuspensionPoints) & "</a></li>"
            rsCorrelative.MoveNext
        Loop
    End If
    rsCorrelative.Close
    Set rsCorrelative = Nothing
    GetCorrelative = strCorrelative
End Function

Private Function GetPrevPhoto()
    Dim rsPrev, sqlPrev
    sqlPrev = "Select Top 1 PhotoID,UpdateTime,InfoPurview,InfoPoint From PE_Photo Where ChannelID=" & ChannelID & " and Deleted=" & PE_False & " and Status=3 and ClassID=" & rsPhoto("ClassID") & " and PhotoID<" & rsPhoto("PhotoID") & " order by PhotoID DESC"
    Set rsPrev = Conn.Execute(sqlPrev)
    If rsPrev.BOF And rsPrev.EOF Then
        GetPrevPhoto = ""
    Else
        GetPrevPhoto = GetPhotoUrl(ParentDir, ClassDir, rsPrev("UpdateTime"), rsPrev("PhotoID"), ClassPurview, rsPrev("InfoPurview"), rsPrev("InfoPoint"))
    End If
    rsPrev.Close
    Set rsPrev = Nothing
End Function

Private Function GetNextPhoto()
    Dim rsNext, sqlNext
    sqlNext = "Select Top 1 PhotoID,UpdateTime,InfoPurview,InfoPoint From PE_Photo Where ChannelID=" & ChannelID & " and Deleted=" & PE_False & " and Status=3 and ClassID=" & rsPhoto("ClassID") & " and PhotoID>" & rsPhoto("PhotoID") & " order by PhotoID ASC"
    Set rsNext = Conn.Execute(sqlNext)
    If rsNext.BOF And rsNext.EOF Then
        GetNextPhoto = ""
    Else
        GetNextPhoto = GetPhotoUrl(ParentDir, ClassDir, rsNext("UpdateTime"), rsNext("PhotoID"), ClassPurview, rsNext("InfoPurview"), rsNext("InfoPoint"))
    End If
    rsNext.Close
    Set rsNext = Nothing
End Function

Private Function GetHits(iInfoPoint, iHits, HitsType)
    'HitsType 1�������������2�������������3�����������
    Dim strHits
    If UseCreateHTML > 0 And Not (ClassPurview > 0 Or iInfoPoint > 0) Then
        strHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?HitsType=" & HitsType & "&PhotoID=" & PhotoID & "'></script>"
    Else
        strHits = iHits
    End If
    GetHits = strHits
End Function

Private Function GetPhotoProperty()
    Dim strProperty
    If rsPhoto("OnTop") = True Then
        strProperty = strProperty & XmlText_Class("ShowPhoto/OnTop", "<font color=blue>��</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsPhoto("Hits") >= HitsOfHot Then
        strProperty = strProperty & XmlText_Class("ShowPhoto/Hot", "<font color=red>��</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsPhoto("Elite") = True Then
        strProperty = strProperty & XmlText_Class("ShowPhoto/Elite", "<font color=green>��</font>")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;"
    End If
    GetPhotoProperty = strProperty
End Function

Private Function GetStars(Stars)
    GetStars = "<font color='" & XmlText_Class("ShowPhoto/Star_Color", "#009900") & "'>" & String(Stars, XmlText_Class("ShowPhoto/Star", "��")) & "</font>"
End Function

Public Sub ReplaceViewPhoto()
    Dim FoundErr, ErrMsg, PurviewChecked, ConsumePoint
    FoundErr = False
    ErrMsg = ""

    If ClassPurview > 0 Or rsPhoto("InfoPurview") > 0 Or rsPhoto("InfoPoint") > 0 Then
        Dim ErrMsg_NoLogin, ErrMsg_PurviewCheckedErr, ErrMsg_PurviewCheckedErr2, ErrMsg_NoMail, ErrMsg_NoCheck, ErrMsg_NeedPoint, ErrMsg_UsePoint, ErrMsg_OutTime, ErrMsg_Overflow_Total, ErrMsg_Overflow_Today
        ErrMsg_NoLogin = Replace(Replace(Replace(R_XmlText_Class("PhotoContent/Nologin", "<br>&nbsp;&nbsp;&nbsp;&nbsp;�㻹ûע�᣿����û�е�¼����{$ItemUnit}Ҫ�������Ǳ�վ��ע���Ա�����Ķ���<br><br>&nbsp;&nbsp;&nbsp;&nbsp;����㻹ûע�ᣬ��Ͻ�<a href='{$InstallDir}Reg/User_Reg.asp'><font color=red>���ע��</font></a>�ɣ�<br><br>&nbsp;&nbsp;&nbsp;&nbsp;������Ѿ�ע�ᵫ��û��¼����Ͻ�<a href='{$InstallDir}User/User_Login.asp'><font color=red>��˵�¼</font></a>�ɣ�<br><br>"), "{$ItemUnit}", ChannelItemUnit & ChannelShortName), "{$ChannelItemUnit}", ChannelItemUnit), "{$InstallDir}", strInstallDir)
        If UserLogined <> True Then
            FoundErr = True
            ErrMsg = ErrMsg & ErrMsg_NoLogin
        Else
            Call GetUser(UserName)
            ErrMsg_PurviewCheckedErr = XmlText("BaseText", "PurviewCheckedErr", "<li>�Բ�����û�в鿴����Ŀ���ݵ�Ȩ�ޣ�</li>")
            ErrMsg_PurviewCheckedErr2 = XmlText("BaseText", "PurviewCheckedErr2", "<li>�Բ�����û�в鿴����Ϣ��Ȩ�ޣ�</li>")
            ErrMsg_NoMail = "<li>" & R_XmlText_Class("PhotoContent/NoMail", "�Բ�������δͨ���ʼ���֤�����ܲ鿴��{$ChannelShortName}") & "</li>"
            ErrMsg_NoCheck = "<li>" & R_XmlText_Class("PhotoContent/NoCheck", "�Բ�������δͨ������Ա��ˣ����ܲ鿴�շ�{$ChannelShortName}") & "</li>"
            ErrMsg_NeedPoint = Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("PhotoContent/NeedPoint", "<p align='center'><br><br>�Բ��𣬲鿴��ͼƬ��Ҫ���� <b><font color=red>{$NeedPoint}</font></b> {$PointUnit}{$PointName}������Ŀǰֻ�� <b><font color=blue>{$NowPoint}</font></b> {$PointUnit}{$PointName}���á�{$PointName}�����㣬�޷��鿴��ͼƬ������������ϵ���г�ֵ��</p>"), "{$InfoPoint}", rsPhoto("InfoPoint")), "{$NeedPoint}", rsPhoto("InfoPoint")), "{$NowPoint}", UserPoint), "{$PointName}", PointName), "{$PointUnit}", PointUnit)
            ErrMsg_UsePoint = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("PhotoContent/UsePoint", "<p align='center'><br><br>�鿴��ͼƬ��Ҫ���� <b><font color=red>{$InfoPoint}</font></b> {$PointUnit}{$PointName}����Ŀǰ���� <b><font color=blue>{$NowPoint}</font></b> {$PointUnit}{$PointName}���á��鿴��ͼƬ���㽫ʣ�� <b><font color=green>{$FinalPoint}</font></b> {$PointUnit}{$PointName}<br><br>��ȷʵԸ�⻨�� <b><font color=red>{$InfoPoint}</font></b> {$PointUnit}{$PointName}���鿴��ͼƬ��<br><br><a href='{$FileName}?Pay=yes&PhotoID={$PhotoID}'>��Ը��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='{$InstallDir}index.asp'>�Ҳ�Ը��</a></p>"), "{$InfoPoint}", rsPhoto("InfoPoint")), "{$NowPoint}", UserPoint), "{$FinalPoint}", UserPoint - rsPhoto("InfoPoint")), "{$FileName}", strFileName), "{$PhotoID}", PhotoID), "{$InstallDir}", strInstallDir), "{$PointName}", PointName), "{$PointUnit}", PointUnit)
            ErrMsg_OutTime = R_XmlText_Class("PhotoContent/OutTime", "<p align='center'><br><br><font color=red>�Բ��𣬱�ͼƬΪ�շ����ݣ���������Ч���Ѿ����ڣ������޷��鿴��ͼƬ������������ϵ���г�ֵ��</font></p>")
            ErrMsg_Overflow_Total = "<li>" & R_XmlText_Class("PhotoContent/Overflow_Total", "���Ѿ��ﵽ�򳬹���Ч�������ܲ鿴����Ϣ������") & "</li>"
            ErrMsg_Overflow_Today = "<li>" & R_XmlText_Class("PhotoContent/Overflow_Today", "���Ѿ��ﵽ�򳬹��������ܲ鿴����Ϣ������") & "</li>"
            Select Case rsPhoto("InfoPurview")
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
                PurviewChecked = FoundInArr(rsPhoto("arrGroupID"), GroupID, ",")
                If PurviewChecked = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & ErrMsg_PurviewCheckedErr2
                End If
            End Select
            If PurviewChecked = True Then
                If rsPhoto("InfoPoint") > 0 And rsPhoto("InfoPoint") < 9999 Then
                    If GroupType < 1 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_NoMail
                    ElseIf GroupType = 1 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & ErrMsg_NoCheck
                    Else
                        Dim trs, ValidConsumeLogID, DividePoint

                        If UserChargeType = 0 Then   '��������
                            ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, PhotoID, rsPhoto("ChargeType"), rsPhoto("PitchTime"), rsPhoto("ReadTimes"))
                            If ValidConsumeLogID = 0 Then   '���û���ҵ���¼���ѣ���Ҫ��ʼ�Ʒ�
                                If UserPoint < rsPhoto("InfoPoint") Then  '����û��ĵ���С��Ҫ�۵ĵ���
                                    FoundErr = True
                                    ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                Else
                                    If LCase(Trim(Request("Pay"))) = "yes" Then  '����û�ȷ��Ҫ�۵�
                                        Conn.Execute "update PE_User set UserPoint=UserPoint-" & rsPhoto("InfoPoint") & " where UserName='" & UserName & "'"
                                        Call AddConsumeLog("System", ModuleType, UserName, PhotoID, rsPhoto("InfoPoint"), 2, "���ڲ鿴�շ�" & ChannelShortName & "��" & rsPhoto("PhotoName"))
                                        If rsPhoto("DividePercent") <= 0 Then
                                            DividePoint = 0
                                        ElseIf rsPhoto("DividePercent") > 0 And rsPhoto("DividePercent") <= 100 Then
                                            DividePoint = PE_CLng(rsPhoto("InfoPoint") * rsPhoto("DividePercent") / 100)
                                        Else
                                            DividePoint = rsPhoto("InfoPoint")
                                        End If
                                        If DividePoint > 0 Then
                                            Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsPhoto("Inputer") & "'"
                                            Call AddConsumeLog("System", ModuleType, rsPhoto("Inputer"), 0, DividePoint, 1, "�ӡ�" & rsPhoto("PhotoName") & "�����շ��зֳ�")
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
                                    ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, PhotoID, rsPhoto("ChargeType"), rsPhoto("PitchTime"), rsPhoto("ReadTimes"))
                                    If ValidConsumeLogID = 0 Then   '���û���ҵ���¼���ѣ���Ҫ��ʼ�Ʒ�
                                        If UserPoint < rsPhoto("InfoPoint") Then  '����û��ĵ���С��Ҫ�۵ĵ���
                                            FoundErr = True
                                            ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                        Else
                                            If LCase(Trim(Request("Pay"))) = "yes" Then  '����û�ȷ��Ҫ�۵�
                                                Conn.Execute "update PE_User set UserPoint=UserPoint-" & rsPhoto("InfoPoint") & " where UserName='" & UserName & "'"
                                                Call AddConsumeLog("System", ModuleType, UserName, PhotoID, rsPhoto("InfoPoint"), 2, "���ڲ鿴�շ�" & ChannelShortName & "��" & rsPhoto("PhotoName"))
                                                If rsPhoto("DividePercent") <= 0 Then
                                                    DividePoint = 0
                                                ElseIf rsPhoto("DividePercent") > 0 And rsPhoto("DividePercent") <= 100 Then
                                                    DividePoint = PE_CLng(rsPhoto("InfoPoint") * rsPhoto("DividePercent") / 100)
                                                Else
                                                    DividePoint = rsPhoto("InfoPoint")
                                                End If
                                                If DividePoint > 0 Then
                                                    Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsPhoto("Inputer") & "'"
                                                    Call AddConsumeLog("System", ModuleType, rsPhoto("Inputer"), 0, DividePoint, 1, "�ӡ�" & rsPhoto("PhotoName") & "�����շ��зֳ�")
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
                                    ValidConsumeLogID = GetValidConsumeLogID(UserName, ModuleType, PhotoID, rsPhoto("ChargeType"), rsPhoto("PitchTime"), rsPhoto("ReadTimes"))
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
                                                Call AddConsumeLog("System", ModuleType, UserName, PhotoID, 0, 2, "��Ч���ڲ鿴�շ�" & ChannelShortName & "��" & rsPhoto("PhotoName") & "��Ӧ�۵�����" & rsPhoto("InfoPoint") & "")
                                            Else  '�۵�
                                                If UserPoint >= rsPhoto("InfoPoint") Then   '��������㹻
                                                    '�����Ŀ۷���ʾ
                                                    If LCase(Trim(Request("Pay"))) = "yes" Then  '����û�ȷ��Ҫ�۵�
                                                        Conn.Execute "update PE_User set UserPoint=UserPoint-" & rsPhoto("InfoPoint") & " where UserName='" & UserName & "'"
                                                        Call AddConsumeLog("System", ModuleType, UserName, PhotoID, rsPhoto("InfoPoint"), 2, "��Ч���ڲ鿴�շ�" & ChannelShortName & "��" & rsPhoto("PhotoName"))
                                                        If rsPhoto("DividePercent") <= 0 Then
                                                            DividePoint = 0
                                                        ElseIf rsPhoto("DividePercent") > 0 And rsPhoto("DividePercent") <= 100 Then
                                                            DividePoint = PE_CLng(rsPhoto("InfoPoint") * rsPhoto("DividePercent") / 100)
                                                        Else
                                                            DividePoint = rsPhoto("InfoPoint")
                                                        End If
                                                        If DividePoint > 0 Then
                                                            Conn.Execute "update PE_User set UserPoint=UserPoint+" & DividePoint & " where UserName='" & rsPhoto("Inputer") & "'"
                                                            Call AddConsumeLog("System", ModuleType, rsPhoto("Inputer"), 0, DividePoint, 1, "�ӡ�" & rsPhoto("PhotoName") & "�����շ��зֳ�")
                                                        End If
                                                    Else    '���û�û��ȷ��ǰ���Ƚ��п۷���ʾ
                                                        FoundErr = True
                                                        ErrMsg = ErrMsg & ErrMsg_UsePoint
                                                    End If
                                                Else   '����������ʱ
                                                    If UserChargeType = 2 Then '�����������Ч�ڵ��ں󣬾Ͳ��ɲ鿴�շ����ݡ�
                                                        FoundErr = True
                                                        ErrMsg = ErrMsg & ErrMsg_NeedPoint
                                                    Else   '��Ч�����Ȼ���Ч�ڹ��ں͵������꣬�Ų��ɲ鿴�շ����ݣ���ʱֻ���¼
                                                        Call AddConsumeLog("System", ModuleType, UserName, PhotoID, 0, 2, "��Ч���ڲ鿴�շ�" & ChannelShortName & "��" & rsPhoto("PhotoName") & "��Ӧ�۵�����" & rsPhoto("InfoPoint") & "")
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
        strHtml = Replace(strHtml, "{$ViewPhoto}", ErrMsg)
        strHtml = Replace(strHtml, "{$PhotoUrlList}", "")
        strHtml = Replace(strHtml, "{$PhotoUrl}", "")
        strHtml = Replace(strHtml, "{$GetUrlArray}", "")
        regEx.Pattern = "\{\$ViewPhoto\((.*?)\)\}"
        strHtml = regEx.Replace(strHtml, ErrMsg)
        regEx.Pattern = "\{\$PhotoUrlList\((.*?)\)\}"
        strHtml = regEx.Replace(strHtml, "")
    Else
        Call ReplacePhotoContent
    End If
End Sub

Private Sub ReplacePhotoContent()
    strHtml = Replace(strHtml, "{$ViewPhoto}", ViewPhoto(600, 0, False))
    'strHTML = Replace(strHTML, "{$PhotoUrlList}", GetPhotoUrlList(0,0,0,6,0))
    strHtml = Replace(strHtml, "{$PhotoUrlList}", GetPhotoUrlList(1, 130, 0, 5, 5))
    strHtml = Replace(strHtml, "{$PhotoUrl}", GetFirstPhotoUrl())
    strHtml = Replace(strHtml, "{$GetUrlArray}", GetUrlArray())
    
    Dim arrTemp, strViewPhoto, strUrlList
    regEx.Pattern = "\{\$ViewPhoto\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        Select Case UBound(arrTemp)
        Case 0
            strViewPhoto = ViewPhoto(PE_CLng(arrTemp(0)), 0 ,False)
        Case 1
            strViewPhoto = ViewPhoto(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), False)
        Case 2
            strViewPhoto = ViewPhoto(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CBool(arrTemp(2)))			
        End Select
        strHtml = Replace(strHtml, Match.value, strViewPhoto)
    Next

    regEx.Pattern = "\{\$PhotoUrlList\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 4 Then
            strUrlList = "����ʽ��ǩ��{$GetPhotoUrlList(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Else
            strUrlList = GetPhotoUrlList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)))
        End If
        strHtml = Replace(strHtml, Match.value, strUrlList)
    Next
    
End Sub

Private Function ViewPhoto(ImgWidth, ImgHeight, ShowPicInfo)
    Dim PhotoUrl, ImgSetting, strViewPhoto
    PhotoUrl = GetViewFirstPhotoUrl()
    If PhotoUrl = "" Then
        PhotoUrl = "images/nopic.gif"
    End If
    If ImgWidth > 0 Then
        ImgSetting = " onload='if(this.width>" & ImgWidth & ") this.width=" & ImgWidth & "'"
    Else
        ImgWidth = 550
    End If
    If ImgHeight <= 0 Then
        ImgHeight = 400
    End If
    ShowPicInfo = PE_CBool(ShowPicInfo)
	
    strViewPhoto = strViewPhoto & "<div id='imgBox'></div>" & vbCrLf
    strViewPhoto = strViewPhoto & "<script language='javascript'>" & vbCrLf
    strViewPhoto = strViewPhoto & "function ViewPhoto(PhotoUrl,PhotoDesc){" & vbCrLf
    strViewPhoto = strViewPhoto & "  var strHtml;" & vbCrLf
    strViewPhoto = strViewPhoto & "  var FileExt=PhotoUrl.substr(PhotoUrl.lastIndexOf('.')+1).toLowerCase();" & vbCrLf
    strViewPhoto = strViewPhoto & "  if(FileExt=='gif'||FileExt=='jpg'||FileExt=='png'||FileExt=='bmp'||FileExt=='jpeg'){" & vbCrLf
    strViewPhoto = strViewPhoto & "    strHtml=""<a href='""+PhotoUrl+""' target='PhotoView'><img alt='""+PhotoDesc+""'  src='""+PhotoUrl+""'  border='0'" & ImgSetting & " ></a>"
    If ShowPicInfo = True Then strViewPhoto = strViewPhoto & "<br><a href='""+PhotoUrl+""' target='PhotoView' classs='imgBoxInfo'>""+PhotoDesc+""</a>" 
    strViewPhoto = strViewPhoto &""";"& vbCrLf
    strViewPhoto = strViewPhoto & "  }else if(FileExt=='swf'){" & vbCrLf
    strViewPhoto = strViewPhoto & "    strHtml=""<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' width='" & ImgWidth & "' height='" & ImgHeight & "' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'><param name='movie' value='""+PhotoUrl+""'><param name='quality' value='high'><embed src='""+PhotoUrl+""' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='550' height='400'></embed></object>"";" & vbCrLf
    strViewPhoto = strViewPhoto & "  }else{" & vbCrLf
    strViewPhoto = strViewPhoto & "    strHtml=PhotoUrl;" & vbCrLf
    strViewPhoto = strViewPhoto & "  }" & vbCrLf
    strViewPhoto = strViewPhoto & "  imgBox.innerHTML=strHtml;" & vbCrLf
    strViewPhoto = strViewPhoto & "}" & vbCrLf
    strViewPhoto = strViewPhoto & "ViewPhoto('" & PhotoUrl & "');" & vbCrLf
    strViewPhoto = strViewPhoto & "</script>" & vbCrLf
    ViewPhoto = strViewPhoto
End Function

Private Function GetViewFirstPhotoUrl()
    Dim arrPhotoUrls, arrUrls
    If InStr(rsPhoto("PhotoUrl"), "$$$") > 0 Then
        arrPhotoUrls = Split(rsPhoto("PhotoUrl"), "$$$")
        arrUrls = Split(arrPhotoUrls(0), "|")
    Else
        arrUrls = Split(rsPhoto("PhotoUrl"), "|")
    End If
    If UBound(arrUrls) <> 1 Then
        GetViewFirstPhotoUrl = ""
    Else
        If arrUrls(1) = "" Or LCase(arrUrls(1)) = "http://" Then
            GetViewFirstPhotoUrl = ""
        Else
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                GetViewFirstPhotoUrl = ChannelUrl & "/" & UploadDir & "/" & arrUrls(1)&"','"&arrUrls(0)
            Else
                GetViewFirstPhotoUrl = arrUrls(1)&"','"&arrUrls(0)
            End If
        End If
    End If
End Function
Private Function GetFirstPhotoUrl()
    Dim arrPhotoUrls, arrUrls
    If InStr(rsPhoto("PhotoUrl"), "$$$") > 0 Then
        arrPhotoUrls = Split(rsPhoto("PhotoUrl"), "$$$")
        arrUrls = Split(arrPhotoUrls(0), "|")
    Else
        arrUrls = Split(rsPhoto("PhotoUrl"), "|")
    End If
    If UBound(arrUrls) <> 1 Then
        GetFirstPhotoUrl = ""
    Else
        If arrUrls(1) = "" Or LCase(arrUrls(1)) = "http://" Then
            GetFirstPhotoUrl = ""
        Else
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                GetFirstPhotoUrl = ChannelUrl & "/" & UploadDir & "/" & arrUrls(1)
            Else
                GetFirstPhotoUrl = arrUrls(1)
            End If
        End If
    End If
End Function


Private Function GetUrlArray()
    Dim strArray, arrPhotoUrls, iTemp, arrUrls, PhotoUrl
    strArray = "<script language='javascript'>" & vbCrLf
    strArray = strArray & "var arrUrlName=new Array();" & vbCrLf
    strArray = strArray & "var arrUrl=new Array();" & vbCrLf
    arrPhotoUrls = Split(rsPhoto("PhotoUrl"), "$$$")
    For iTemp = 0 To UBound(arrPhotoUrls)
        arrUrls = Split(arrPhotoUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If arrUrls(1) <> "" And arrUrls(1) <> "http://" Then
                If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then  '�����ļ�
                    PhotoUrl = ChannelUrl & "/" & UploadDir & "/" & arrUrls(1)
                Else
                    If FoundInArr("gif,jpg,jpeg,jpe,bmp,png,swf", LCase(Mid(arrUrls(1), InStrRev(arrUrls(1), ".") + 1)), ",") = True Then
                        PhotoUrl = arrUrls(1)
                    End If
                End If
                strArray = strArray & "arrUrlName[" & iTemp & "]='" & arrUrls(0) & "';" & vbCrLf
                strArray = strArray & "arrUrl[" & iTemp & "]='" & PhotoUrl & "';" & vbCrLf
            End If
        End If
    Next
    strArray = strArray & "</script>" & vbCrLf
    GetUrlArray = strArray
End Function

Private Function GetPhotoUrlList(ShowType, ImgWidth, ImgHeight, Cols, MaxPerPage)
    If rsPhoto("PhotoUrl") & "" = "" Then
        GetPhotoUrlList = ""
        Exit Function
    End If
    If Cols < 1 Then Cols = 1
    If MaxPerPage < 1 Then MaxPerPage = 1
    Dim strUrls, ImgSetting
    strUrls = GetUrlArray()
    strUrls = strUrls & "<div id='PhotoUrlList'></div>"
    strUrls = strUrls & "<script language='javascript'>" & vbCrLf
    If ShowType = 0 Then   '����
        strUrls = strUrls & "for(var i=0;i<arrUrl.length;i++){" & vbCrLf
        strUrls = strUrls & "  document.write(""<a href='#Title' onclick=ViewPhoto('""+arrUrl[i]+""')>""+arrUrlName[i]+""</a>&nbsp;&nbsp;"");" & vbCrLf
        strUrls = strUrls & "  if((i+1)%" & Cols & "==0&&i+1<arrUrl.length){document.write('<br>');}" & vbCrLf
        strUrls = strUrls & "}" & vbCrLf
    Else    'ͼƬ
        If ImgWidth > 0 Then
            ImgSetting = " width='" & ImgWidth & "'"
        End If
        If ImgHeight > 0 Then
            ImgSetting = ImgSetting & " height='" & ImgHeight & "'"
        End If

        strUrls = strUrls & "function ShowUrlList(page){" & vbCrLf
        strUrls = strUrls & "  if(arrUrl.length<=1) return '';" & vbCrLf
        strUrls = strUrls & "  var dTotalPage=arrUrl.length/" & MaxPerPage & ";" & vbCrLf
        strUrls = strUrls & "  var TotalPage;" & vbCrLf
        strUrls = strUrls & "  var MaxPerPage=" & MaxPerPage & ";" & vbCrLf
        strUrls = strUrls & "  if(arrUrl.length%MaxPerPage==0){TotalPage=Math.floor(dTotalPage);}else{TotalPage=Math.floor(dTotalPage)+1;}" & vbCrLf

        strUrls = strUrls & "  if(page<1) page=1;" & vbCrLf
        strUrls = strUrls & "  if(page>TotalPage) page=TotalPage;" & vbCrLf
        strUrls = strUrls & "  var strPage='<table><tr>';" & vbCrLf
        strUrls = strUrls & "  for(var i=(page-1)*MaxPerPage;i<arrUrl.length&&i<page*MaxPerPage;i++){" & vbCrLf
        strUrls = strUrls & "    strPage+=""<td><a href='#Title' onclick=ViewPhoto('""+arrUrl[i]+""','""+arrUrlName[i]+""')><img src='""+arrUrl[i]+""' border='0' " & ImgSetting & "></a></td>"";" & vbCrLf
        strUrls = strUrls & "    if((i+1)%" & Cols & "==0&&i+1<page*MaxPerPage){strPage+='</tr><tr>';}" & vbCrLf
        strUrls = strUrls & "  }" & vbCrLf
        strUrls = strUrls & "  strPage+=""</tr></table>"";" & vbCrLf
        strUrls = strUrls & "  if(TotalPage>1){strPage+=""<table><tr><td><a href='javascript:ShowUrlList(1)'>��ҳ</a> <a href='javascript:ShowUrlList(""+(page-1)+"")'>��һҳ</a> <a href='javascript:ShowUrlList(""+(page+1)+"")'>��һҳ</a> <a href='javascript:ShowUrlList(""+TotalPage+"")'>βҳ</a></td></tr></table>"";}" & vbCrLf
        'strUrls = strUrls & "  alert(strPage);" & vbcrlf
        strUrls = strUrls & "  PhotoUrlList.innerHTML=strPage;" & vbCrLf
        strUrls = strUrls & "}" & vbCrLf
        strUrls = strUrls & "ShowUrlList(1);" & vbCrLf
    End If
    strUrls = strUrls & "</script>" & vbCrLf
    GetPhotoUrlList = strUrls
End Function


Public Function GetCustomFromTemplate(strValue)   '�õ��Զ����б�İ�����Ƶ�HTML����
    Dim strCustom, strParameter
	strCustom = strValue
    regEx.Pattern = "��PhotoList\((.*?)\)��([\s\S]*?)��\/PhotoList��"
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
    regEx.Pattern = "\{\$GetPhotoList\((.*?)\)\}"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        strList = PE_Replace(strList, Match.value, GetListFromLabel(Match.SubMatches(0)))
    Next
    GetListFromTemplate = strList
End Function

Public Function GetPicFromTemplate(ByVal strValue)
    Dim strPicList
    strPicList = strValue
    regEx.Pattern = "\{\$GetPicPhoto\((.*?)\)\}"
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
    regEx.Pattern = "\{\$GetSlidePicPhoto\((.*?)\)\}"
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
        GetSlidePicFromLabel = GetSlidePicPhoto(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), -1)
    Case 13
        GetSlidePicFromLabel = GetSlidePicPhoto(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)))
    Case Else
        GetSlidePicFromLabel = "����ʽ��ǩ��{$GetSlidePicPhoto(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
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
        GetPicFromLabel = "����ʽ��ǩ��{$GetPicPhoto(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
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
    
    GetPicFromLabel = GetPicPhoto(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CBool(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CLng(arrTemp(16)))
End Function

Private Function GetListFromLabel(ByVal strSource)
    Dim arrTemp
    Dim tChannelID, PhotoNum, arrClassID, tSpecialID, OrderType, OpenType
    If strSource = "" Then
        GetListFromLabel = ""
        Exit Function
    End If
    
    strSource = Replace(strSource, Chr(34), "")
    strSource = FillInArrStr(strSource, "1,listA,listbg,listbg2", 28)
    arrTemp = Split(strSource, ",")
    If UBound(arrTemp) + 1 < 28 Then
        GetListFromLabel = "����ʽ��ǩ��{$GetPhotoList(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
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
        PhotoNum = 8
    Case "TopNumber"
        PhotoNum = 8
    Case Else
        PhotoNum = PE_CLng(arrTemp(5))
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

    GetListFromLabel = GetPhotoList(tChannelID, arrClassID, PE_CBool(arrTemp(2)), tSpecialID, PE_CLng(arrTemp(4)), PhotoNum, PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), arrTemp(8), PE_CLng(arrTemp(9)), OrderType, PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CBool(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CLng(arrTemp(17)), PE_CBool(arrTemp(18)), PE_CBool(arrTemp(19)), PE_CBool(arrTemp(20)), PE_CBool(arrTemp(21)), PE_CBool(arrTemp(22)), OpenType, PE_CLng(arrTemp(24)), Trim(arrTemp(25)), Trim(arrTemp(26)), Trim(arrTemp(27)))

End Function

Private Function GetCustomFromLabel(strTemp, strList)
    Dim arrTemp
    Dim strPhotoPic, strPicTemp, arrPicTemp
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
        GetCustomFromLabel = "�Զ����б��ǩ����PhotoList(�����б�)���б����ݡ�/PhotoList���Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
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
    Set rsField = Conn.Execute("select FieldName,LabelName,FieldType from PE_Field where ChannelID=-3 or ChannelID=" & ChannelID & "")
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
        sqlCustom = sqlCustom & "P.PhotoIntro,"
    End If
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlCustom = sqlCustom & "P." & ArrField(0, iField) & ","
        Next
    End If
    sqlCustom = sqlCustom & "P.PhotoID,P.ChannelID,P.ClassID,P.PhotoName,P.Keyword,P.PhotoThumb,P.DayHits,P.WeekHits,P.MonthHits"
    sqlCustom = sqlCustom & ",P.Author,P.CopyFrom,P.Inputer,P.Editor,P.UpdateTime,P.Stars,P.Hits,P.OnTop,P.Elite,P.InfoPoint,P.InfoPurview"
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

        strTemp = PE_Replace(strTemp, "{$PhotoID}", rsCustom("PhotoID"))
        If InStr(strTemp, "{$PhotoUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoUrl}", GetPhotoUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("PhotoID"), rsCustom("ClassPurview"), rsCustom("InfoPurview"), rsCustom("InfoPoint")))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        strTemp = PE_Replace(strTemp, "{$Author}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$CopyFrom}", rsCustom("CopyFrom"))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$Inputer}", rsCustom("Inputer"))
        strTemp = PE_Replace(strTemp, "{$Editor}", rsCustom("Editor"))
        If InStr(strTemp, "{$InfoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$InfoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
        If InStr(strTemp, "{$PhotoPoint}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoPoint}", GetInfoPoint(rsCustom("InfoPoint")))
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
            strTemp = PE_Replace(strTemp, "{$PhotoName}", GetSubStr(rsCustom("PhotoName"), TitleLen, ShowSuspensionPoints))
        Else
            strTemp = PE_Replace(strTemp, "{$PhotoName}", rsCustom("PhotoName"))
        End If
        strTemp = PE_Replace(strTemp, "{$PhotoNameOriginal}", rsCustom("PhotoName"))
        If ContentLen > 0 Then
            If InStr(strTemp, "{$PhotoIntro}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoIntro}", Left(nohtml(rsCustom("PhotoIntro")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$PhotoIntro}", "")
        End If
        If InStr(strTemp, "{$PhotoThumb}") > 0 Then strTemp = PE_Replace(strTemp, "{$PhotoThumb}", GetPhotoThumb(rsCustom("PhotoThumb"), 130, 0))
        If InStr(strTemp, "{$DayHits}") > 0 Then strTemp = PE_Replace(strTemp, "{$DayHits}", GetHits(rsCustom("InfoPoint"), rsCustom("DayHits"), 1))
        If InStr(strTemp, "{$WeekHits}") > 0 Then strTemp = PE_Replace(strTemp, "{$WeekHits}", GetHits(rsCustom("InfoPoint"), rsCustom("WeekHits"), 2))
        If InStr(strTemp, "{$MonthHits}") > 0 Then strTemp = PE_Replace(strTemp, "{$MonthHits}", GetHits(rsCustom("InfoPoint"), rsCustom("MonthHits"), 3))
        
        '�滻ͼƬ����ͼ
        regEx.Pattern = "\{\$PhotoThumb\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strPhotoPic = GetPhotoThumb(Trim(rsCustom("PhotoThumb")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)))
            strTemp = Replace(strTemp, Match.value, strPhotoPic)
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





Public Sub GetHtml_Index()
    Dim strTemp, arrTemp, iCols, iClassID
    Dim PhotoList_ChildClass, PhotoList_ChildClass2

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
    regEx.Pattern = "��PhotoList_ChildClass��([\s\S]*?)��\/PhotoList_ChildClass��"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        PhotoList_ChildClass = Match.SubMatches(0)
        strHtml = regEx.Replace(strHtml, "{$PhotoList_ChildClass}")
        
        '�õ�ÿ����ʾ������
        iCols = 1
        regEx.Pattern = "��Cols=(\d{1,2})��"
        Set Matches2 = regEx.Execute(PhotoList_ChildClass)
        PhotoList_ChildClass = regEx.Replace(PhotoList_ChildClass, "")
        For Each Match2 In Matches2
            If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
        Next
     
        '��ʼѭ�����õ���������Ŀ�б��HTML����
        PhotoList_ChildClass2 = ""
        iClassID = 0
        Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=0 and ShowOnIndex=" & PE_True & " order by RootID")
        Do While Not rsClass.EOF
            strTemp = PhotoList_ChildClass
            
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
                PhotoList_ChildClass2 = PhotoList_ChildClass2 & strTemp
                If iCols > 1 Then PhotoList_ChildClass2 = PhotoList_ChildClass2 & "</tr><tr>"
            Else
                PhotoList_ChildClass2 = PhotoList_ChildClass2 & strTemp
                If iCols > 1 Then PhotoList_ChildClass2 = PhotoList_ChildClass2 & "<td width='1'></td>"
            End If
        Loop
        rsClass.Close
        Set rsClass = Nothing

        strHtml = Replace(strHtml, "{$PhotoList_ChildClass}", PhotoList_ChildClass2)
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
    If Child > 0 Then    '�����ǰ��Ŀ������Ŀ
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(0, 0, 3, 3, 0, True))
    Else
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(ParentID, 0, 3, 3, 0, True))
    End If
    
    Dim PhotoList_CurrentClass, PhotoList_CurrentClass2, PhotoList_ChildClass, PhotoList_ChildClass2
    If Child > 0 And ClassShowType <> 2 Then    '�����ǰ��Ŀ������Ŀ
        ItemCount = PE_CLng(Conn.Execute("select Count(*) from PE_Photo where ClassID=" & ClassID & "")(0))
        If ItemCount <= 0 Then     '�����ǰ��Ŀû������
            regEx.Pattern = "��PhotoList_CurrentClass��([\s\S]*?)��\/PhotoList_CurrentClass��"
            strHtml = regEx.Replace(strHtml, "") '��ȥ����ʾ��ǰ��Ŀ��ֻ���ڱ���Ŀ�������б�
        Else      '�����ǰ��Ŀ������Ŀ���ҵ�ǰ��Ŀ�����ݣ�����Ҫ��ʾ������
            regEx.Pattern = "��PhotoList_CurrentClass��([\s\S]*?)��\/PhotoList_CurrentClass��"
            Set Matches = regEx.Execute(strHtml)
            For Each Match In Matches
                PhotoList_CurrentClass = Match.SubMatches(0)
                strHtml = regEx.Replace(strHtml, "{$PhotoList_CurrentClass}")
                
                strTemp = PhotoList_CurrentClass
                strTemp = GetCustomFromTemplate(strTemp)
                strTemp = GetListFromTemplate(strTemp)
                strTemp = GetPicFromTemplate(strTemp)
                
                strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
                strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", ReadMe)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", ClassName)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", ClassID)
                
                strHtml = Replace(strHtml, "{$PhotoList_CurrentClass}", strTemp)
            Next
        End If
        
        '�õ�����Ŀ�б�İ�����Ƶ�HTML����
        regEx.Pattern = "��PhotoList_ChildClass��([\s\S]*?)��\/PhotoList_ChildClass��"
        Set Matches = regEx.Execute(strHtml)
        For Each Match In Matches
            PhotoList_ChildClass = Match.SubMatches(0)
            strHtml = regEx.Replace(strHtml, "{$PhotoList_ChildClass}")
            
            '�õ�ÿ����ʾ������
            iCols = 1
            regEx.Pattern = "��Cols=(\d{1,2})��"
            Set Matches2 = regEx.Execute(PhotoList_ChildClass)
            PhotoList_ChildClass = regEx.Replace(PhotoList_ChildClass, "")
            For Each Match2 In Matches2
                If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
            Next
            
            '��ʼѭ�����õ���������Ŀ�б��HTML����
            iClassID = 0
            Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=" & ClassID & " and IsElite=" & PE_True & " and ClassType=1 order by RootID,OrderID")
            Do While Not rsClass.EOF
                strTemp = PhotoList_ChildClass
                
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
                    PhotoList_ChildClass2 = PhotoList_ChildClass2 & strTemp
                    If iCols > 1 Then PhotoList_ChildClass2 = PhotoList_ChildClass2 & "</tr><tr>"
                Else
                    PhotoList_ChildClass2 = PhotoList_ChildClass2 & strTemp
                    If iCols > 1 Then PhotoList_ChildClass2 = PhotoList_ChildClass2 & "<td width='1'></td>"
                End If
            Loop
            rsClass.Close
            Set rsClass = Nothing

            strHtml = Replace(strHtml, "{$PhotoList_ChildClass}", PhotoList_ChildClass2)
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


Public Sub GetHtml_Photo()
    strHtml = GetCustomFromTemplate(strHtml)  '�����Ƚ����Զ����б��ǩ

    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If

    strHtml = PE_Replace(strHtml, "{$PhotoID}", PhotoID)
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    
    Call ReplaceCommonLabel   '����ͨ�ñ�ǩ�������Զ����ǩ
    
    strHtml = GetCustomFromTemplate(strHtml)  '�����Ƚ����Զ����б��ǩ
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    strHtml = PE_Replace(strHtml, "{$PhotoID}", PhotoID)
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = Replace(strHtml, "{$PageTitle}", ReplaceText(PhotoName, 2))
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    
    If InStr(strHtml, "{$MY_") > 0 Then
        Dim rsField
        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-3 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            strHtml = PE_Replace(strHtml, rsField("LabelName"), PE_HTMLEncode(rsPhoto(Trim(rsField("FieldName")))))
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
    
    If InStr(strHtml, "{$Author}") > 0 Then strHtml = PE_Replace(strHtml, "{$Author}", GetAuthorInfo(rsPhoto("Author"), ChannelID))
    If InStr(strHtml, "{$CopyFrom}") > 0 Then strHtml = PE_Replace(strHtml, "{$CopyFrom}", GetCopyFromInfo(rsPhoto("CopyFrom"), ChannelID))
    If InStr(strHtml, "{$Hits}") > 0 Then strHtml = PE_Replace(strHtml, "{$Hits}", GetHits(rsPhoto("InfoPoint"), rsPhoto("Hits"), 0))
    If InStr(strHtml, "{$UpdateDate}") > 0 Then strHtml = PE_Replace(strHtml, "{$UpdateDate}", FormatDateTime(rsPhoto("UpdateTime"), 2))
    strHtml = PE_Replace(strHtml, "{$UpdateTime}", rsPhoto("UpdateTime"))
    strHtml = PE_Replace(strHtml, "{$Inputer}", rsPhoto("Inputer"))
    strHtml = PE_Replace(strHtml, "{$Editor}", rsPhoto("Editor"))
    If InStr(strHtml, "{$Stars}") > 0 Then strHtml = PE_Replace(strHtml, "{$Stars}", GetStars(rsPhoto("Stars")))
    If InStr(strHtml, "{$PhotoProperty}") > 0 Then strHtml = PE_Replace(strHtml, "{$PhotoProperty}", GetPhotoProperty())
    strHtml = PE_Replace(strHtml, "{$Rss}", "")
    If InStr(strHtml, "{$Keyword}") > 0 Then strHtml = PE_Replace(strHtml, "{$Keyword}", GetKeywords(",", rsPhoto("Keyword")))
    If InStr(strHtml, "{$InfoPoint}") > 0 Then strHtml = PE_Replace(strHtml, "{$InfoPoint}", GetInfoPoint(rsPhoto("InfoPoint")))
    If InStr(strHtml, "{$PhotoPoint}") > 0 Then strHtml = PE_Replace(strHtml, "{$PhotoPoint}", GetInfoPoint(rsPhoto("InfoPoint")))
    If InStr(strHtml, "{$PhotoIntro}") > 0 Then strHtml = PE_Replace(strHtml, "{$PhotoIntro}", ReplaceKeyLink(ReplaceText(rsPhoto("PhotoIntro"), 1)))
    '�滻{$PhotoIntro(Type,InfoLength)}��ǩ
    Dim strPhotoIntro
    regEx.Pattern = "\{\$PhotoIntro\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strPhotoIntro= "����ʽ��ǩ��{$PhotoIntro(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"

        Else
            Select Case PE_Clng(arrTemp(0))
            Case 1
                strPhotoIntro = ReplaceKeyLink(ReplaceText(rsPhoto("PhotoIntro"), 1))
            Case 2
                If PE_Clng(arrTemp(1))>0 then
                    strPhotoIntro = GetSubStr(nohtml(rsPhoto("PhotoIntro")),PE_Clng(arrTemp(1)),False)
                Else
                    strPhotoIntro = nohtml(rsPhoto("PhotoIntro"))
                End IF
            End Select
        End If
        strHtml = Replace(strHtml, Match.Value, strPhotoIntro)
	Next

    strHtml = Replace(strHtml, "{$PhotoProtect}", "")
    If InStr(strHtml, "{$PhotoThumb}") > 0 Then strHtml = Replace(strHtml, "{$PhotoThumb}", GetPhotoThumb(rsPhoto("PhotoThumb"), 130, 0))
    If InStr(strHtml, "{$PhotoName}") > 0 Then strHtml = Replace(strHtml, "{$PhotoName}", ReplaceText(rsPhoto("PhotoName"), 2))
    strHtml = Replace(strHtml, "{$PhotoSize}", "")
    If InStr(strHtml, "{$DayHits}") > 0 Then strHtml = Replace(strHtml, "{$DayHits}", GetHits(rsPhoto("InfoPoint"), rsPhoto("DayHits"), 1))
    If InStr(strHtml, "{$WeekHits}") > 0 Then strHtml = Replace(strHtml, "{$WeekHits}", GetHits(rsPhoto("InfoPoint"), rsPhoto("WeekHits"), 2))
    If InStr(strHtml, "{$MonthHits}") > 0 Then strHtml = Replace(strHtml, "{$MonthHits}", GetHits(rsPhoto("InfoPoint"), rsPhoto("MonthHits"), 3))
    If InStr(strHtml, "{$PrevPhotoUrl}") > 0 Then strHtml = Replace(strHtml, "{$PrevPhotoUrl}", GetPrevPhoto())
    If InStr(strHtml, "{$NextPhotoUrl}") > 0 Then strHtml = Replace(strHtml, "{$NextPhotoUrl}", GetNextPhoto())
    If InStr(strHtml, "{$Vote}") > 0 Then strHtml = Replace(strHtml, "{$Vote}", GetVoteOfContent(PhotoID)) 'ͶƱ��ǩ
    If InStr(strHtml, "{$CorrelativePhoto}") > 0 Then strHtml = Replace(strHtml, "{$CorrelativePhoto}", GetCorrelative(5, 50))

    Dim arrTemp
    Dim strPhotoThumb
    '�滻ͼƬ����ͼ
    regEx.Pattern = "\{\$PhotoThumb\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strPhotoThumb = GetPhotoThumb(Trim(rsPhoto("PhotoThumb")), PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
        strHtml = Replace(strHtml, Match.value, strPhotoThumb)
    Next

    Dim strCorrelativePhoto
    regEx.Pattern = "\{\$CorrelativePhoto\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strCorrelativePhoto = GetCorrelative(arrTemp(0), arrTemp(1))
        strHtml = Replace(strHtml, Match.value, strCorrelativePhoto)
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
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��ר��", False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��ר��", False))
    
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

Public Sub GetHtml_Search()
    Dim SearchChannelID
    SearchChannelID = ChannelID
    If ChannelID > 0 Then
        strHtml = GetTemplate(ChannelID, 5, 0)
    Else
        strHtml = GetTemplate(ChannelID, 3, 0)
        ChannelID = PE_CLng(Conn.Execute("select min(ChannelID) from PE_Channel where ModuleType=3 and Disabled=" & PE_False & "")(0))
        Call GetChannel(ChannelID)
    End If
    Select Case strField
    Case "Title"
        strField = "PhotoName"
    Case "Content"
        strField = "PhotoIntro"
    End Select
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

Public Sub ShowFavorite()
    Dim sqlFavorite, rsFavorite, iCount
    iCount = 0
    
    Response.Write "<table width='100%' cellpadding='0' cellspacing='5' border='0' align='center' class='border'><tr valign='top' class='tdbg'>"
    
    sqlFavorite = "select P.PhotoID,P.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,P.InfoPurview,P.PhotoName,P.UpdateTime,P.PhotoThumb,P.InfoPoint from PE_Photo P inner join PE_Class C on P.ClassID=C.ClassID where P.Deleted=" & PE_False & " and P.Status=3 "
    sqlFavorite = sqlFavorite & " and PhotoID in (select InfoID from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & ")"
    sqlFavorite = sqlFavorite & " order by P.PhotoID desc"

    Set rsFavorite = Server.CreateObject("ADODB.Recordset")
    rsFavorite.Open sqlFavorite, Conn, 1, 1
    If rsFavorite.BOF And rsFavorite.EOF Then
        totalPut = 0
        Response.Write "<td align='center'><img class='pic3' src='" & InstallDir & "images/nopic.gif' width='130' height='90' border='0'><br>û���ղ��κ�" & ChannelShortName & "</td>"
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
            Response.Write "<td align='center'><a href='" & GetPhotoUrl(rsFavorite("ParentDir"), rsFavorite("ClassDir"), rsFavorite("UpdateTime"), rsFavorite("PhotoID"), rsFavorite("ClassPurview"), rsFavorite("InfoPurview"), rsFavorite("InfoPoint")) & "' target='_blank'>"
            Response.Write GetPhotoThumb(rsFavorite("PhotoThumb"), 130, 90)
            Response.Write "<br>" & rsFavorite("PhotoName") & "</a>"
            Response.Write "</td>"
            rsFavorite.MoveNext
            iCount = iCount + 1
            If iCount >= MaxPerPage Then Exit Do
            If ((iCount Mod 4 = 0) And (Not rsFavorite.EOF)) Then Response.Write "</tr><tr valign='top' class='tdbg'>"
        Loop
    End If
    rsFavorite.Close
    Set rsFavorite = Nothing
    Response.Write "</tr></table>"
    Response.Write ShowPage("User_Favorite.asp?ChannelID=" & ChannelID & "", totalPut, 20, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False)
End Sub

Function XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    XmlText_Class = XmlText("Photo", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Photo", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>
