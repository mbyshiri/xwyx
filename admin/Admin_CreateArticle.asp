<!--#include file="Admin_CreateCommon.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim PE_Content
Set PE_Content = New Article
PE_Content.Init
tmpPageTitle = strPageTitle    '����ҳ����⵽��ʱ�����У�����Ϊ��Ŀ������ҳѭ������ʱ��ʼֵ
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
        '��������ҳǰ��Ҫ����ĿID��ר��ID��Ϊ0
        ClassID = 0
        arrChildID = 0
        SpecialID = 0
        Call CreateIndex

        Call CreateSiteIndex     '������վ��ҳ
        Call CreateSiteSpecial   '����ȫվר��
    End If
Case "CreateOther" '��ʱ���ɴ�������������ҳ
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
            Call CreateSiteIndex    '������վ��ҳ
        End If


        TimingCreateNum = TimingCreateNum + 1
        strFileName = "Admin_Timing.asp?Action=DoTiming&TimingCreateNum=" & TimingCreateNum & "&TimingCreate=" & Trim(Request("TimingCreate"))
    Else    '�ɼ�������
        CreateNum = PE_CLng(Trim(Request("CreateNum")))
        Call CreateClass
        Call CreateSpecial
        Call CreateIndex
        Call CreateSiteIndex     '������վ��ҳ
        '��������JS
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
    ErrMsg = ErrMsg & "<li>��������</li>"
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
        Response.Write "<b>��������" & ChannelShortName & "ҳ�桭�����Ժ�<font color='red'>�ڴ˹���������ˢ�´�ҳ�棡����</font></b><br>"
        Response.Flush
    End If

    Select Case CreateType
    Case 1 'ѡ��������
        If IsValidID(ArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷָ��Ҫ���ɵ�" & ChannelShortName & "ID</li>"
            Exit Sub
        End If
        If InStr(ArticleID, ",") > 0 Then
            sql = sql & " and ArticleID in (" & ArticleID & ")"
        Else
            sql = sql & " and ArticleID=" & ArticleID
        End If
        strUrlParameter = "&ArticleID=" & ArticleID
    Case 2 'ѡ������Ŀ
        ClassID = PE_CLng(Trim(Request("ClassID")))
        If ClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ���ɵ���ĿID</li>"
            Exit Sub
        End If
        Call GetClass
        If ClassPurview > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ŀ���ǿ�����Ŀ�����Դ���Ŀ�µ����²�������HTML��"
        End If
        If FoundErr = True Then Exit Sub
        If InStr(arrChildID, ",") > 0 Then
            sql = sql & " and ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and ClassID=" & ClassID
        End If
    Case 3 '��������
        
    Case 4 '���µ�����
        Dim TopNew
        TopNew = PE_CLng(Trim(Request("TopNew")))
        If TopNew <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����Ч����Ŀ��"
            Exit Sub
        End If
        sql = "select top " & TopNew & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID & ""
        strUrlParameter = "&TopNew=" & TopNew
    Case 5 'ָ������ʱ��
        Dim BeginDate, EndDate
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        If Not (IsDate(BeginDate) And IsDate(EndDate)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������Ч�����ڣ�</li>"
            Exit Sub
        End If
        If SystemDatabaseType = "SQL" Then
            sql = sql & " and UpdateTime between '" & BeginDate & "' and '" & EndDate & "'"
        Else
            sql = sql & " and UpdateTime between #" & BeginDate & "# and #" & EndDate & "#"
        End If
        strUrlParameter = "&BeginDate=" & Replace(BeginDate,"/","-") & "&EndDate=" & Replace(EndDate,"/","-")
    Case 6 'ָ��ID��Χ
        Dim BeginID, EndID
        BeginID = Trim(Request("BeginID"))
        EndID = Trim(Request("EndID"))
        If Not (IsNumeric(BeginID) And IsNumeric(EndID)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������֣�</li>"
            Exit Sub
        End If
        sql = sql & " and ArticleID between " & BeginID & " and " & EndID & ""
        strUrlParameter = "&BeginID=" & BeginID & "&EndID=" & EndID
    Case 7 '�ɼ���������
        TimingCreate = Trim(Request("TimingCreate"))
        CollectionCreateHTML = Trim(Request("CollectionCreateHTML"))
        CreateNum = PE_CLng(Trim(Request("CreateNum")))
        IsShowReturn = True

        If CollectionCreateHTML = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>��ָ��Ҫ���ɵ���Ŀ��"
            Exit Sub
        Else
            ChannelID = PE_CLng(Trim(Request("ChannelID")))
            ClassID = PE_CLng(Trim(Request("ClassID")))
            SpecialID = ReplaceBadChar(Trim(Request("SpecialID")))
            ArticleNum = PE_CLng(Trim(Request("ArticleNum")))

            sql = "select top " & ArticleNum & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and ChannelID=" & ChannelID & " and ClassID=" & ClassID & ""
        End If
        strUrlParameter = "&CollectionCreateHTML=" & CollectionCreateHTML & "&CreateNum=" & CreateNum & "&ArticleNum=" & ArticleNum & "&TimingCreate=" & TimingCreate

    Case 8 '��ʱ��������
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
    Case 9 '����δ���ɵ�����
        sql = "select top " & MaxPerPage_Create & " * from PE_Article where Deleted=" & PE_False & " and Status=3 and ReceiveType=0 and InfoPoint=0 And InfoPurview=0 and ChannelID=" & ChannelID & ""
		sql = sql & " and (CreateTime is null or CreateTime<=UpdateTime)"
    Case Else
        Response.Write "��������"
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

    PageTitle = "����" '�õ�Ƶ������
    strFileName = ChannelUrl_ASPFile & "/ShowArticle.asp" '�õ�·��
    strTemplate = GetTemplate(ChannelID, 3, tmpTemplateID) '�õ�Ƶ�������ĵ�Ĭ��ģ��
    
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
            Response.Write "<li><font color='red'>IDΪ " & rsArticle("ArticleID") & " ��" & ChannelShortName & "��Ϊ�������Ķ�Ȩ�ޣ�����û�����ɡ�</font></li>"
            Response.Flush
        Else
            SpecialID = 0
            CurrentPage = 1
            ArticlePath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle("UpdateTime"))

            If CreateMultiFolder(ArticlePath) = False Then
                Response.Write "�����������ϵͳ���ܴ��������ļ�����Ҫ���ļ��С�"
                Exit Sub
            End If
            ArticlePath = ArticlePath & GetItemFileName(FileNameType, ChannelDir, rsArticle("UpdateTime"), ArticleID)
                
            tmpFileName = ArticlePath & FileExt_Item

            '����ҳ��ʱ�ж�ת������
            If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then
                Call WriteToFile(tmpFileName, PE_Content.GetLinkUrlContent(rsArticle("LinkUrl"), ArticleID))
                Response.Write "<li>�ɹ����ɵ� <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & "��&nbsp;&nbsp;ID��" & ArticleID & " &nbsp;&nbsp;���⣺" & rsArticle("Title") & " &nbsp;&nbsp;��ַ��<a href='" & tmpFileName & "' target='_blank'>" & tmpFileName & "</a></li>" & vbCrLf
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
                'д�����ɵ�ַ
                Call WriteToFile(tmpFileName, tmpArticle)
                Response.Write "<li>�ɹ����ɵ� <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & "��&nbsp;&nbsp;ID��" & ArticleID & " &nbsp;&nbsp;���⣺" & rsArticle("Title") & " &nbsp;&nbsp;��ַ��<a href='" & tmpFileName & "' target='_blank'>" & tmpFileName & "</a></li>" & vbCrLf
                Response.Flush
                
                For CurrentPage = 2 To PE_Content.TotalPage
                    tmpFileName = ArticlePath & "_" & CurrentPage & FileExt_Item
                    tmpArticle = PE_Content.ReplaceContentLabel(strHtml)
                    If InStr(tmpArticle, "{$ShowPageContent}") > 0 Then tmpArticle = Replace(tmpArticle, "{$ShowPageContent}", "")
                    Call WriteToFile(tmpFileName, tmpArticle)
                    Response.Write "<br>&nbsp;&nbsp;&nbsp;�ɹ����ɵ� <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & "�ĵ� <font color='blue'>" & CurrentPage & "</font> ҳ��<a href='" & tmpFileName & "' target='_blank'>" & tmpFileName & "</a>" & vbCrLf
                    Response.Flush
                Next
            End If
            '�������ݽ������������ݵ�����ʱ��
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
