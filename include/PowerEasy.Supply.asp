<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim SupplyID, rsSupply, SearchStyleFlag

Class Supply

Private OpenType

Public Sub Init()

    'ClassID = PE_CLng(Trim(Request("ClassID")))
    'SpecialID = PE_CLng(Trim(Request("SpecialID")))
    SupplyID = PE_CLng(Trim(Request("SupplyId")))
    
    If IsValidID(SupplyID) = False Then
        SupplyID = ""
    End If
    PrevChannelID = ChannelID
        
    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
     
    ChannelShortName = "������Ϣ"
    
    strNavPath = XmlText("BaseText", "Nav", "�����ڵ�λ�ã�") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
    strPageTitle = SiteTitle
    
    Call GetChannel(ChannelID)
    
    If Trim(ChannelName) <> "" And ShowNameOnPath <> False Then
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath' href='" & ChannelUrl & "/Index.asp"
        strNavPath = strNavPath & "'>" & ChannelName & "</a>"
        strPageTitle = strPageTitle & " >> " & ChannelName
    End If
End Sub

'��ǩ�����ӿ�
Private Function getInfoListLable()
    regEx.Pattern = "\{\$SupplyInfoList\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        strHtml = Replace(strHtml, Match.value, ReplaceInfoListLabel(Match.SubMatches(0)))
    Next

    regEx.Pattern = "\{\$SupplyInfoType\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        strHtml = Replace(strHtml, Match.value, ReplaceSupplyInfoType(Match.SubMatches(0)))
    Next

    regEx.Pattern = "\{\$Navigation\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        strHtml = Replace(strHtml, Match.value, ReplaceNavigationLabel(Match.SubMatches(0)))
    Next

End Function

Private Function ReplaceSupplyInfoType(ByVal strTemp)
    Dim arrTemp
    arrTemp = Split(strTemp, ",")
    
    If PE_CLng(arrTemp(0)) > 4 Or PE_CLng(arrTemp(0)) < 0 Then
        arrTemp(0) = 0
    End If
    If UBound(arrTemp) <> 6 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ҳ���ǩ{$SupplyInfoType(...)}����</li>"
        Exit Function
    End If
    If PE_CLng(arrTemp(1)) = 0 Then
        arrTemp(1) = 10
    End If
    Select Case PE_CLng(arrTemp(0))
        Case 0  ' ����
            ReplaceSupplyInfoType = getLasterSupplyInfo(PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)))
        Case 1  '����
            ReplaceSupplyInfoType = getHotSupplyInfo(PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)))
        Case 2  '�Ƽ�
            ReplaceSupplyInfoType = getCommandSupplyInfo(PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)))
        Case 3  '��ͼƬ��������Ϣ
            ReplaceSupplyInfoType = getPicLasterInfo(PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)))
    End Select
End Function

'���ݱ�ǩ���������滻��ͬ�ĺ���
Private Function ReplaceInfoListLabel(ByVal strTemp)    '0,0,1,100,0,True
    Dim arrTemp
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) = 10 Then
        strTemp = strTemp & ",0,True"
    End If
    If CheckSupplyLabel(strTemp) Then
        Exit Function
    End If
    arrTemp = Split(strTemp, ",")
    If PE_CLng(arrTemp(5)) > 3 Or PE_CLng(arrTemp(5)) < 0 Then
        arrTemp(5) = 0
    End If
    If ClassID > 0 Then
        arrTemp(1) = ClassID
    End If
    Select Case PE_CLng(arrTemp(5))
        Case 0
            ReplaceInfoListLabel = getInfoList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12))) 'һ�ж���
        Case 1
            ReplaceInfoListLabel = getDetailInfoList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12))) 'һ��
        Case 2
            ReplaceInfoListLabel = getListPicInfoList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12))) 'ͼƬ��ʽһ
        Case 3
            ReplaceInfoListLabel = getPicInfoList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12))) 'ͼƬ��ʽ��
    End Select
End Function

Private Function CheckSupplyLabel(ByVal strTemp)
    Dim arrTemp
    arrTemp = Split(strTemp, ",")
    CheckSupplyLabel = False
    
    If UBound(arrTemp) <> 12 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ҳ���ǩ{$SupplyInfoList(...)}����̫�����̫�ٴ���</li>"
        CheckSupplyLabel = True
    Else
        If PE_CLng(arrTemp(11)) > PE_CLng(getSupplyTypeNum("//SupplyType/Type")) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ҳ���ǩ{$SupplyInfoList(...)}�ĵ�12����������</li>"
            CheckSupplyLabel = True
        End If
    End If
End Function

Private Function getSupplyTypeNum(ByVal NodeName)
    Dim LangRoot, strTemp, XmlDoc, ShowLength
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    XmlDoc.async = False
    XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
    Set LangRoot = XmlDoc.selectNodes(NodeName)
    getSupplyTypeNum = LangRoot.Length
End Function

Private Function ReplaceNavigationLabel(ByVal strTemp)
    Dim arrTemp
    arrTemp = Split(strTemp, ",")
    ReplaceNavigationLabel = GetClass_Navigation(2, PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
End Function

Public Function ReplaceSearchCondition(ByVal strTemp)
    Dim arrTemp
    arrTemp = Split(strTemp, ",")
    ReplaceSearchCondition = ShowSearchCondition(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
End Function

Private Function getClassInfoNum(ByVal ClassID)
    Dim strSql
    strSql = "Select Count(*) From PE_Supply Where ClassId IN (" & ClassID & ")"
    getClassInfoNum = PE_CLng(Conn.Execute(strSql)(0))
End Function


Public Sub GetHtml_Supply()
    If PrevChannelID <> ChannelID Then
        Call GetChannel(ChannelID)
    End If
    strHtml = PE_Replace(strHtml, "{$SupplyID}", SupplyID)
    Call ReplaceCommon
    strHtml = PE_Replace(strHtml, "{$ClassUrl}", GetClassUrl(ClassID))
    strHtml = Replace(strHtml, "{$SupplyAction}", GetSupplyAction())
    strPageTitle = rsSupply("SupplyTitle")
    strHtml = Replace(strHtml, "{$SupplyInfoType}", GetSupplyInfoType(rsSupply("SupplyType"), "//SupplyType/Type"))
    strHtml = Replace(strHtml, "{$SupplyInfoTitle}", rsSupply("SupplyTitle"))
    strHtml = Replace(strHtml, "{$TradeType}", GetSupplyInfoType(rsSupply("TradeType"), "//TradeType/Type"))
    strHtml = PE_Replace(strHtml, "{$SupplyName}", rsSupply("SupplyName"))
    strHtml = PE_Replace(strHtml, "{$PriceIntro}", rsSupply("PriceIntro"))
    strHtml = PE_Replace(strHtml, "{$UpdateTime}", rsSupply("UpdateTime"))
    If rsSupply("SupplyPeriod") <> -1 Then
        strHtml = PE_Replace(strHtml, "{$EndTime}", DateAdd("d", rsSupply("SupplyPeriod"), rsSupply("UpdateTime")))
    Else
        strHtml = PE_Replace(strHtml, "{$EndTime}", "������Ч")
    End If
    strHtml = Replace(strHtml, "{$SupplyIntro}", Replace(Replace(rsSupply("SupplyIntro"), "[InstallDir_ChannelDir]", ChannelUrl & "/"), "{$UploadDir}", UploadDir))
    strHtml = PE_Replace(strHtml, "{$UserName}", rsSupply("UserName"))
    strHtml = PE_Replace(strHtml, "{$Province}", rsSupply("Province"))
    strHtml = PE_Replace(strHtml, "{$City}", rsSupply("City"))
    strHtml = PE_Replace(strHtml, "{$Address}", rsSupply("Address"))
    strHtml = PE_Replace(strHtml, "{$ZipCode}", rsSupply("ZipCode"))
    strHtml = PE_Replace(strHtml, "{$Email}", rsSupply("Email"))
    strHtml = PE_Replace(strHtml, "{$CompanyName}", rsSupply("Company"))
    strHtml = PE_Replace(strHtml, "{$Department}", rsSupply("Department"))
    strHtml = PE_Replace(strHtml, "{$CompanyAddress}", rsSupply("CompanyAddress")) '��˾��ַ
    strHtml = PE_Replace(strHtml, "{$RealName}", rsSupply("TrueName")) '��ʵ����
    strHtml = PE_Replace(strHtml, "{$Sex}", getUserSex(rsSupply("Sex"))) '�Ա�
    strHtml = PE_Replace(strHtml, "{$Position}", rsSupply("Position")) 'ְ��
    strHtml = PE_Replace(strHtml, "{$Operation}", rsSupply("Operation")) '�����ҵ��
    strHtml = PE_Replace(strHtml, "{$OfficePhone}", rsSupply("OfficePhone")) '�칫�ҵ绰
    strHtml = PE_Replace(strHtml, "{$Fax}", rsSupply("Fax"))  '����
    strHtml = PE_Replace(strHtml, "{$Mobile}", rsSupply("Mobile")) '�ƶ��绰
    strHtml = PE_Replace(strHtml, "{$QQ}", rsSupply("QQ")) 'qq
    strHtml = PE_Replace(strHtml, "{$Msn}", rsSupply("Msn")) 'msn
    strHtml = PE_Replace(strHtml, "{$Homepage}", rsSupply("Homepage")) '��ַA.LoginTimes,A.LastLoginTime
    strHtml = PE_Replace(strHtml, "{$LoginTimes}", rsSupply("LoginTimes")) '��¼����
    strHtml = PE_Replace(strHtml, "{$LastLoginTime}", rsSupply("LastLoginTime")) '�����¼ʱ��
    strHtml = PE_Replace(strHtml, "{$UserType}", getUserType(rsSupply("UserType"))) '��Ա����
End Sub


Public Sub GetHtml_Special()
    strHtml = PE_Replace(strHtml, "{$SpecialID}", SpecialID)
    Call ReplaceCommon
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)
    strHtml = PE_Replace(strHtml, "{$SpecialPicUrl}", SpecialPicUrl)

    Dim strPath
    strPath = ChannelUrl & "/Special/" & SpecialDir
    Call getInfoListLable
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
End Sub


Public Sub GetHtml_Class()
    Dim strTemp, iCols, iClassID
    If Child > 0 And ClassShowType <> 2 Then
        strHtml = arrTemplate(0)
    Else
        strHtml = arrTemplate(1)
    End If
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommon
    strHtml = PE_Replace(strHtml, "{$ClassPicUrl}", ClassPicUrl)
    strHtml = PE_Replace(strHtml, "{$Meta_Keywords_Class}", Meta_Keywords_Class)
    strHtml = PE_Replace(strHtml, "{$Meta_Description_Class}", Meta_Description_Class)
    strHtml = Replace(strHtml, "{$ClassUrl}", GetClassUrl(ClassID))
    strHtml = Replace(strHtml, "{$ClassListUrl}", GetClass_1Url(ClassID))
    
    Dim ArticleList_CurrentClass, ArticleList_CurrentClass2, ArticleList_ChildClass, ArticleList_ChildClass2
    If Child > 0 And ClassShowType <> 2 Then    '�����ǰ��Ŀ������Ŀ
        If InStr(strHtml, "{$ShowChildClass}") > 0 Then strHtml = Replace(strHtml, "{$ShowChildClass}", GetChildClass(0, 0, 3, 3, 0, True))
        
        Dim strChildClass, arrTemp
        regEx.Pattern = "\{\$ShowChildClass\((.*?)\)\}"
        Set Matches = regEx.Execute(strHtml)
        For Each Match In Matches
            arrTemp = Split(Match.SubMatches(0), ",")
            strChildClass = GetChildClass(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
            strHtml = Replace(strHtml, Match.value, strChildClass)
        Next
        
        ItemCount = PE_CLng(Conn.Execute("select Count(*) from PE_Supply where ClassID=" & ClassID & "")(0))
        If ItemCount <= 0 Then     '�����ǰ��Ŀû������
            strHtml = regEx.Replace(strHtml, "") '��ȥ����ʾ��ǰ��Ŀ��ֻ���ڱ���Ŀ�������б�
        Else      '�����ǰ��Ŀ������Ŀ���ҵ�ǰ��Ŀ�����ݣ�����Ҫ��ʾ������
            strTemp = ArticleList_CurrentClass
            strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(ClassID))
            strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", ReadMe)
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", ClassName)
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", ClassID)
        End If
     
        '�õ�ÿ����ʾ������
        iCols = 1
        regEx.Pattern = "��Cols=(\d{1,2})��"
        Set Matches = regEx.Execute(ArticleList_ChildClass)
        ArticleList_ChildClass = regEx.Replace(ArticleList_ChildClass, "")
        For Each Match In Matches
            If Match.SubMatches(0) > 1 Then iCols = Match.SubMatches(0)
        Next
        '��ʼѭ�����õ���������Ŀ�б��HTML����
        iClassID = 0
        Dim rsClass
        Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=" & ClassID & " and IsElite=" & PE_True & " and ClassType=1 order by RootID,OrderID")
        Do While Not rsClass.EOF
            strTemp = ArticleList_ChildClass
            strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(rsClass("ClassID")))
            strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", rsClass("Readme"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", rsClass("ClassName"))
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
    End If
    Dim strPath
    strPath = ChannelUrl
    Call getInfoListLable
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    
End Sub


Public Sub GetHtml_List()
    Call ReplaceCommon
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)
    Call getInfoListLable
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
End Sub


Public Sub GetHTML_Index()
    strHtml = GetTemplate(ChannelID, 1, Template_Index)
    ClassID = 0
    Call ReplaceCommon
    strHtml = Replace(strHtml, "{$ShowChannelCount}", GetChannelCount())
    Call getInfoListLable
End Sub

Public Sub GetHTML_Search()
    Dim SearchChannelID
    SearchChannelID = ChannelID
    If ChannelID > 0 Then
        strHtml = GetTemplate(ChannelID, 5, 0)
    Else
        ChannelID = PE_CLng(Conn.Execute("select min(ChannelID) from PE_Channel where ModuleType=1 and Disabled=" & PE_False & "")(0))
        strHtml = GetTemplate(ChannelID, 3, 0)
        CurrentChannelID = ChannelID
        Call GetChannel(ChannelID)
    End If
    ClassID = PE_CLng(Trim(Request("SelClass")))
    Call ReplaceCommon
    strHtml = Replace(strHtml, "{$Keyword}", Keyword)
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    Call GetClass
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)

    strHtml = Replace(strHtml, "{$SearchResul}", SearchResult())
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage_SearchResult, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage_SearchResult, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    Call getInfoListLable
End Sub

Private Sub ReplaceCommon()

    Call ReplaceCommonLabel
    
    strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, ShowClassTreeGuide))
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))

    strHtml = PE_Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = PE_Replace(strHtml, "{$ShowPath}", ShowPath())

    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
End Sub

Private Function getUserType(ByVal UserType)
    If PE_CLng(UserType) = 0 Then
        getUserType = "���˻�Ա"
    Else
        getUserType = "��ҵ��Ա"
    End If
End Function

Private Function getUserSex(ByVal UserSex)
    Select Case PE_CLng(UserSex)
        Case 0
            getUserSex = "Ů"
        Case 1
            getUserSex = "��"
        Case Else
            getUserSex = "����"
    End Select
    
End Function


Private Function ShowSearchCondition(ByVal ShowStyle, ByVal SupplyTypeNum)
    Dim strTemp
    strTemp = "<script language='javascript'>" & vbCrLf
    strTemp = strTemp & "function CheckInput()" & vbCrLf
    strTemp = strTemp & "{" & vbCrLf
    strTemp = strTemp & "   if(document.Searchform.KeyWord.value=='')" & vbCrLf
    strTemp = strTemp & "   {" & vbCrLf
    strTemp = strTemp & "     alert('�ؼ��ֲ�����Ϊ��!');" & vbCrLf
    strTemp = strTemp & "     return false;" & vbCrLf
    strTemp = strTemp & "    }" & vbCrLf
    strTemp = strTemp & "}" & vbCrLf
    strTemp = strTemp & "</script>" & vbCrLf
    SearchStyleFlag = ShowStyle
    Select Case ShowStyle
        Case 0
            strTemp = strTemp & "<table cellSpacing=0 cellPadding=2 width='100%' align=center border=0>"
            strTemp = strTemp & "<Form Name='Searchform' action='Search.asp' method=GET onSubmit='return CheckInput();'>"
            strTemp = strTemp & "<tr><td>" & GetSupplyInfo_Radio(0, "SupplyType", "//SupplyType/Type", SupplyTypeNum) & "<INPUT Class='SupplySearchKeyWordStyle' TYPE=""Text"" NAME=""KeyWord"" Id=""KeyWord"" MaxLength='250' >&nbsp;&nbsp;<INPUT TYPE=""submit"" Value=""������Ϣ""></td></tr>"
            strTemp = strTemp & "<Input Type='Hidden' value='" & ShowStyle & "' Name='SearchType' id='SearchType'>"
            strTemp = strTemp & "</Form>"
            strTemp = strTemp & "</Table>"
            ShowSearchCondition = strTemp
        Case 1
            strTemp = strTemp & "<table cellSpacing=0 cellPadding=2 width='100%' align=center border=0>"
            strTemp = strTemp & "<tr align=left>"
            strTemp = strTemp & "<Form Name='Searchform' action='Search.asp' method='GET' onSubmit='return CheckInput();'>"
            strTemp = strTemp & "<td>�����ؼ��֣�<Input Class='SupplySearchKeyWordStyle' Name='KeyWord' id='KeyWord' type='text' MaxLength='200' size='20'>"
            strTemp = strTemp & "&nbsp;&nbsp;<Select Name='selClass' id='selClass'><option Value=-1>������Ŀ</option>" & GetClass_Option(0) & "</Select></td>"
            strTemp = strTemp & "<td>������:<Select Name='mySelectProvince' Id='mySelectProvince' OnChange=""getSelected('Region.asp?Province='+this.value,mySelectCity)""><Option value=-1>����ʡ��</Option></Select><Select Name='mySelectCity' Id='mySelectCity' ><option value=-1>���г���</option></Select></td>"
            strTemp = strTemp & "<td><INPUT TYPE=""submit"" Value=""������Ϣ""></td>"
            strTemp = strTemp & "<Input Type='Hidden' value='" & ShowStyle & "' Name='SearchType' id='SearchType'>"
            strTemp = strTemp & "</form></tr>"
            strTemp = strTemp & "</Table>"
            ShowSearchCondition = strTemp
   End Select
End Function

Private Function SearchResult()
    Dim Province, City, SelClass, SupplyType, SearchType, QuerySql, strSql, strTemp
    Province = ReplaceBadChar(Trim(Request("mySelectProvince")))
    City = ReplaceBadChar(Trim(Request("mySelectCity")))
    SelClass = PE_CLng(Trim(Request("SelClass")))
    SearchType = PE_CLng(Trim(Request("SearchType")))
    SupplyType = PE_CLng(Trim(Request("SupplyType")))
    
    strSql = "Select Top " & MaxPerPage_SearchResult & " A.SupplyId,A.SupplyTitle,A.SupplyName,A.SupplyType,A.TradeType, "
    strSql = strSql & " A.UpDateTime,B.Country,B.Province,B.City From PE_Supply A , PE_Contacter B, PE_User C"
    QuerySql = " Where  A.UserName=C.UserName and C.ContacterID = B.ContacterID and A.Status=1 And Deleted=" & PE_False & " "
    Select Case SearchType
        Case 0
            QuerySql = QuerySql & " And A.SupplyType =" & SupplyType
            If CurrentPage > 1 Then
                strSql = strSql & QuerySql & " And A.SupplyId <=(Select Min(SupplyId) From (Select Top " & (CurrentPage - 1) * MaxPerPage_SearchResult + 1 & " SupplyId From PE_Supply A,PE_Contacter B,PE_User C " & QuerySql & " And A.SupplyType =" & SupplyType & " And A.SupplyTitle like '%" & Keyword & "%' Order By SupplyId DESC) As QueryTable) Order by A.SupplyId Desc"
            Else
                strSql = strSql & QuerySql & " And A.SupplyType =" & SupplyType & " And A.SupplyTitle like '%" & Keyword & "%' Order By A.SupplyID Desc"
            End If
            strFileName = ChannelUrl & "/Search.asp?SupplyType=" & SupplyType & "&KeyWord=" & Keyword & "&SearchType=" & SearchType & ""
            totalPut = PE_CLng(Conn.Execute("Select Count(*) From PE_Supply A , PE_Contacter B, PE_User C Where  A.UserName=C.UserName and C.ContacterID = B.ContacterID and A.Status=1 And A.SupplyType =" & SupplyType & " And A.SupplyTitle like '%" & Keyword & "%' ")(0))
        Case 1
            If Province <> "-1" Then
                QuerySql = QuerySql & "And B.Province='" & Province & "'"
            End If
            If City <> "-1" Then
                QuerySql = QuerySql & " And B.City = '" & City & "'"
            End If
            If SelClass > -1 Then
                 If Child > 0 Then
                    QuerySql = QuerySql & " and A.ClassID in (" & arrChildID & ")"
                Else
                    QuerySql = QuerySql & " and A.ClassID=" & ClassID
                End If
            End If
            If CurrentPage > 1 Then
                strSql = strSql & QuerySql & "And A.SupplyId <=(Select Min(SupplyId) From (Select Top " & (CurrentPage - 1) * MaxPerPage_SearchResult + 1 & " SupplyId From PE_Supply A,PE_Contacter B,PE_User C " & QuerySql & " And A.SupplyTitle like '%" & Keyword & "%' Order by A.SupplyId DESC) As QueryTable) Order By A.SupplyID DESC"
            Else
                strSql = strSql & QuerySql & " And A.SupplyTitle like '%" & Keyword & "%' Order By A.SupplyID DESC "
            End If
            totalPut = PE_CLng(Conn.Execute("Select Count(*) From PE_Supply A , PE_Contacter B, PE_User C " & QuerySql & " And A.SupplyTitle like '%" & Keyword & "%'")(0))
            strFileName = ChannelUrl & "/Search.asp?KeyWord=" & Keyword & "&selClass=" & SelClass & "&mySelectProvince=" & Province & "&mySelectCity=" & City & "&SearchType=" & SearchType & ""
    End Select
    Dim rsSupply
    
    Set rsSupply = Server.CreateObject("Adodb.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    If rsSupply.EOF And rsSupply.BOF Then
        SearchResult = "<li>û���������κ���Ϣ</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            strTemp = strTemp & "<tr><td><Img src='" & ChannelUrl & "/Images/article_common.gif' border=0/>"
            strTemp = strTemp & "<font color=red>[" & GetSupplyInfoType(rsSupply("SupplyType"), "//SupplyType/Type") & "]</font>"
            strTemp = strTemp & "<a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & ">" & rsSupply("SupplyTitle") & "</a></td><td align='center'>" & GetSupplyInfoType(rsSupply("TradeType"), "//TradeType/Type") & "</td><td>" & rsSupply("Province") & "/" & rsSupply("City") & "</td><td>" & rsSupply("UpDateTime") & "</td></tr>" & vbCrLf
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    SearchResult = "<Table>" & strTemp & "</Table>"
End Function




'==================================================
'��������ShowPath
'��  �ã���ʾ������������λ�á�������Ϣ
'��  ������
'==================================================
Private Function ShowPath()
    If PageTitle <> "" Then
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
    End If
    ShowPath = strNavPath
End Function

Private Function GetClassUrl(iClassID)
    GetClassUrl = ChannelUrl & "/ShowClass.asp?ClassID=" & iClassID
End Function

Private Function GetClass_1Url(iClassID)
    GetClass_1Url = ChannelUrl & "/ShowClass.asp?ShowType=2&ClassID=" & iClassID
End Function


Private Function GetExecuteSql(ByVal PageType, ByVal ClassID, ByVal CommandType, ByVal IsNew, ByVal IsHot, ByVal ShowNum, ByVal SupplyType, ByVal Flage)
    Dim strSql, QuerySql, MaxShowInfo
    QuerySql = QuerySql & "From PE_Supply A,PE_Contacter B,PE_User C "
    QuerySql = QuerySql & "Where A.UserName=C.UserName and C.ContacterID = B.ContacterID and A.Status=1 And Deleted=" & PE_False & " "
    If Flage Then
        QuerySql = QuerySql & " And A.SupplyPicUrl<>'' "
    End If
    Select Case PageType
        Case 0
            MaxShowInfo = ShowNum
        Case 1 '��Ŀҳ
            MaxShowInfo = MaxPerPage
            If Child > 0 Then
                totalPut = getInfoCounts(arrChildID, 0, SupplyType)
            Else
                totalPut = getInfoCounts(ClassID, 0, SupplyType)
            End If
        Case 2  '�Ƽ�ҳ
            MaxShowInfo = MaxPerPage_Elite
            CommandType = 3
            If Flage Then
                totalPut = getInfoCounts(0, 5, SupplyType) '������ͼƬ���Ƽ���Ϣ
            Else
                totalPut = getInfoCounts(0, 2, SupplyType) '��ͼƬ���Ƽ���Ϣ
            End If
        Case 3 '�ȵ�ҳ
            MaxShowInfo = MaxPerPage_Hot
            totalPut = getInfoCounts(0, 3, SupplyType)
        Case 4 'ר��ҳ
            MaxShowInfo = MaxPerPage_Special
            totalPut = getInfoCounts(0, 4, SupplyType)
        Case 5 '����ҳ
            MaxShowInfo = MaxPerPage_New
            totalPut = getInfoCounts(0, 1, SupplyType) '���µ���Ϣ
        Case Else
            MaxShowInfo = MaxPerPage_Index
    End Select
    If ClassID > 0 Then
        If Child > 0 Then
            QuerySql = QuerySql & " and A.ClassID in (" & arrChildID & ")"
        Else
            QuerySql = QuerySql & " and A.ClassID=" & ClassID
        End If
    ElseIf SpecialID > 0 Then
        QuerySql = QuerySql & "And A.SpecialId=" & SpecialID
    End If
   
    Select Case CommandType
        Case 1
            QuerySql = QuerySql & "And A.CommandType=" & CommandType & " And DateDiff(" & PE_DatePart_D & ",A.UpdateTime," & PE_Now & ") < A.CommandChannelDays "
        Case 2
            QuerySql = QuerySql & "And A.CommandType=" & CommandType & " And DateDiff(" & PE_DatePart_D & ",A.UpdateTime," & PE_Now & ") < A.CommandClassDays "
        Case 3
            QuerySql = QuerySql & "And A.CommandType<>0"
    End Select
    If IsNew Then
        QuerySql = QuerySql & " And DateDiff(" & PE_DatePart_D & "," & PE_Now & ",A.UpdateTime)< " & DaysOfNew & ""
    End If

    If IsHot Then
        QuerySql = QuerySql & " And Hits >= " & HitsOfHot & ""
    End If
    If SupplyType >= 0 Then
        QuerySql = QuerySql & " And  SupplyType = " & SupplyType & ""
    End If

    If CurrentPage > 1 Then
        QuerySql = QuerySql & " and A.SupplyID < (select min(SupplyId) from (select top " & ((CurrentPage - 1) * MaxShowInfo) & " A.SupplyId " & QuerySql & " order by A.SupplyId desc) as QueryArticle) "
    End If
    strSql = "Select Top " & MaxShowInfo & " A.SupplyId,A.ClassId,A.SupplyTitle,A.SupplyName,A.SupplyType,"
    strSql = strSql & "A.TradeType,A.SupplyPicUrl,A.UpDateTime,B.Country,B.Province,B.City "
    
    strSql = strSql & QuerySql & " order by A.SupplyId desc "
    GetExecuteSql = strSql
End Function

'*****************************
'��ö���ʽ��Ϣ�б�
'ClassId --- ����
'CommandType --- �Ƽ�����
'iCols--ÿ����ʾ����
'iLength--ÿ����Ϣ��ʾ�೤
'IsNew  ---- �Ƿ���ʾ������Ϣ
'������
'2005-12-21
'****************************************
Private Function getInfoList(PageType, ClassID, CommandType, iCols, iLength, IsNew, IsHot, iWeight, iHeight, ShowNum, SupplyType, ShowInfoType)
    Dim strTable, strTemp, Rows, rsSupply, strSql
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    strSql = GetExecuteSql(PageType, ClassID, CommandType, IsNew, IsHot, ShowNum, SupplyType, False)
    rsSupply.Open strSql, Conn, 1, 1
    Rows = 0
    If rsSupply.EOF And rsSupply.BOF Then
        getInfoList = "<li>û���κ���Ϣ!</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            Rows = Rows + 1
            strTemp = strTemp & "<td><Img src='" & ChannelUrl & "/Images/article_common.gif' border=0/>"
            If ShowInfoType Then
                strTemp = strTemp & "<a href='" & ChannelUrl & "/Search.asp?SupplyType=" & rsSupply("SupplyType") & "&SearchType=0'><font color=red>[" & GetSupplyInfoType(rsSupply("SupplyType"), "//SupplyType/Type") & "]</font></a>"
            End If
            strTemp = strTemp & "<a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target ='_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & ">"
            strTemp = strTemp & Left(rsSupply("SupplyTitle"), iLength) & "</a></td>" & vbCrLf
            If (Rows Mod iCols) = 0 Then
                strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
                strTemp = ""
            End If
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
    getInfoList = "<Table Width='100%'>" & strTable & "</Table>"
End Function

Private Function getHotSupplyInfo(ByVal iCols, ByVal iLength, ByVal ShowNum)
    Dim strSql, strTable, strTemp, Rows, rsSupply
    strSql = "Select Top " & ShowNum & " SupplyID,SupplyTitle From PE_Supply Where Hits >= " & HitsOfHot & " And Deleted=" & PE_False & " And Status=1 Order By SupplyID DESC"
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    Rows = 0
    If rsSupply.EOF And rsSupply.BOF Then
        getHotSupplyInfo = "<li>û���ȵ���Ϣ!</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            Rows = Rows + 1
            strTemp = strTemp & "<td><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If

            strTemp = strTemp & ">"
            strTemp = strTemp & Left(rsSupply("SupplyTitle"), iLength) & "</a></td>" & vbCrLf
            If (Rows Mod iCols) = 0 Then
                strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
                strTemp = ""
            End If
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
    getHotSupplyInfo = "<Table Width='100%'>" & strTable & "</Table>"
End Function

Private Function getCommandSupplyInfo(ByVal CommandType, ByVal iCols, ByVal iLength, ByVal ShowNum)
    Dim strSql, strTable, strTemp, Rows, rsSupply, QuerySql
    Select Case CommandType
        Case 1
            QuerySql = "CommandType=" & CommandType & " And DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ") < CommandChannelDays "
        Case 2
            QuerySql = "CommandType=" & CommandType & " And DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ") < CommandClassDays "
        Case Else
            QuerySql = "CommandType<>0"
    End Select
    strSql = "Select Top " & ShowNum & " SupplyId,SupplyTitle From PE_Supply Where " & QuerySql & " And Deleted=" & PE_False & " And Status=1 Order By SupplyId DESC"
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    Rows = 0
    If rsSupply.EOF And rsSupply.BOF Then
        getCommandSupplyInfo = "<li>û���Ƽ���Ϣ!</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            Rows = Rows + 1
            strTemp = strTemp & "<td><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target ='_blank' "
            End If
            strTemp = strTemp & ">"
            strTemp = strTemp & Left(rsSupply("SupplyTitle"), iLength) & "</a></td>" & vbCrLf
            If (Rows Mod iCols) = 0 Then
                strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
                strTemp = ""
            End If
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
    getCommandSupplyInfo = "<Table Width='100%'>" & strTable & "</Table>"
End Function

Private Function getLasterSupplyInfo(ByVal iCols, ByVal iLength, ByVal ShowNum)
    Dim strSql, strTable, strTemp, Rows, rsSupply
    strSql = "Select Top " & ShowNum & " SupplyId,SupplyTitle From PE_Supply Where DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)< " & DaysOfNew & " And Deleted=" & PE_False & " And Status=1 Order by SupplyId DESC"
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    Rows = 0
    If rsSupply.EOF And rsSupply.BOF Then
        getLasterSupplyInfo = "<li>û��������Ϣ!</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            Rows = Rows + 1
            strTemp = strTemp & "<td Class='LasterStyle'><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & ">"
            strTemp = strTemp & Left(rsSupply("SupplyTitle"), iLength) & "</a></td>" & vbCrLf
            If (Rows Mod iCols) = 0 Then
                strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
                strTemp = ""
            End If
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
    getLasterSupplyInfo = "<Table Width='100%'>" & strTable & "</Table>"
End Function
Private Function getPicLasterInfo(ByVal iCols, ByVal iLength, ByVal ShowNum, ByVal iWidth, ByVal iHeight)
    Dim strSql, rsSupply, Rows, strTemp, strTable, QuerySql
    If ClassID > 0 Then
        If Child > 0 Then
            QuerySql = " And ClassID in (" & arrChildID & ")"
        Else
            QuerySql = " And ClassID=" & ClassID
        End If
    ElseIf SpecialID > 0 Then
        QuerySql = "And SpecialId=" & SpecialID
    End If

    strSql = "Select Top " & ShowNum & " SupplyID,SupplyTitle,SupplyPicUrl From PE_Supply Where SupplyPicUrl<>'' And Deleted=" & PE_False & " And Status=1 And DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)< " & DaysOfNew & " " & QuerySql & " Order By SupplyID"
   
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    Rows = 0
    If rsSupply.EOF And rsSupply.BOF Then
        getPicLasterInfo = "<li>û������ͼƬ��Ϣ!</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            Rows = Rows + 1
            strTemp = strTemp & "<Td><Table><tr><td><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & ">"
            strTemp = strTemp & "<Img src=" & UploadDir & "/" & getDefaultPicUrl(rsSupply("SupplyPicUrl")) & " border='0 'width ='" & iWidth & "' height='" & iHeight & "' alt=" & rsSupply("SupplyTitle") & " /></a></td></tr><tr><td><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'>" & Left(rsSupply("SupplyTitle"), iLength) & "</a></td></tr></Table></td>"
            If (Rows Mod iCols) = 0 Then
                strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
                strTemp = ""
            End If
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
    getPicLasterInfo = "<Table Width='100%'>" & strTable & "</Table>"
End Function
'*******************************************
'���ͼƬ��Ϣ�б���ʽ
'ShowNum -----------ÿҳ��ʾ����Ϣ��
'CurrentPage -------��ǰҳ��
'KeyWords ----------�ؼ���
'Flag --------------�Ƿ��з�ҳ
'iLength -----------���ⳤ��
'iWidth  -----------ͼƬ�Ŀ��
'iHeight -----------ͼƬ�ĸ߶�
'******************************************
Private Function getListPicInfoList(ByVal PageType, ByVal ClassID, ByVal CommandType, ByVal iCols, ByVal iLength, ByVal IsNew, ByVal IsHot, ByVal iWidth, ByVal iHeight, ByVal ShowNum, ByVal SupplyType, ByVal ShowInfoType)
    Dim strSql, rsSupply, Rows, strTemp
    
    strSql = GetExecuteSql(PageType, ClassID, CommandType, IsNew, IsHot, ShowNum, SupplyType, True)
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    If rsSupply.EOF And rsSupply.BOF Then
        getListPicInfoList = "<li>û��ͼƬ��Ϣ!</li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            strTemp = strTemp & "<tr><td><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & "><img width =" & iWidth & " height=" & iHeight & " src=" & UploadDir & "/" & getDefaultPicUrl(rsSupply("SupplyPicUrl")) & " border=0 /></a></td><td>"
            If ShowInfoType Then
                strTemp = strTemp & "<a href='" & ChannelUrl & "/Search.asp?SupplyType=" & rsSupply("SupplyType") & "&SearchType=0'><font color=red>[" & GetSupplyInfoType(rsSupply("SupplyType"), "//SupplyType/Type") & "]</font></a>"
            End If
            strTemp = strTemp & "<a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'>" & Left(rsSupply("SupplyTitle"), iLength) & "</a></td><td>" & rsSupply("Province") & "/" & rsSupply("City") & "</td><td>" & GetSupplyInfoType(rsSupply("TradeType"), "//TradeType/Type") & "</td></tr>"
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
   
    getListPicInfoList = "<Table Width='100%'>" & strTemp & "</Table>"
End Function
Private Function getDefaultPicUrl(ByVal PicUrl)
    Dim arrPicUrl
    If Not (IsNull(PicUrl)) Or PicUrl <> "" Then
        arrPicUrl = Split(PicUrl, "|")
        getDefaultPicUrl = arrPicUrl(0)
    End If
End Function
'���ͼƬ�б���ʽһ
Private Function getPicInfoList(ByVal PageType, ByVal ClassID, ByVal CommandType, ByVal iCols, ByVal iLength, ByVal IsNew, ByVal IsHot, ByVal iWidth, ByVal iHeight, ByVal ShowNum, ByVal SupplyType, ByVal ShowInfoType)
    Dim strSql, rsSupply, Rows, strTemp, strTable
    
   strSql = GetExecuteSql(PageType, ClassID, CommandType, IsNew, IsHot, ShowNum, SupplyType, True)

    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    Rows = 0
    If rsSupply.EOF And rsSupply.BOF Then
        getPicInfoList = "<li>û��ͼƬ��Ϣ!<li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            Rows = Rows + 1
            strTemp = strTemp & "<Td><Table><tr><td><a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'"
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & "><Img src=" & UploadDir & "/" & getDefaultPicUrl(rsSupply("SupplyPicUrl")) & " border='0 'width ='" & iWidth & "' height='" & iHeight & "' alt=" & rsSupply("SupplyTitle") & " /></a></td></tr><tr><td>"
            If ShowInfoType Then
                strTemp = strTemp & "<a href='" & ChannelUrl & "/Search.asp?SupplyType=" & rsSupply("SupplyType") & "&SearchType=0'><font color=red>[" & GetSupplyInfoType(rsSupply("SupplyType"), "//SupplyType/Type") & "]</font></a>"
            End If
            strTemp = strTemp & "<a href='" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & "'>" & Left(rsSupply("SupplyTitle"), iLength) & "</a></td></tr></Table></td>"
            If (Rows Mod iCols) = 0 Then
                strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
                strTemp = ""
            End If
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    strTable = strTable & "<tr>" & strTemp & "</tr>" & vbCrLf
    getPicInfoList = "<Table Width='100%'>" & strTable & "</Table>"
End Function


'�����Ϣ�б�
'2005-11-21
'������
Private Function getDetailInfoList(PageType, ClassID, CommandType, iCols, iLength, IsNew, IsHot, iWeight, iHeight, ShowNum, SupplyType, ShowInfoType)
    Dim strSql, strTable, strTemp, Rows, rsSupply
 
    strSql = GetExecuteSql(PageType, ClassID, CommandType, IsNew, IsHot, ShowNum, SupplyType, False)
    
    Set rsSupply = Server.CreateObject("ADODB.RecordSet")
    rsSupply.Open strSql, Conn, 1, 1
    If rsSupply.EOF And rsSupply.BOF Then
        getDetailInfoList = "<li>û����Ϣ!<li>"
        rsSupply.Close
        Set rsSupply = Nothing
        Exit Function
    Else
        Do While Not rsSupply.EOF
            
            strTemp = strTemp & "<tr><td><Img src='" & ChannelUrl & "/Images/article_common.gif' border=0/>"
            If ShowInfoType Then
                strTemp = strTemp & "<a href='" & ChannelUrl & "/Search.asp?SupplyType=" & rsSupply("SupplyType") & "&SearchType=0'><font color=red>[" & GetSupplyInfoType(rsSupply("SupplyType"), "//SupplyType/Type") & "]</font></a>"
            End If
            strTemp = strTemp & "<a href=" & ChannelUrl & "/ShowSupply.asp?SupplyId=" & rsSupply("SupplyId") & ""
            If OpenType = 0 Then
                strTemp = strTemp & " target = '_self' "
            Else
                strTemp = strTemp & " target = '_blank' "
            End If
            strTemp = strTemp & ">" & Left(rsSupply("SupplyTitle"), iLength) & "</a></td><td align='center'>" & GetSupplyInfoType(rsSupply("TradeType"), "//TradeType/Type") & "</td><td>" & rsSupply("Province") & "/" & rsSupply("City") & "</td><td>" & rsSupply("UpDateTime") & "</td></tr>" & vbCrLf
            rsSupply.MoveNext
        Loop
    End If
    rsSupply.Close
    Set rsSupply = Nothing
    getDetailInfoList = "<Table Width='100%'>" & strTemp & "</Table>"
End Function

'���ĳһ����µ�������Ϣ��
'2005-11-18
'������
Private Function getInfoCounts(ByVal ClassID, ByVal iType, ByVal SupplyType)
    'Call OpenConn()
    Dim strSql, QuerySql
    If SupplyType >= 0 Then
        QuerySql = " And SupplyType=" & SupplyType & " "
    End If
    Select Case iType
        Case 0
            strSql = "Select Count(*) From PE_Supply Where Status=1 And Deleted=" & PE_False & " And ClassId in (" & ClassID & ")"
        Case 1 '����ҳ������ͳ��
            strSql = "Select Count(*) From PE_Supply Where Status=1 And Deleted=" & PE_False & " And DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)<" & DaysOfNew & ""
        Case 2 '�Ƽ�ҳ������ͳ��
            strSql = "Select Count(*) From PE_Supply Where Status=1 And Deleted=" & PE_False & " And CommandType<>0"
        Case 3
            strSql = "Select Count(*) From PE_Supply Where Status=1 And Deleted=" & PE_False & " And Hits>=" & HitsOfHot & ""
        Case 4
            strSql = "Select Count(*) From PE_Supply Where Status=1 And Deleted=" & PE_False & " And SpecialId=" & SpecialID & ""
        Case 5
            strSql = "Select Count(*) From PE_Supply Where Status=1 And Deleted=" & PE_False & " And CommandType<>0 And SupplyPicUrl<>''"
    End Select
    getInfoCounts = PE_CLng(Conn.Execute(strSql & QuerySql)(0))
End Function

'=================================================
'��������ShowChannelCount
'��  �ã���ʾƵ��ͳ����Ϣ
'��  ������
'=================================================
Private Function GetChannelCount()
    Dim HitsCount_Channel, rs
    Set rs = Conn.Execute("select sum(Hits) from PE_Supply where ChannelID=" & ChannelID)
    If IsNull(rs(0)) Then
        HitsCount_Channel = 0
    Else
        HitsCount_Channel = rs(0)
    End If
    rs.Close
    Set rs = Nothing
    GetChannelCount = Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ChannelCount", "{$ChannelShortName}������ {$ItemChecked_Channel} {$ChannelItemUnit}<br>����{$ChannelShortName}�� {$UnItemChecked} {$ChannelItemUnit}<br>���������� {$CommentCount_Channel} ��<br>ר�������� {$SpecialCount_Channel} ��<br>{$ChannelShortName}�Ķ��� {$HitsCount_Channel} �˴�<br>"), "{$ItemChecked_Channel}", ItemChecked_Channel), "{$ChannelItemUnit}", ChannelItemUnit), "{$UnItemChecked}", ItemCount_Channel - ItemChecked_Channel), "{$CommentCount_Channel}", CommentCount_Channel), "{$SpecialCount_Channel}", SpecialCount_Channel), "{$HitsCount_Channel}", HitsCount_Channel)
End Function


Private Function GetSupplyAction()
    GetSupplyAction = Replace(Replace(Replace(Replace(R_XmlText_Class("SupplyAction", "��<a href='{$ChannelUrl}/Comment.asp?SupplyID={$SupplyID}' target='_blank'>��������</a>����<a href='{$InstallDir}User/User_Favorite.asp?Action=Add&ChannelID={$ChannelID}&InfoID={$SupplyID}' target='_blank'>�����ղ�</a>����<a href='javascript:window.close();'>�رմ���</a>��"), "{$ChannelUrl}", ChannelUrl), "{$SupplyID}", SupplyID), "{$InstallDir}", strInstallDir), "{$ChannelID}", ChannelID)
End Function


Private Function GetSupplyUrl(ByVal SupplyID)
    GetSupplyUrl = strInstallDir & ChannelDir & "/ShowSupply.asp?SupplyID=" & SupplyID
End Function


Private Sub GetRegionValue()
    Response.Write "<script language='javascript'> " & vbCrLf
    Response.Write "getSelected('Region.asp',-1);" & vbCrLf
    Response.Write " var http_request = false; " & vbCrLf
    Response.Write " function InitRequest() {//��ʼ����ָ������������������ĺ��� " & vbCrLf
    Response.Write "     http_request = false; " & vbCrLf
    Response.Write "     //��ʼ��ʼ��XMLHttpRequest���� " & vbCrLf
    Response.Write "     if(window.XMLHttpRequest) { //Mozilla ����� " & vbCrLf
    Response.Write "         http_request = new XMLHttpRequest(); " & vbCrLf
    Response.Write "         if (http_request.overrideMimeType) {//����MiME��� " & vbCrLf
    Response.Write "             http_request.overrideMimeType('text/xml'); " & vbCrLf
    Response.Write "         } " & vbCrLf
    Response.Write "     } " & vbCrLf
    Response.Write "     else if (window.ActiveXObject) { // IE����� " & vbCrLf
    Response.Write "         try { " & vbCrLf
    Response.Write "             http_request = new ActiveXObject('Msxml2.XMLHTTP'); " & vbCrLf
    Response.Write "         } catch (e) { " & vbCrLf
    Response.Write "             try { " & vbCrLf
    Response.Write "                 http_request = new ActiveXObject('Microsoft.XMLHTTP'); " & vbCrLf
    Response.Write "             } catch (e) {} " & vbCrLf
    Response.Write "         } " & vbCrLf
    Response.Write "     } " & vbCrLf
    Response.Write "     if (!http_request) { // �쳣����������ʵ��ʧ�� " & vbCrLf
    Response.Write "         window.alert('���ܴ���XMLHttpRequest����ʵ��.'); " & vbCrLf
    Response.Write "         return false; " & vbCrLf
    Response.Write "     } " & vbCrLf
    Response.Write "      " & vbCrLf
    Response.Write " } " & vbCrLf
    Response.Write " //�趨��ʼֵ " & vbCrLf
    Response.Write " function getSelectValue(url,SelectName) " & vbCrLf
    Response.Write " { " & vbCrLf
    Response.Write "     InitRequest(); " & vbCrLf
    Response.Write "     http_request.onreadystatechange = function() " & vbCrLf
    Response.Write "     { " & vbCrLf
    Response.Write "         if (http_request.readyState == 4)  " & vbCrLf
    Response.Write "         { // �ж϶���״̬ " & vbCrLf
    Response.Write "             if (http_request.status == 200)  " & vbCrLf
    Response.Write "             { // ��Ϣ�Ѿ��ɹ����أ���ʼ������Ϣ " & vbCrLf
    Response.Write "                //alert(unescape(http_request.responseText));" & vbCrLf
    Response.Write "                getClass(unescape(http_request.responseText),SelectName); " & vbCrLf
    Response.Write "             } else { //ҳ�治���� " & vbCrLf
    Response.Write "                 alert('���������ҳ�����쳣��'); " & vbCrLf
    Response.Write "             } " & vbCrLf
    Response.Write "         } " & vbCrLf
    Response.Write "     }        " & vbCrLf
    Response.Write "     // ȷ����������ķ�ʽ��URL�Լ��Ƿ�ͬ��ִ���¶δ��� " & vbCrLf
    Response.Write "     http_request.open('GET',url, false); " & vbCrLf
    Response.Write "     http_request.send(null); " & vbCrLf
    Response.Write " } " & vbCrLf
    Response.Write " function getClass(node,SelectName) " & vbCrLf
    Response.Write " { " & vbCrLf
    Response.Write "     SelectName.options.length =1 ;"
    Response.Write "     var arrstr = new Array(); " & vbCrLf
    Response.Write "     arrstr = node.split(','); " & vbCrLf
    Response.Write "     for(var i=0;i<arrstr.length-1;i++) " & vbCrLf
    Response.Write "     { " & vbCrLf
    Response.Write "         SelectName.options[i+1] =new Option(arrstr[i]); " & vbCrLf
    Response.Write "         SelectName.options[i+1].value = arrstr[i]; " & vbCrLf
    Response.Write "     } " & vbCrLf
    Response.Write " } " & vbCrLf
    Response.Write " function getSelected(url,selValue)" & vbCrLf
    Response.Write " {" & vbCrLf
    Response.Write "    if(selValue==-1)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        getSelectValue(url,document.Searchform.mySelectProvince);" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        getSelectValue(url,selValue);" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "</script> " & vbCrLf
End Sub

Public Sub ShowFavorite()
    Response.Write "<table width='100%' cellpadding='2' cellspacing='1' border='0' class='border'>"
    Response.Write "  <tr class='title' align='center'><td width='30'>ѡ��</td><td>" & ChannelShortName & "����</td><td width='100'>������</td><td width='80'>����ʱ��</td><td width='80'>����</td></tr>"
    
    Dim sqlFavorite, rsFavorite, iCount, strLink
    iCount = 0
    
    sqlFavorite = "select A.ChannelID,A.SupplyID,A.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,A.SupplyTitle,A.UserName,A.UpdateTime,A.SupplyPicUrl from PE_Supply A left join PE_Class C on A.ClassID=C.ClassID where A.Deleted=" & PE_False & " and A.Status=1 "
    sqlFavorite = sqlFavorite & " and SupplyID in (select InfoID from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & ")"
    sqlFavorite = sqlFavorite & " order by A.SupplyID desc"
    MaxPerPage = 20
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
            strLink = "[<a href='" & GetClassUrl(rsFavorite("ClassID")) & "'>" & rsFavorite("ClassName") & "</a>] "
            strLink = strLink & "<a href='" & GetSupplyUrl(rsFavorite("SupplyID")) & "' target='_blank'>" & rsFavorite("SupplyTitle") & "</a>"
            Response.Write "<tr class='tdbg'>"
            Response.Write "<td align='center' width='30'><input type='checkbox' name='InfoID' value='" & rsFavorite("SupplyID") & "'></td>"
            Response.Write "<td align='left'>" & strLink & "</td>"
            Response.Write "<td width='100' align='center'>" & rsFavorite("UserName") & "</td>"
            Response.Write "<td width='80' align='right'>" & Year(rsFavorite("UpdateTime")) & "-" & Right("0" & Month(rsFavorite("UpdateTime")), 2) & "-" & Right("0" & Day(rsFavorite("UpdateTime")), 2) & "</td>"
            Response.Write "<td width='80' align='center'><a href='User_Favorite.asp?Action=Remove&ChannelID=" & ChannelID & "&InfoID=" & rsFavorite("SupplyID") & "' onclick=""return confirm('ȷʵ�����ղش�" & ChannelShortName & "��');"">ȡ���ղ�</a></td>"
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
    XmlText_Class = XmlText("Supply", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Supply", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>
