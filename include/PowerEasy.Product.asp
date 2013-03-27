<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim ProductID, ProductName, ProductUrl
Dim rsProduct

Class Product

Private ChannelID
Private strPrice_Market, strPrice_Shop, strPrice_Member, strPrice_Original, strPrice_Te, strPrice_Time, strPrice_Now

'��������ȫ�ֵı���
Private rsClass, NoPrice, NoPrice_Member, NoPrice_Market

Public Sub Init()
    FoundErr = False
    ErrMsg = ""
    ChannelID = 1000
    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
    '*****************************
    '��ȡ���԰��е��ַ�����
    ChannelShortName = XmlText_Class("ChannelShortName", "��Ʒ")
    strListStr_Font = XmlText_Class("ProductList/UpdateTimeColor_New", "color=""red""")
    strTop = XmlText_Class("ProductList/t4", "�̶�")
    strElite = XmlText_Class("ProductList/t3", "�Ƽ�")
    strCommon = XmlText_Class("ProductList/t5", "��ͨ")
    strHot = XmlText_Class("ProductList/t7", "�ȵ�")
    strNew = XmlText_Class("ProductList/t6", "����")
    strTop2 = XmlText_Class("ProductList/Top", " ��")
    strElite2 = XmlText_Class("ProductList/Elite", " ��")
    strHot2 = XmlText_Class("ProductList/Hot", " ��")
    Character_Author = XmlText("Product", "Include/Author", "[{$Text}]")
    Character_Date = XmlText("Product", "Include/Date", "[{$Text}]")
    Character_Hits = XmlText("Product", "Include/Hits", "[{$Text}]")
    Character_Class = XmlText("Product", "Include/ClassChar", "[{$Text}]")
    SearchResult_Content_NoPurview = XmlText("BaseText", "SearchPurviewContent", "��������Ҫ��ָ��Ȩ�޲ſ���Ԥ��")
    SearchResult_ContentLenth = PE_CLng(XmlText_Class("ShowSearch/Content_Lenght", "200"))
    strList_Content_Div = XmlText_Class("ProductList/Content_DIV", "style=""padding:0px 20px""")
    strList_Title = R_XmlText_Class("ProductList/Title", "{$ChannelShortName}���⣺{$Title}{$br}��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�{$Author}{$br}����ʱ�䣺{$UpdateTime}")
    strComment = XmlText_Class("ProductList/CommentLink", "<font color=""red"">����</font>")
    
    NoPrice = XmlText_Class("NoPrice", "0")
    NoPrice_Member = XmlText_Class("NoPrice_Member", "��")
    NoPrice_Market = XmlText_Class("NoPrice_Market", "��")
    strPrice_Market = XmlText_Class("ProductPrice/Price_Market", "�г��ۣ�")
    strPrice_Shop = XmlText_Class("ProductPrice/Price", "�̳Ǽۣ�")
    strPrice_Member = XmlText_Class("ProductPrice/Price_Member", "��Ա�ۣ�")
    strPrice_Original = XmlText_Class("ProductPrice/Price_Original", "ԭ&nbsp;&nbsp;�ۣ�")
    strPrice_Te = XmlText_Class("ProductPrice/Price_Te", "��&nbsp;&nbsp;�ۣ�")
    strPrice_Time = XmlText_Class("ProductPrice/Price_Time", "ʱ&nbsp;&nbsp;�䣺")
    strPrice_Now = XmlText_Class("ProductPrice/Price_Now", "��&nbsp;&nbsp;�ۣ�")
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
    Dim strCount
    strCount = Replace(Replace(Replace(Replace(R_XmlText_Class("ChannelCount", "{$ChannelShortName}������ {$ItemChecked_Channel} {$ChannelItemUnit}<br>���������� {$CommentCount_Channel} ��<br>ר�������� {$SpecialCount_Channel} ��<br>"), "{$ItemChecked_Channel}", ItemChecked_Channel), "{$ChannelItemUnit}", ChannelItemUnit), "{$CommentCount_Channel}", CommentCount_Channel), "{$SpecialCount_Channel}", SpecialCount_Channel)
    GetChannelCount = strCount
End Function

Private Function GetSqlStr(arrClassID, IncludeChild, iSpecialID, ProductType, IsHot, IsElite, DateNum, OrderType, ShowClassName, IsPicUrl)
    Dim strSql, IDOrder
    iSpecialID = PE_CLng(iSpecialID)
    If IsValidID(arrClassID) = False Then
        arrClassID = 0
    Else
        arrClassID = ReplaceLabelBadChar(arrClassID)
    End If	
    If UseCreateHTML > 0 Or ShowClassName = True Then
        strSql = ",C.ClassName,C.ParentDir,C.ClassDir,C.Readme,C.ClassPurview"
        If iSpecialID > 0 Then
            strSql = strSql & " from PE_InfoS I inner join (PE_Product P left join PE_Class C on P.ClassID=C.ClassID) on I.ItemID=P.ProductID"
        Else
            strSql = strSql & " from PE_Product P left join PE_Class C on P.ClassID=C.ClassID"
        End If
    Else
        If iSpecialID > 0 Then
            strSql = " from PE_InfoS I inner join PE_Product P on I.ItemID=P.ProductID"
        Else
            strSql = " from PE_Product P"
        End If
    End If
    strSql = strSql & " where P.Deleted=" & PE_False & " and P.EnableSale=" & PE_True

    If InStr(arrClassID, ",") > 0 Then
        strSql = strSql & " and P.ClassID in (" & FilterArrNull(arrClassID, ",") & ") "
    Else
        arrClassID = PE_CLng(arrClassID)
        If arrClassID > 0 Then
            If IncludeChild = True Then
                Dim trs
                Set trs = Conn.Execute("select arrChildID from PE_Class where ChannelID=" & ChannelID & " and ClassID=" & arrClassID & "")
                If trs.BOF And trs.EOF Then
                    arrClassID = 0
                Else
                    If IsNull(trs(0)) Or Trim(trs(0)) = "" Then
                        arrClassID = 0
                    Else
                        arrClassID = trs(0)
                    End If
                End If
                Set trs = Nothing
            End If
            If InStr(arrClassID, ",") > 0 Then
                strSql = strSql & " and P.ClassID in (" & arrClassID & ") "
            Else
                If PE_CLng(arrClassID) > 0 Then strSql = strSql & " and P.ClassID=" & PE_CLng(arrClassID)
            End If
        End If
    End If
    If iSpecialID > 0 Then
        strSql = strSql & " and I.ModuleType=5 and I.SpecialID=" & iSpecialID
    End If
    Select Case PE_CLng(ProductType)
    Case 1   '��������
        strSql = strSql & " and (P.ProductType=1 or (P.ProductType=3 and (BeginDate>" & PE_Now & " or EndDate<" & PE_Now & ")))"
    Case 2   '�Ǽ���Ʒ
        strSql = strSql & " and P.ProductType=2"
    Case 3   '�ؼ۴�����Ʒ
        strSql = strSql & " and (P.ProductType=3 and BeginDate<=" & PE_Now & " and EndDate>=" & PE_Now & ")"
    Case 4 '�������Ʒ
        strSql = strSql & " and P.ProductType=4"
    Case 5    '�������ۺ��Ǽ���Ʒ
        strSql = strSql & " and P.ProductType<3"
    Case 6    '���۴�����Ʒ
        strSql = strSql & " and P.ProductType=5"
    Case Else
        '��ָ����Ʒ����
    End Select
    If IsHot = True Then
        strSql = strSql & " and P.IsHot=" & PE_True & ""
    End If
    If IsElite = True Then
        strSql = strSql & " And P.IsElite=" & PE_True & ""
    End If
    If DateNum > 0 Then
        strSql = strSql & " and DateDiff(" & PE_DatePart_D & ",P.UpdateTime," & PE_Now & ")<" & DateNum
    End If

    If IsPicUrl = True Then
        strSql = strSql & " and P.ProductThumb<>'' "
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
        strSql = strSql & "P.ProductID " & IDOrder
    End If
    GetSqlStr = strSql
End Function

'=================================================
'��������GetProductList
'��  �ã���ʾ��Ʒ���Ƶ���Ϣ
'��  ����
'0        arrClassID ---��ĿID���飬0Ϊ������Ŀ
'1        IncludeChild ----�Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'2        iSpecialID ------ר��ID��0Ϊ���в�Ʒ������ר���Ʒ�������Ϊ����0����ֻ��ʾ��Ӧר��Ĳ�Ʒ
'3        ProductNum ---��Ʒ����������0����ֻ��ѯǰ������Ʒ
'4        ProductType ---- ��Ʒ���ͣ�1Ϊ����������Ʒ��2Ϊ�Ǽ���Ʒ��3Ϊ������Ʒ��4Ϊ������Ʒ��5Ϊ�������ۺ��Ǽ���Ʒ��0Ϊ������Ʒ
'5        IsHot ------------�Ƿ������Ų�Ʒ��TrueΪֻ��ʾ���Ų�Ʒ��FalseΪ��ʾ���в�Ʒ
'6        IsElite ----------�Ƿ����Ƽ���Ʒ��TrueΪֻ��ʾ�Ƽ���Ʒ��FalseΪ��ʾ���в�Ʒ
'7        DateNum ----���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µĲ�Ʒ
'8        OrderType ----����ʽ��1--����ƷID����2--����ƷID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'9        ShowType -----��ʾ��ʽ��1Ϊ��ͨ��ʽ��2Ϊ���ʽ��3Ϊ�������ʽ 4Ϊ���DIV��ʽ 5Ϊ���RSS��ʽ
'10       TitleLen  ----��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ��������
'11       ContentLen ---��Ʒ�������ַ�����һ������=����Ӣ���ַ���Ϊ0ʱ����ʾ��
'12       ShowClassName -----�Ƿ���ʾ������Ŀ���ƣ�TrueΪ��ʾ��FalseΪ����ʾ
'13       ShowPropertyType ------��ʾ��Ʒ���ԣ��̶�/�Ƽ�/��ͨ���ķ�ʽ��0Ϊ����ʾ��1ΪСͼƬ��2Ϊ����
'14       ShowDateType ------��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���գ�3Ϊ�ԡ���-�ա���ʽ��ʾ���ա�
'15       ShowHotSign -----------�Ƿ���ʾ���Ų�Ʒ��־��TrueΪ��ʾ��FalseΪ����ʾ
'16       ShowNewSign -------�Ƿ���ʾ�²�Ʒ��־��TrueΪ��ʾ��FalseΪ����ʾ
'17       UsePage -----------�Ƿ��ҳ��ʾ��TrueΪ��ҳ��ʾ��FalseΪ����ҳ��ʾ��ÿҳ��ʾ�Ĳ�Ʒ������MaxPerPageָ��
'18       OpenType -----��Ʒ�򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�

'19       UrlType ---- ���ӵ�ַ���ͣ�0Ϊ���·����1Ϊ����ַ�ľ���·����ע��˲����ڱ�ǩ�в�����

'20       IntervalLines ---- ÿ��N�пհ�һ�У�Ϊ0ʱ������
'21       Cols ----ÿ�е������������������ͻ��С�

'���²���ֻ�е�ShowType������Ϊ���ʽʱ����Ч
'22       ShowTableTitle ---- �Ƿ���ʾ���ͷ����ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'23       TableTitleStr ---- ���ͷ�����֣����磺��Ʒ����|�ͺ�|���|����ʱ��|��λ|�����|����|�г���|�̳Ǽ�|�Żݼ�|��Ա��|�ۿ���|������������ɾ����Ŀ��ֻ�е�ShowType������Ϊ���ʽʱ����Ч�����Ϊ�գ�����ʾ���ͷ��
'24       ShowProductModel ---- �Ƿ���ʾ��Ʒ�ͺţ�ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'25       ShowProductStandard ---- �Ƿ���ʾ��Ʒ���ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'26       ShowUnit ---- �Ƿ���ʾ��Ʒ��λ��ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'27       ShowStocksType ----��ʾ��Ʒ��淽ʽ��ֻ�е�ShowType������Ϊ���ʽʱ����Ч��0Ϊ����ʾ��1Ϊ��ʾ�����棬2Ϊ��ʾʵ�ʿ��
'28       ShowWeight ---- �Ƿ���ʾ��Ʒ������ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'29       ShowPrice_Market ---- �Ƿ���ʾ�г��ۣ�ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'30       ShowPrice_Original ---- �Ƿ���ʾԭ�ۣ�ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'31       ShowPrice ---- �Ƿ���ʾ��ǰ���ۼۣ�ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'32       ShowPrice_Member ---- �Ƿ���ʾ��Ա�ۣ�ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'33       ShowDiscount  ---- �Ƿ���ʾ�ۿ��ʣ�ֻ�е�ShowType������Ϊ���ʽʱ����Ч��TrueΪ��ʾ��FalseΪ����ʾ
'34       ShowButtonType ---- ��ť��ʾ��ʽ��0Ϊ����ʾ��1Ϊ��ʾ����ť��2Ϊ��ʾ��ϸ��ť��3Ϊ�ղذ�ť��4Ϊ��ʾ������ϸ��ť��5Ϊ��ʾ�����ղذ�ť��6Ϊ��ϸ���ղذ�ť��7Ϊ��������ʾ
'35       ButtonStyle ---- ��ť��ʽ

'36       CssNameTable ---- ����CSS��������ѡ����
'37       CssNameTitle ---- ���ͷ���е�CSS��������ѡ����
'38       CssNameA ---- �б����������ӵ��õ�CSS��������ѡ����
'39       CssName1 ---- �б��������е�CSSЧ������������ѡ����
'40       CssName2 ---- �б���ż���е�CSSЧ������������ѡ����
'=================================================

Public Function GetProductList(arrClassID, IncludeChild, iSpecialID, ProductNum, ProductType, IsHot, IsElite, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowDateType, ShowHotSign, ShowNewSign, UsePage, OpenType, UrlType, IntervalLines, Cols, ShowTableTitle, TableTitleStr, ShowProductModel, ShowProductStandard, ShowUnit, ShowStocksType, ShowWeight, ShowPrice_Market, ShowPrice_Original, ShowPrice, ShowPrice_Member, ShowDiscount, ShowButtonType, ButtonStyle, CssNameTable, CssNameTitle, CssNameA, CssName1, CssName2)
    Dim sqlProduct, rsProductList, iCount, iNumber, strProductList, TitleStr, strThisClass, trs, strProperty, strLink
    Dim CssName, iAuthor
    Dim arrTitle, iLines
    iCount = 0
    strThisClass = ""
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)
    If Cols <= 0 Then Cols = 1
    
    IntervalLines = PE_CLng(IntervalLines)
    
    If TitleLen < 0 Or TitleLen > 200 Then
        TitleLen = 50
    End If

    If IsNull(CssNameTable) Or CssNameTable = "" Then CssNameA = "productlist_table"
    If IsNull(CssNameTitle) Or CssNameTitle = "" Then CssNameA = "productlist_title"
    If IsNull(CssNameA) Or CssNameA = "" Then CssNameA = "productlist_A"
    If IsNull(CssName1) Or CssName1 = "" Then CssName1 = "productlist_tr1"
    If IsNull(CssName2) Or CssName2 = "" Then CssName2 = "productlist_tr2"
    CssName = CssName1

    If TableTitleStr = "" Then
        TableTitleStr = "��Ʒ����|�ͺ�|���|����ʱ��|��λ|�����|����|�г���|�̳Ǽ�|�Żݼ�|��Ա��|�ۿ���|����"
    End If
    arrTitle = Split(TableTitleStr, "|")
    If UBound(arrTitle) <> 12 Then
        arrTitle = Split("��Ʒ����|�ͺ�|���|����ʱ��|��λ|�����|����|�г���|�̳Ǽ�|�Żݼ�|��Ա��|�ۿ���|����", "|")
    End If

    If ProductNum > 0 Then
        sqlProduct = "select top " & ProductNum & " "
    Else
        sqlProduct = "select "
    End If
    If ContentLen > 0 Then
        sqlProduct = sqlProduct & "P.ProductExplain,"
    End If
    sqlProduct = sqlProduct & "P.ClassID,P.ProductID,P.ProductNum,P.ProductName,P.UpdateTime,P.ProductThumb,P.ProductIntro,P.Hits"
    sqlProduct = sqlProduct & ",P.IsHot,P.IsElite,P.OnTop,P.ProductModel,P.ProductStandard,P.ProducerName,P.Unit,P.Stocks,P.OrderNum,P.Stars"
    sqlProduct = sqlProduct & ",P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,P.BeginDate,P.EndDate,P.Discount,P.Weight"
    sqlProduct = sqlProduct & GetSqlStr(arrClassID, IncludeChild, iSpecialID, ProductType, IsHot, IsElite, DateNum, OrderType, ShowClassName, False)
    Set rsProductList = Server.CreateObject("ADODB.Recordset")
    rsProductList.Open sqlProduct, Conn, 1, 1
    If rsProductList.BOF And rsProductList.EOF Then
        If UsePage = True And ShowType < 5 Then totalPut = 0
        If ShowType < 5 Then
            If IsHot = False And IsElite = False Then
                strProductList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & ChannelShortName & "</li>"
            ElseIf IsHot = True And IsElite = False Then
                strProductList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & strHot & ChannelShortName & "</li>"
            ElseIf IsHot = False And IsElite = True Then
                strProductList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & strElite & ChannelShortName & "</li>"
            Else
                strProductList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & strHot & strElite & ChannelShortName & "</li>"
            End If
        End If
        rsProductList.Close
        Set rsProductList = Nothing
        GetProductList = strProductList
        Exit Function
    End If
    If UsePage = True And ShowType < 5 Then
        totalPut = rsProductList.RecordCount
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
                rsProductList.Move (CurrentPage - 1) * MaxPerPage - iMod
            Else
                CurrentPage = 1
            End If
        End If
    End If

    iLines = 0

    Select Case PE_CLng(ShowType)
    Case 1, 3
        If Cols > 1 Then
            strProductList = "<table width='100%' cellpadding='0' cellspacing='0' class='" & CssNameTable & "'><tr>"
        Else
            strProductList = ""
        End If
    Case 2
        strProductList = "<table width='100%' cellpadding='2' cellspacing='1' class='" & CssNameTable & "'>"
        If ShowTableTitle = True Then
            strProductList = strProductList & "<tr align='center'><td class='" & CssNameTitle & "'></td>"
            strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(0) & "</td>"
            If ShowProductModel = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(1) & "</td>"
            End If
            If ShowProductStandard = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(2) & "</td>"
            End If
            If ShowDateType > 0 Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(3) & "</td>"
            End If
            If ShowUnit = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(4) & "</td>"
            End If
            If ShowStocksType > 0 Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(5) & "</td>"
            End If
            If ShowWeight = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(6) & "</td>"
            End If
            If ShowPrice_Market = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(7) & "</td>"
            End If
            If ShowPrice_Original = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(8) & "</td>"
            End If
            If ShowPrice = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(9) & "</td>"
            End If
            If ShowPrice_Member = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(10) & "</td>"
            End If
            If ShowDiscount = True Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(11) & "</td>"
            End If
            If ShowButtonType > 0 Then
                strProductList = strProductList & "<td class='" & CssNameTitle & "'>" & arrTitle(12) & "</td>"
            End If
            strProductList = strProductList & "</tr><tr>"
        End If
        strProductList = strProductList & "<tr>"
    End Select
    If ShowType = 5 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")

    Do While Not rsProductList.EOF
        If UsePage = True Then
            iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1
        Else
            iNumber = iCount + 1
        End If
        If TitleLen > 0 Then
            TitleStr = ReplaceText(GetSubStr(rsProductList("ProductName"), TitleLen, ShowSuspensionPoints), 2)
        Else
            TitleStr = ReplaceText(rsProductList("ProductName"), 2)
        End If
        Select Case PE_CLng(ShowPropertyType)
        Case 0
            strProperty = ""
        Case 1
            If rsProductList("OnTop") = True Then
                strProperty = "<img src=""" & ChannelUrl & "/images/Product_ontop.gif"" alt=""" & strTop & ChannelShortName & """>"
            ElseIf rsProductList("IsElite") = True Then
                strProperty = "<img src=""" & ChannelUrl & "/images/Product_elite.gif"" alt=""" & strElite & ChannelShortName & """>"
            Else
                strProperty = "<img src=""" & ChannelUrl & "/images/Product_common.gif"" alt=""" & strCommon & ChannelShortName & """>"
            End If
        Case 2
            strProperty = "��"
        Case 11
            strProperty = iNumber
        Case Else
            If rsProductList("OnTop") = True Then
                strProperty = "<img src=""" & ChannelUrl & "/images/Product_ontop" & ShowPropertyType - 1 & ".gif"" alt=""" & strTop & ChannelShortName & """>"
            ElseIf rsProductList("IsElite") = True Then
                strProperty = "<img src=""" & ChannelUrl & "/images/Product_elite" & ShowPropertyType - 1 & ".gif"" alt=""" & strElite & ChannelShortName & """>"
            Else
                strProperty = "<img src=""" & ChannelUrl & "/images/Product_common" & ShowPropertyType - 1 & ".gif"" alt=""" & strCommon & ChannelShortName & """>"
            End If
        End Select

        Dim Product_ClassID
        If UseCreateHTML > 0 Or ShowClassName = True Then
            Product_ClassID = rsProductList("ClassID")
        Else
            Product_ClassID = 0
        End If
        If ShowClassName = True And Product_ClassID <> -1 Then
            strLink = Replace(Character_Class, "{$Text}", "<a href=""" & GetClassUrl(rsProductList("ParentDir"), rsProductList("ClassDir"), Product_ClassID, 0) & """>" & rsProductList("ClassName") & "</a>")
        Else
            strLink = ""
        End If
        
        If ShowType = 5 Then
            strLink = "http://" & Trim(Request.ServerVariables("HTTP_HOST"))
            If Not (UrlType = 0 Or Left(ChannelUrl, 1) <> "/") Then
                If UseCreateHTML > 0 Then
                    strLink = strLink & GetProductUrl(rsProductList("ParentDir"), rsProductList("ClassDir"), rsProductList("UpdateTime"), rsProductList("ProductID")) & """"
                Else
                    strLink = strLink & GetProductUrl("", "", "", rsProductList("ProductID")) & """"
                End If
            End If
        Else
            If UrlType = 0 Or Left(ChannelUrl, 1) <> "/" Then
                strLink = strLink & "<a class=""productlist_A"" href="""
            Else
                strLink = strLink & "<a class=""productlist_A"" href=""http://" & Trim(Request.ServerVariables("HTTP_HOST"))
            End If
            If UseCreateHTML > 0 Then
                strLink = strLink & GetProductUrl(rsProductList("ParentDir"), rsProductList("ClassDir"), rsProductList("UpdateTime"), rsProductList("ProductID")) & """"
            Else
                strLink = strLink & GetProductUrl("", "", "", rsProductList("ProductID")) & """"
            End If
            strLink = strLink & " title=""" & rsProductList("ProductName") & """"
        
            If OpenType = 0 Then
                strLink = strLink & " target=""_self"">"
            Else
                strLink = strLink & " target=""_blank"">"
            End If
        End If
        
        Select Case PE_CLng(ShowType)
        Case 1
            If Cols > 1 Then
                strProductList = strProductList & "<td  class='" & CssName & "'>"
            End If
            strProductList = strProductList & strProperty & "&nbsp;" & strLink & TitleStr & "</a>"
            
            If ShowDateType > 0 Then
                strProductList = strProductList & "&nbsp;(" & GetUpdateTimeStr(rsProductList("UpdateTime"), ShowDateType) & ")"
            End If
            If ShowHotSign = True And rsProductList("IsHot") = True Then
                strProductList = strProductList & "<img src='" & strInstallDir & "images/hot.gif' alt='" & strHot & ChannelShortName & "'>"
            End If
            If ShowNewSign = True And DateDiff("D", rsProductList("UpdateTime"), Now()) < DaysOfNew Then
                strProductList = strProductList & "<img src='" & strInstallDir & "images/new.gif' alt='" & strNew & ChannelShortName & "'>"
            End If
            If ShowButtonType > 0 Then
                strProductList = strProductList & " " & GetButtons(ShowButtonType, ButtonStyle, rsProductList("ProductID"), strLink)
            End If
            If ContentLen > 0 Then
                strProductList = strProductList & "<div " & strList_Content_Div & ">" & Left(Replace(Replace(nohtml(rsProductList("ProductIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen) & "����</div>"
            End If
            strProductList = strProductList & "<br>"
            iCount = iCount + 1
            If Cols > 1 Then
                strProductList = strProductList & "</td>"
                If iCount Mod Cols = 0 Then
                    strProductList = strProductList & "</tr><tr>"
                    iLines = iLines + 1
                    If IntervalLines > 0 Then
                        If iLines Mod IntervalLines = 0 Then strProductList = strProductList & "<td></td></tr><tr>"
                    End If
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                End If
            Else
                If IntervalLines > 0 Then
                    If iCount Mod IntervalLines = 0 Then strProductList = strProductList & "<table class='IntervalHeight'><tr><td></td></tr></table>"
                End If
            End If
        Case 2
            If strProperty <> "" Then
                strProductList = strProductList & "<td width='10' class='" & CssName & "'>" & strProperty & "</td>"
            End If
            strProductList = strProductList & "<td class='" & CssName & "'>" & strLink & TitleStr & "</a>"
            
            If ShowHotSign = True And rsProductList("IsHot") = True Then
                strProductList = strProductList & "<img src='" & strInstallDir & "images/hot.gif' alt='" & strHot & ChannelShortName & "'>"
            End If
            If ShowNewSign = True And DateDiff("D", rsProductList("UpdateTime"), Now()) < DaysOfNew Then
                strProductList = strProductList & "<img src='" & strInstallDir & "images/new.gif' alt='" & strNew & ChannelShortName & "'>"
            End If
            strProductList = strProductList & "</td>"

            
            If ShowProductModel = True Then
                strProductList = strProductList & "<td class='" & CssName & "'>" & rsProductList("ProductModel") & "</td>"
            End If
            If ShowProductStandard = True Then
                strProductList = strProductList & "<td class='" & CssName & "'>" & rsProductList("ProductStandard") & "</td>"
            End If
            If ShowDateType > 0 Then
                strProductList = strProductList & "<td class='" & CssName & "'>" & GetUpdateTimeStr(rsProductList("UpdateTime"), ShowDateType) & "</td>"
            End If
            If ShowUnit = True Then
                strProductList = strProductList & "<td align='center' class='" & CssName & "'>" & rsProductList("Unit") & "</td>"
            End If
            Select Case PE_CLng(ShowStocksType)
            Case 1
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & rsProductList("Stocks") - rsProductList("OrderNum") & "</td>"
            Case 2
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & rsProductList("Stocks") & "</td>"
            End Select
            If ShowWeight = True Then
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & rsProductList("Weight") & "Kg</td>"
            End If
            If ShowPrice_Market = True Then
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & GetPrice_Market(rsProductList("Price_Market")) & "</td>"
            End If
            If ShowPrice_Original = True Then
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & GetPrice_FilterZero(rsProductList("Price_Original")) & "</td>"
            End If
            If ShowPrice = True Then
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & GetCurrentPrice(rsProductList("ProductType"), rsProductList("BeginDate"), rsProductList("EndDate"), rsProductList("Price_Original"), rsProductList("Price")) & "</td>"
            End If
            If ShowPrice_Member = True Then
                strProductList = strProductList & "<td align='right' class='" & CssName & "'>" & GetPrice_Member(rsProductList("Price_Member")) & "</td>"
            End If
            If ShowDiscount = True Then
                strProductList = strProductList & "<td align='center' class='" & CssName & "'>" & GetDiscount(rsProductList("ProductType"), rsProductList("Discount"), rsProductList("BeginDate"), rsProductList("EndDate")) & "</font></td>"
            End If
            If ShowButtonType > 0 Then
                strProductList = strProductList & "<td align='center' class='" & CssName & "'>" & GetButtons(ShowButtonType, ButtonStyle, rsProductList("ProductID"), strLink) & "</td>"
            End If

            iCount = iCount + 1
            If (iCount Mod Cols = 0) Or ContentLen > 0 Then
                strProductList = strProductList & "</tr>"
                If ContentLen > 0 Then
                    strProductList = strProductList & "<tr><td colspan='10'>" & Left(Replace(Replace(nohtml(rsProductList("ProductIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen) & "����</td></tr>"
                End If
                strProductList = strProductList & "<tr>"
                iLines = iLines + 1
                If IntervalLines > 0 Then
                    If iLines Mod IntervalLines = 0 Then strProductList = strProductList & "<td class='IntervalHeight'></td></tr><tr>"
                End If
                If iCount Mod (Cols * 2) = 0 Then
                    CssName = CssName1
                Else
                    CssName = CssName2
                End If
            End If
        Case 3
            If Cols > 1 Then
                strProductList = strProductList & "<td class='" & CssName & "'>"
            End If
            strProductList = strProductList & strProperty & "&nbsp;" & strLink & TitleStr & "</a>"
            If ShowDateType > 0 Then
                strProductList = strProductList & Replace(Character_Date, "{$Text}", GetUpdateTimeStr(rsProductList("UpdateTime"), ShowDateType))
            End If
            If ShowHotSign = True And rsProductList("IsHot") = True Then
                strProductList = strProductList & "<img src='" & strInstallDir & "images/hot.gif' alt='" & strHot & ChannelShortName & "'>"
            End If
            If ShowNewSign = True And DateDiff("D", rsProductList("UpdateTime"), Now()) < DaysOfNew Then
                strProductList = strProductList & "<img src='" & strInstallDir & "images/new.gif' alt='" & strNew & ChannelShortName & "'>"
            End If
            If ShowButtonType > 0 Then
                strProductList = strProductList & " " & GetButtons(ShowButtonType, ButtonStyle, rsProductList("ProductID"), strLink)
            End If
            If ContentLen > 0 Then
                strProductList = strProductList & "<div " & strList_Content_Div & ">" & Left(Replace(Replace(nohtml(rsProductList("ProductIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen) & "����</div>"
            End If
            strProductList = strProductList & "<br>"
            iCount = iCount + 1
            If Cols > 1 Then
                strProductList = strProductList & "</td>"
                If iCount Mod Cols = 0 Then
                    strProductList = strProductList & "</tr><tr>"
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                End If
            End If
        Case 4
            strProductList = strProductList & "<div class=""" & CssName & """>"
            strProductList = strProductList & strProperty & "&nbsp;" & strLink & TitleStr & "</a>"
            If ShowDateType > 0 Then
                strProductList = strProductList & Replace(Character_Date, "{$Text}", GetUpdateTimeStr(rsProductList("UpdateTime"), ShowDateType))
            End If
            If ShowHotSign = True And rsProductList("IsHot") = True Then
                strProductList = strProductList & "<img src=""" & strInstallDir & "images/hot.gif"" alt=""" & strHot & ChannelShortName & """>"
            End If
            If ShowNewSign = True And DateDiff("D", rsProductList("UpdateTime"), Now()) < DaysOfNew Then
                strProductList = strProductList & "<img src=""" & strInstallDir & "images/new.gif"" alt=""" & strNew & ChannelShortName & """>"
            End If
            If ShowButtonType > 0 Then
                strProductList = strProductList & "<div class=""product_list_button"">" & GetButtons(ShowButtonType, ButtonStyle, rsProductList("ProductID"), strLink) & "</div>"
            End If
            If ContentLen > 0 Then
                strProductList = strProductList & "<div class=""product_list_content"">" & Left(Replace(Replace(nohtml(rsProductList("ProductIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen) & "����</div>"
            End If
            strProductList = strProductList & "</div>"
            iCount = iCount + 1
            If iCount Mod 2 = 0 Then
                CssName = CssName1
            Else
                CssName = CssName2
            End If
        Case 5
            If Trim(rsProductList("ProducerName") & "") = "" Then
                iAuthor = "��"
            Else
                If Right(rsProductList("ProducerName"), 1) = "|" Then
                    iAuthor = xml_nohtml(Left(rsProductList("ProducerName"), Len(rsProductList("ProducerName")) - 1))
                Else
                    iAuthor = xml_nohtml(rsProductList("ProducerName"))
                End If
            End If
            If ShowClassName = True And Product_ClassID <> -1 Then
                strThisClass = rsProductList("ClassName")
            Else
                strThisClass = ""
            End If
            XMLDOM.appendChild (XMLDOM.createElement("item"))
            Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("title"))
            Node.Text = xml_nohtml(TitleStr)
            Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("link"))
            Node.Text = strLink
            Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("description"))
            If ContentLen > 0 Then
                Node.Text = Left(Replace(Replace(nohtml(rsProductList("ProductIntro")), ">", "&gt;"), "<", "&lt;"), ContentLen)
            End If
            Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("author"))
            Node.Text = iAuthor
            Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("category"))
            If strThisClass <> "" Then Node.Text = strThisClass
            Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("pubDate"))
            Node.Text = rsProductList("UpdateTime")
            strProductList = strProductList & XMLDOM.documentElement.xml
            iCount = iCount + 1
        End Select
        rsProductList.MoveNext
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    If ShowType = 2 Or Cols > 1 Then
        strProductList = strProductList & "</table>"
    End If
    rsProductList.Close
    Set rsProductList = Nothing
    If ShowType = 5 And RssCodeType = False Then strProductList = unicode(strProductList)
    GetProductList = strProductList
End Function

'=================================================
'��������GetPicProduct
'��  �ã���ʾ��ƷͼƬ
'��  ����
'0        arrClassID  ----��ĿID���飬0Ϊ������Ŀ
'1        IncludeChild ----�Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'2        iSpecialID ---- ר��ID��0Ϊ���в�Ʒ������ר���Ʒ�������Ϊ����0����ֻ��ʾ��Ӧר��Ĳ�Ʒ
'3        ProductNum  ---- �����ʾ���ٸ���Ʒ
'4        ProductType ---- ��Ʒ���ͣ�1Ϊ����������Ʒ��2Ϊ�Ǽ���Ʒ��3Ϊ������Ʒ��4Ϊ������Ʒ��5Ϊ�������ۺ��Ǽ���Ʒ��0Ϊ������Ʒ
'5        IsHot        ----�Ƿ������Ų�Ʒ
'6        IsElite      ----�Ƿ����Ƽ���Ʒ
'7        DateNum ----���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µĲ�Ʒ
'8        OrderType ----����ʽ��1--����ƷID����2--����ƷID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'9        ShowType   ----��ʾ��ʽ��1ΪͼƬ+����+�۸�+��ť����������
'                                  2Ϊ��ͼƬ+���ƣ��������У�+������+�۸�+��ť������������
'                                  3ΪͼƬ+������+�۸�+��ť���������У�����������
'                                  4ΪͼƬ+����+�۸���������
'                                  5Ϊ��ͼƬ+���ƣ��������У�+�۸���������
'                                  6ΪͼƬ+������+�۸��������У�����������
'                                  7ΪͼƬ+����+��ť����������
'                                  8ΪͼƬ+���ƣ���������
'                                  9ΪͼƬ+��ť����������
'                                 10Ϊֻ��ʾͼƬ
'                                 11Ϊ���DIV��ʽ
'10       ImgWidth   ----��Ʒ���
'11       ImgHeight  ----��Ʒ�߶�
'12       TitleLen   ----��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ���ƣ���Ϊ-1������ʾ��������
'13       Cols       ----ÿ�е������������������ͻ��С�

'14       UrlType ---- ���ӵ�ַ���ͣ�0Ϊ���·����1Ϊ����ַ�ľ���·����ע��˲��������⹫��

'14       ShowPriceType ---- ��ʾ�۸�ʽ��ֻ�е�ShowType������Ϊ���۸�ʽʱ����Ч��0Ϊ�Զ���ʾ��1Ϊֻ��ʾԭ�ۣ�2Ϊֻ��ʾ��ǰ�ۣ�3Ϊֻ��ʾ�г�����ԭ�ۣ�4Ϊֻ��ʾ�г����뵱ǰ�ۣ�5Ϊֻ��ʾԭ���뵱ǰ�ۣ�6Ϊֻ��ʾԭ�����Ա�ۣ�7Ϊ��ʾ�г��ۡ�ԭ�ۺ͵�ǰ�ۣ�8Ϊ��ʾ�г��ۡ�ԭ�ۺͻ�Ա�ۣ�9Ϊ��ʾ�г��ۡ���ǰ�ۺͻ�Ա�ۣ�10Ϊ��ʾ�г��ۡ�ԭ�ۡ���ǰ�ۺͻ�Ա��
'15       ShowDiscount  ---- �Ƿ���ʾ�ۿ��ʣ�ֻ�е�ShowType������Ϊ���۸�ʽʱ����Ч
'16       ShowButtonType ---- ��ť��ʾ��ʽ��ֻ�е�ShowType������Ϊ����ť��ʽʱ����Ч��1Ϊ��ʾ����ť��2Ϊ��ʾ��ϸ��ť��3Ϊ�ղذ�ť��4Ϊ��ʾ������ϸ��ť��5Ϊ��ʾ�����ղذ�ť��6Ϊ��ϸ���ղذ�ť��7Ϊ��������ʾ
'17       ButtonStyle ----  ��ť��ʽ
'18       OpenType ---- �򿪷�ʽ��0Ϊԭ���ڴ򿪣�1Ϊ�´��ڴ�
'=================================================
Public Function GetPicProduct(arrClassID, IncludeChild, iSpecialID, ProductNum, ProductType, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, Cols, UrlType, ShowPriceType, ShowDiscount, ShowButtonType, ButtonStyle, OpenType)
    Dim sqlPic, rsPic, iCount, TitleStr, strPic, strLink, trs

    iCount = 0
    ProductNum = PE_CLng(ProductNum)
    ShowType = PE_CLng(ShowType)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)
    Cols = PE_CLng(Cols)
    UrlType = PE_CLng(UrlType)

    If ProductNum < 0 Or ProductNum >= 100 Then ProductNum = 10
    If ImgWidth < 0 Or ImgWidth > 500 Then ImgWidth = 150
    If ImgHeight < 0 Or ImgHeight > 500 Then ImgHeight = 150
    If Cols <= 0 Then Cols = 5
    
    If ProductNum > 0 Then
        sqlPic = "select top " & ProductNum
    Else
        sqlPic = "select "
    End If
    sqlPic = sqlPic & " P.ProductID,P.ClassID,P.ProductName,P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,BeginDate,EndDate,P.Discount,P.UpdateTime,P.ProductThumb"
    sqlPic = sqlPic & GetSqlStr(arrClassID, IncludeChild, iSpecialID, ProductType, IsHot, IsElite, DateNum, OrderType, False, False)

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    If ShowType < 11 Then strPic = "<table cellpadding='0' cellspacing='5' border='0' width='100%'><tr valign='top'>"
    If rsPic.BOF And rsPic.EOF Then
        If ProductNum = 0 Then totalPut = 0
        If ShowType < 11 Then
            strPic = strPic & "<td align='center'><img class='pic5' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>" & R_XmlText_Class("PicProduct/NoFound", "û���κ�ͼƬ{$ChannelShortName}") & "</td></tr></table>"
        Else
            strPic = "<div class=""pic_product""><img class=""pic5"" src=""" & strInstallDir & "images/nopic.gif"" width=""" & ImgWidth & """ height=""" & ImgHeight & """ border=""0""><br>" & R_XmlText_Class("PicProduct/NoFound", "û���κ�ͼƬ{$ChannelShortName}") & "</div>"
        End If
        rsPic.Close
        Set rsPic = Nothing
        GetPicProduct = strPic
        Exit Function
    End If

    If ProductNum = 0 Then
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
    Do While Not rsPic.EOF
        If TitleLen <> 0 Then
            If TitleLen > 0 Then
                TitleStr = GetSubStr(rsPic("ProductName"), TitleLen, ShowSuspensionPoints)
            ElseIf TitleLen = -1 Then
                TitleStr = rsPic("ProductName")
            End If
        End If
        
        If UrlType = 0 Or Left(ChannelUrl, 1) <> "/" Then
            strLink = "<a href="""
        Else
            strLink = "<a href=""http://" & Trim(Request.ServerVariables("HTTP_HOST"))
        End If
        
        If UseCreateHTML > 0 Then
            strLink = strLink & GetProductUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ProductID")) & """"
        Else
            strLink = strLink & GetProductUrl("", "", "", rsPic("ProductID")) & """"
        End If
        strLink = strLink & " title=""" & rsPic("ProductName") & """"
        If OpenType = 0 Then
            strLink = strLink & " target=""_self"">"
        Else
            strLink = strLink & " target=""_blank"">"
        End If
        If ShowType < 11 Then
            strPic = strPic & "<td><table width='100%' cellspacing='2' border='0'>"
        Else
            strPic = strPic & "<div class=""pic_product"">" & vbCrLf
        End If
        
        Select Case PE_CLng(ShowType)
        Case 1  'ͼƬ+����+�۸�+��ť����������
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td></tr>"
            strPic = strPic & "<tr><td align='center'>" & strLink & TitleStr & "</a></td></tr>"
            strPic = strPic & "<tr><td align='left'>" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</td></tr>"
            strPic = strPic & "<tr><td align='center'>" & GetButtons(ShowButtonType, ButtonStyle, rsPic("ProductID"), strLink) & "</td></tr>"
        Case 2  '��ͼƬ+���ƣ��������У�+������+�۸�+��ť������������	
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td>"
            strPic = strPic & "<td align='left'>" & strLink & TitleStr & "</a><br>" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</td></tr>"
            strPic = strPic & "<tr><td align='center'>" & strLink & TitleStr & "</a></td>"
            strPic = strPic & "<td align='left'>" & GetButtons(ShowButtonType, ButtonStyle, rsPic("ProductID"), strLink) & "</td></tr>"
        Case 3  'ͼƬ+������+�۸�+��ť���������У�����������
            strPic = strPic & "<tr><td align='center' rowspan='2'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td>"
            strPic = strPic & "<td align='left'>" & strLink & TitleStr & "</a><br>" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</td></tr>"
            strPic = strPic & "<tr><td align='left' valign='bottom'>" & GetButtons(ShowButtonType, ButtonStyle, rsPic("ProductID"), strLink) & "</td></tr>"
        Case 4  'ͼƬ+����+�۸���������
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td></tr>"
            strPic = strPic & "<tr><td align='center'>" & strLink & TitleStr & "</a></td></tr>"
            strPic = strPic & "<tr><td align='left'>" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</td></tr>"
        Case 5  '��ͼƬ+���ƣ��������У�+�۸���������
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "<br>" & TitleStr & "</a></td>"
            strPic = strPic & "<td align='left'>" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</td></tr>"
        Case 6  'ͼƬ+������+�۸��������У�����������
            strPic = strPic & "<tr><td align='center'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td>"
            strPic = strPic & "<td align='left'>" & strLink & TitleStr & "</a><br>" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</td></tr>"
        Case 7  'ͼƬ+����+��ť����������
            IF ShowButtonType=0 then ShowButtonType=4		
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td></tr>"
            strPic = strPic & "<tr><td align='center'>" & strLink & TitleStr & "</a></td></tr>"
            strPic = strPic & "<tr><td align='center'>" & GetButtons(ShowButtonType, ButtonStyle, rsPic("ProductID"), strLink) & "</td></tr>"
        Case 8  'ͼƬ+���ƣ���������
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td></tr>"
            strPic = strPic & "<tr><td align='center'>" & strLink & TitleStr & "</a></td></tr>"
        Case 9  'ͼƬ+��ť����������
            IF ShowButtonType=0 then ShowButtonType=4
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td></tr>"
            strPic = strPic & "<tr><td align='center'>" & GetButtons(ShowButtonType, ButtonStyle, rsPic("ProductID"), strLink) & "</td></tr>"
        Case 10  'ֻ��ʾͼƬ
            strPic = strPic & "<tr><td align='center' class='productpic'>" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></td></tr>"
        Case 11  '���DIV��ʽ
            strPic = strPic & "<div class=""pic_product_img"">" & strLink & GetProductThumb(rsPic("ProductThumb"), ImgWidth, ImgHeight, UrlType) & "</a></div>" & vbCrLf
            strPic = strPic & "<div class=""pic_product_title"">" & strLink & TitleStr & "</a></div>" & vbCrLf
            strPic = strPic & "<div class=""pic_product_price"">" & GetProductPrice(ShowPriceType, ShowDiscount, rsPic("ProductType"), rsPic("Price_Original"), rsPic("Price"), rsPic("Price_Market"), rsPic("Price_Member"), rsPic("BeginDate"), rsPic("EndDate"), rsPic("Discount")) & "</div>" & vbCrLf
            strPic = strPic & "<div class=""pic_product_button"">" & GetButtons(ShowButtonType, ButtonStyle, rsPic("ProductID"), strLink) & "</div>" & vbCrLf
        End Select
        If ShowType < 11 Then
            strPic = strPic & "</table></td>"
        Else
            strPic = strPic & "</div>" & vbCrLf
        End If
        rsPic.MoveNext
        iCount = iCount + 1
        If ProductNum = 0 And iCount >= MaxPerPage Then Exit Do
        If ((iCount Mod Cols = 0) And (Not rsPic.EOF)) And ShowType < 11 Then strPic = strPic & "</tr><tr valign='top'>"
    Loop

    If ShowType < 11 Then strPic = strPic & "</tr></table>"
    rsPic.Close
    Set rsPic = Nothing
    GetPicProduct = strPic
End Function

'=================================================
'��������GetSlidePicProduct
'��  �ã���ʾ�õ�Ч����Ʒ
'��  ����
'0        arrClassID  ----��ĿID���飬0Ϊ������Ŀ
'1        IncludeChild ----�Ƿ��������Ŀ������arrClassIDΪ������ĿIDʱ����Ч��True----��������Ŀ��False----������
'2        iSpecialID ----ר��ID��0Ϊ������Ʒ������ר����Ʒ�������Ϊ����0����ֻ��ʾ��Ӧר�����Ʒ
'3        ProductNum ----�����ʾ���ٸ���Ʒ
'4        IsHot        ----�Ƿ���������Ʒ
'5        IsElite      ----�Ƿ����Ƽ���Ʒ
'6        DateNum ----���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ���Ʒ
'7        OrderType ----����ʽ��1--����ƷID����2--����ƷID����3--������ʱ�併��4--������ʱ������5--�����������6--�����������7--������������8--������������
'8        ImgWidth   ----��Ʒ���
'9        ImgHeight  ----��Ʒ�߶�
'10       TitleLen  ----���±����������ƣ�0Ϊ����ʾ��-1Ϊ��ʾ��������
'11       iTimeOut   ----Ч���任���ʱ�䣬�Ժ���Ϊ��λ
'12       effectID   ---- ͼƬת��Ч����0��22ָ��ĳһ����Ч��23��ʾ���Ч��
'=================================================
Public Function GetSlidePicProduct(arrClassID, IncludeChild, iSpecialID, ProductNum, IsHot, IsElite, DateNum, OrderType, ImgWidth, ImgHeight, TitleLen, iTimeOut, effectID)
    Dim sqlPic, rsPic, i, strPic, trs, tmpChannelID
    Dim ProductThumb, TitleStr
    
    ProductNum = PE_CLng(ProductNum)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)
    tmpChannelID = 0

    If ProductNum <= 0 Or ProductNum > 100 Then
        ProductNum = 10
    End If
    If ImgWidth < 0 Or ImgWidth > 1000 Then
        ImgWidth = 150
    End If
    If ImgHeight < 0 Or ImgHeight > 1000 Then
        ImgHeight = 150
    End If
    If iTimeOut < 1000 Or iTimeOut > 100000 Then
        iTimeOut = 5000
    End If
    If effectID < 0 Or effectID > 23 Then effectID = 23
    
    
    sqlPic = "select top " & ProductNum & " P.ChannelID,P.ProductID,P.ClassID,P.ProductName,P.UpdateTime,P.ProductThumb"
    sqlPic = sqlPic & GetSqlStr(arrClassID, IncludeChild, iSpecialID, 0, IsHot, IsElite, DateNum, OrderType, True, True)

    Dim ranNum
    Randomize
    ranNum = Int(900 * Rnd) + 100
    strPic = "<script language=JavaScript>" & vbCrLf
    strPic = strPic & "<!--" & vbCrLf
    strPic = strPic & "var SlidePic_" & ranNum & " = new SlidePic_Product(""SlidePic_" & ranNum & """);" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Width    = " & ImgWidth & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Height   = " & ImgHeight & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".TimeOut  = " & iTimeOut & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".Effect   = " & effectID & ";" & vbCrLf
    strPic = strPic & "SlidePic_" & ranNum & ".TitleLen = " & TitleLen & ";" & vbCrLf

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    Do While Not rsPic.EOF
        If Left(rsPic("ProductThumb"), 1) <> "/" And InStr(rsPic("ProductThumb"), "://") <= 0 Then
            ProductThumb = ChannelUrl & "/" & UploadDir & "/" & rsPic("ProductThumb")
        Else
            ProductThumb = rsPic("ProductThumb")
        End If
        If TitleLen = -1 Then
            TitleStr = rsPic("ProductName")
        Else
            TitleStr = GetSubStr(rsPic("ProductName"), TitleLen, ShowSuspensionPoints)
        End If
        
        strPic = strPic & "var oSP = new objSP_Product();" & vbCrLf
        strPic = strPic & "oSP.ImgUrl         = """ & ProductThumb & """;" & vbCrLf
        strPic = strPic & "oSP.LinkUrl        = """ & GetProductUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ProductID")) & """;" & vbCrLf
        strPic = strPic & "oSP.Title         = """ & TitleStr & """;" & vbCrLf
        strPic = strPic & "SlidePic_" & ranNum & ".Add(oSP);" & vbCrLf
        
        rsPic.MoveNext
    Loop
    strPic = strPic & "SlidePic_" & ranNum & ".Show();" & vbCrLf
    strPic = strPic & "//-->" & vbCrLf
    strPic = strPic & "</script>" & vbCrLf
    
    rsPic.Close
    Set rsPic = Nothing
    GetSlidePicProduct = strPic
End Function

Private Function JS_SlidePic()
    Dim strJS, LinkTarget
    LinkTarget = XmlText_Class("SlidePicProduct/LinkTarget", "_blank")
    strJS = strJS & "<script language=""JavaScript"">" & vbCrLf
    strJS = strJS & "<!--" & vbCrLf
    strJS = strJS & "function objSP_Product() {this.ImgUrl=""""; this.LinkUrl=""""; this.Title="""";}" & vbCrLf
    strJS = strJS & "function SlidePic_Product(_id) {this.ID=_id; this.Width=0;this.Height=0; this.TimeOut=5000; this.Effect=23; this.TitleLen=0; this.PicNum=-1; this.Img=null; this.Url=null; this.Title=null; this.AllPic=new Array(); this.Add=SlidePic_Product_Add; this.Show=SlidePic_Product_Show; this.LoopShow=SlidePic_Product_LoopShow;}" & vbCrLf
    strJS = strJS & "function SlidePic_Product_Add(_SP) {this.AllPic[this.AllPic.length] = _SP;}" & vbCrLf
    strJS = strJS & "function SlidePic_Product_Show() {" & vbCrLf
    strJS = strJS & "  if(this.AllPic[0] == null) return false;" & vbCrLf
    strJS = strJS & "  document.write(""<div align='center'><a id='Url_"" + this.ID + ""' href='' target='" & LinkTarget & "'><img id='Img_"" + this.ID + ""' style='width:"" + this.Width + ""px; height:"" + this.Height + ""px; filter: revealTrans(duration=2,transition=23);' src='javascript:null' border='0'></a>"");" & vbCrLf
    strJS = strJS & "  if(this.TitleLen != 0) {document.write(""<br><span id='Title_"" + this.ID + ""'></span></div>"");}" & vbCrLf
    strJS = strJS & "  else{document.write(""</div>"");}" & vbCrLf
    strJS = strJS & "  this.Img = document.getElementById(""Img_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Url = document.getElementById(""Url_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.Title = document.getElementById(""Title_"" + this.ID);" & vbCrLf
    strJS = strJS & "  this.LoopShow();" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function SlidePic_Product_LoopShow() {" & vbCrLf
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

Private Function GetProductPrice(ShowPriceType, ShowDiscount, ProductType, Price_Original, Price, Price_Market, Price_Member, BeginDate, EndDate, Discount)
    Dim strPrice
    Select Case ShowPriceType
    Case 0
        Select Case PE_CLng(ProductType)
        Case 0, 1, 2
            strPrice = strPrice & strPrice_Market & "<font class=""price""><STRIKE>" & GetPrice_Market(Price_Market) & "</STRIKE></font>"
            strPrice = strPrice & "<br>" & strPrice_Shop & "<font class=""price"">" & GetPrice_FilterZero(Price) & "</font>"
            strPrice = strPrice & "<br>" & strPrice_Member & "<font class=""price"">" & GetPrice_Member(Price_Member) & "</font>"
        Case 3
            If Date >= BeginDate And Date <= EndDate Then
                strPrice = strPrice & strPrice_Original & "<font class=""price""><STRIKE>" & GetPrice_FilterZero(Price_Original) & "</STRIKE></font>"
                strPrice = strPrice & "<br>" & strPrice_Te & "<font class=""price"">" & GetPrice_FilterZero(Price) & "</font>"
                strPrice = strPrice & "<br>" & strPrice_Time & "" & Month(BeginDate) & "/" & Day(BeginDate) & "-" & Month(EndDate) & "/" & Day(EndDate)
            Else
                strPrice = strPrice & strPrice_Market & "<font class=""price""><STRIKE>" & GetPrice_Market(Price_Market) & "</STRIKE></font>"
                strPrice = strPrice & "<br>" & strPrice_Shop & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font>"
            End If
        Case 4, 5
            strPrice = strPrice & strPrice_Original & "<font class=""price""><STRIKE>" & GetPrice_FilterZero(Price_Original) & "</STRIKE></font>"
            strPrice = strPrice & "<br>" & strPrice_Te & "<font class=""price"">" & GetPrice_FilterZero(Price) & "</font>"
        Case Else
            strPrice = strPrice & strPrice_Shop & "<font class=""price"">" & GetPrice_FilterZero(Price) & "</font>"
        End Select

    Case 1  'ֻ��ʾԭ��
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font>"
    Case 2  'ֻ��ʾ��ǰ��
        If ProductType = 3 And Date >= BeginDate And Date <= EndDate Then
            strPrice = strPrice & strPrice_Te & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        Else
            strPrice = strPrice & strPrice_Now & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        End If
    Case 3  'ֻ��ʾ�г�����ԭ��
        strPrice = strPrice & strPrice_Market & "<font class=""price"">" & GetPrice_Market(Price_Market) & "</font><br>"
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font>"
    Case 4  'ֻ��ʾ�г����뵱ǰ��
        strPrice = strPrice & strPrice_Market & "<font class=""price"">" & GetPrice_Market(Price_Market) & "</font><br>"
        If ProductType = 3 And Date >= BeginDate And Date <= EndDate Then
            strPrice = strPrice & strPrice_Te & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        Else
            strPrice = strPrice & strPrice_Now & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        End If
    Case 5  'ֻ��ʾԭ���뵱ǰ��
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font><br>"
        If ProductType = 3 And Date >= BeginDate And Date <= EndDate Then
            strPrice = strPrice & strPrice_Te & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        Else
            strPrice = strPrice & strPrice_Now & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        End If
    Case 6  'ֻ��ʾԭ�����Ա��
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font><br>"
        strPrice = strPrice & strPrice_Member & "<font class=""price"">" & GetPrice_Member(Price_Member) & "</font>"
    Case 7  '��ʾ�г��ۡ�ԭ�ۺ͵�ǰ��
        strPrice = strPrice & strPrice_Market & "<font class=""price"">" & GetPrice_Market(Price_Market) & "</font><br>"
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font><br>"
        If ProductType = 3 And Date >= BeginDate And Date <= EndDate Then
            strPrice = strPrice & strPrice_Te & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font>"
        Else
            strPrice = strPrice & strPrice_Now & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font>"
        End If
    Case 8  '��ʾ�г��ۡ�ԭ�ۺͻ�Ա��
        strPrice = strPrice & strPrice_Market & "<font class=""price"">" & GetPrice_Market(Price_Market) & "</font><br>"
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font><br>"
        strPrice = strPrice & strPrice_Member & "<font class=""price"">" & GetPrice_Member(Price_Member) & "</font>"
    Case 9  '��ʾ�г��ۡ���ǰ�ۺͻ�Ա��
        strPrice = strPrice & strPrice_Market & "<font class=""price"">" & GetPrice_Market(Price_Market) & "</font><br>"
        If ProductType = 3 And Date >= BeginDate And Date <= EndDate Then
            strPrice = strPrice & strPrice_Te & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        Else
            strPrice = strPrice & strPrice_Now & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        End If
        strPrice = strPrice & strPrice_Member & "<font class=""price"">" & GetPrice_Member(Price_Member) & "</font>"
    Case 10  '��ʾ�г��ۡ�ԭ�ۡ���ǰ�ۺͻ�Ա��
        strPrice = strPrice & strPrice_Market & "<font class=""price"">" & GetPrice_Market(Price_Market) & "</font><br>"
        strPrice = strPrice & strPrice_Original & "<font class=""price"">" & GetPrice_FilterZero(Price_Original) & "</font><br>"
        If ProductType = 3 And Date >= BeginDate And Date <= EndDate Then
            strPrice = strPrice & strPrice_Te & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        Else
            strPrice = strPrice & strPrice_Now & "<font class=""price"">" & GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price) & "</font><br>"
        End If
        strPrice = strPrice & strPrice_Member & "<font class=""price"">" & GetPrice_Member(Price_Member) & "</font>"
    End Select

    If ShowDiscount = True Then
        strPrice = strPrice & "<br>�ۿ��ʣ�" & GetDiscount(ProductType, Discount, BeginDate, EndDate)
    End If
    GetProductPrice = strPrice
End Function


Private Function GetCurrentPrice(ProductType, BeginDate, EndDate, Price_Original, Price)
    If ProductType = 3 Then
        If Date >= BeginDate And Date <= EndDate Then
            GetCurrentPrice = GetPrice_FilterZero(Price)
        Else
            GetCurrentPrice = GetPrice_FilterZero(Price_Original)
        End If
    Else
        GetCurrentPrice = GetPrice_FilterZero(Price)
    End If
End Function

Private Function GetPrice_Market(tPrice_Market)
    If tPrice_Market > 0 Then
        GetPrice_Market = "��" & FormatNumber(tPrice_Market, 2, vbTrue, vbFalse, vbFalse)
    Else
        GetPrice_Market = NoPrice_Market
    End If
End Function

Private Function GetPrice_Market_NoSymbol(tPrice_Market)
    If tPrice_Market > 0 Then
        GetPrice_Market_NoSymbol = FormatNumber(tPrice_Market, 2, vbTrue, vbFalse, vbFalse)
    Else
        GetPrice_Market_NoSymbol = NoPrice_Market
    End If
End Function

Private Function GetPrice_Member(tPrice_Member)
    If tPrice_Member > 0 Then
        GetPrice_Member = "��" & FormatNumber(tPrice_Member, 2, vbTrue, vbFalse, vbFalse)
    Else
        GetPrice_Member = NoPrice_Member
    End If
End Function

Private Function GetPrice_Member_NoSymbol(tPrice_Member)
    If tPrice_Member > 0 Then
        GetPrice_Member_NoSymbol = FormatNumber(tPrice_Member, 2, vbTrue, vbFalse, vbFalse)
    Else
        GetPrice_Member_NoSymbol = NoPrice_Member
    End If
End Function

Private Function GetPrice_FilterZero(tPrice)
    If tPrice > 0 Then
        GetPrice_FilterZero = "��" & FormatNumber(tPrice, 2, vbTrue, vbFalse, vbFalse)
    Else
        GetPrice_FilterZero = NoPrice
    End If
End Function

Private Function GetPrice_FilterZero_NoSymbol(tPrice)
    If tPrice > 0 Then
        GetPrice_FilterZero_NoSymbol = FormatNumber(tPrice, 2, vbTrue, vbFalse, vbFalse)
    Else
        GetPrice_FilterZero_NoSymbol = NoPrice
    End If
End Function

Private Function GetDiscount(ProductType, Discount, BeginDate, EndDate)
    Select Case PE_CLng(ProductType)
    Case 1, 2
        GetDiscount = "100%"
    Case 3
        If Date >= BeginDate And Date <= EndDate Then
            GetDiscount = CLng(Discount * 10) & "%"
        Else
            GetDiscount = "100%"
        End If
    Case 5
        GetDiscount = CLng(Discount * 10) & "%"
    Case 4
        GetDiscount = "��"
    End Select
End Function

Private Function GetUpdateTimeStr(UpdateTime, ShowDateType)
    Dim strUpdateTime
    If Not IsDate(UpdateTime) Then
        GetUpdateTimeStr = ""
        Exit Function
    End If
    Select Case PE_CLng(ShowDateType)
    Case 1
        strUpdateTime = Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2)
    Case 2
        strUpdateTime = Month(UpdateTime) & strMonth & Day(UpdateTime) & strDay
    Case 3
        strUpdateTime = Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2)
    Case 4
        strUpdateTime = Year(UpdateTime) & strYear & Month(UpdateTime) & strMonth & Day(UpdateTime) & strDay
    Case 5
        strUpdateTime = UpdateTime
    Case 6
        strUpdateTime = UpdateTime
    End Select
    If DateDiff("D", UpdateTime, Now()) < DaysOfNew Then
        strUpdateTime = "<font " & strListStr_Font & ">" & strUpdateTime & "</font>"
    End If
    GetUpdateTimeStr = strUpdateTime
End Function

Function GetButtons(iShowButtonType, iButtonStyle, iProductID, strLink)
    Dim strButtons, ImgUrl_Buy, ImgUrl_Content, ImgUrl_Fav
    If iButtonStyle > 0 Then
        ImgUrl_Buy = ChannelUrl & "/images/ProductBuy" & iButtonStyle & ".gif"
        ImgUrl_Content = ChannelUrl & "/images/ProductContent" & iButtonStyle & ".gif"
        ImgUrl_Fav = ChannelUrl & "/images/ProductFav" & iButtonStyle & ".gif"
    Else
        ImgUrl_Buy = ChannelUrl & "/images/ProductBuy.gif"
        ImgUrl_Content = ChannelUrl & "/images/ProductContent.gif"
        ImgUrl_Fav = ChannelUrl & "/images/ProductFav.gif"
    End If
    Select Case PE_CLng(iShowButtonType)
    Case 1
        strButtons = "<a href=""" & ChannelUrl_ASPFile & "/ShoppingCart.asp?Action=Add&ProductID=" & iProductID & """ target=""ShoppingCart""><img src=""" & ImgUrl_Buy & """ border=""0""></a>"
    Case 2
        strButtons = "" & strLink & "<img src=""" & ImgUrl_Content & """ border=""0""></a>"
    Case 3
        strButtons = "<a href=""" & strInstallDir & "User/User_Favorite.asp?Action=Add&ChannelID=" & ChannelID & "&InfoID=" & iProductID & """><img src=""" & ImgUrl_Fav & """ border=0></a>"
    Case 4
        strButtons = "<a href=""" & ChannelUrl_ASPFile & "/ShoppingCart.asp?Action=Add&ProductID=" & iProductID & """ target=""ShoppingCart""><img src=""" & ImgUrl_Buy & """ border=""0""></a>"
        strButtons = strButtons & " " & strLink & "<img src=""" & ImgUrl_Content & """ border=""0""></a>"
    Case 5
        strButtons = "<a href=""" & ChannelUrl_ASPFile & "/ShoppingCart.asp?Action=Add&ProductID=" & iProductID & """ target=""ShoppingCart""><img src=""" & ImgUrl_Buy & """ border=""0""></a>"
        strButtons = strButtons & " <a href=""" & strInstallDir & "User/User_Favorite.asp?Action=Add&ChannelID=" & ChannelID & "&InfoID=" & iProductID & """><img src=""" & ImgUrl_Fav & """ border=0></a>"
    Case 6
        strButtons = "" & strLink & "<img src=""" & ImgUrl_Content & """ border=""0""></a>"
        strButtons = strButtons & " <a href=""" & strInstallDir & "User/User_Favorite.asp?Action=Add&ChannelID=" & ChannelID & "&InfoID=" & iProductID & """><img src=""" & ImgUrl_Fav & """ border=0></a>"
    Case 7
        strButtons = "<a href=""" & ChannelUrl_ASPFile & "/ShoppingCart.asp?Action=Add&ProductID=" & iProductID & """ target=""ShoppingCart""><img src=""" & ImgUrl_Buy & """ border=""0""></a>"
        strButtons = strButtons & " " & strLink & "<img src=""" & ImgUrl_Content & """ border=""0""></a>"
        strButtons = strButtons & " <a href=""" & strInstallDir & "User/User_Favorite.asp?Action=Add&ChannelID=" & ChannelID & "&InfoID=" & iProductID & """><img src=""" & ImgUrl_Fav & """ border=0></a>"
    End Select
    GetButtons = strButtons
End Function

Private Function GetProductThumb(ProductThumb, iWidth, iHeight, UrlType)
    Dim strProductThumb, FileType, strPicUrl
    If UrlType = 0 Then
        strPicUrl = ""
    Else
        strPicUrl = "http://" & Trim(Request.ServerVariables("HTTP_HOST"))
    End If
    
    If ProductThumb = "" Then
        strProductThumb = strProductThumb & "<img src=""" & strPicUrl & strInstallDir & "images/nopic.gif"" "
        If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
        If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
        strProductThumb = strProductThumb & " border=""0"">"
    Else
        FileType = LCase(Mid(ProductThumb, InStrRev(ProductThumb, ".") + 1))
        If Left(ProductThumb, 1) <> "/" And InStr(ProductThumb, "://") <= 0 Then
            If Left(ChannelUrl, 1) = "/" Then
                strPicUrl = strPicUrl & ChannelUrl & "/" & UploadDir & "/" & ProductThumb
            Else
                strPicUrl = ChannelUrl & "/" & UploadDir & "/" & ProductThumb
            End If
        Else
            strPicUrl = ProductThumb
        End If
        If FileType = "swf" Then
            strProductThumb = strProductThumb & "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"" "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"""
            strProductThumb = strProductThumb & "><param name=""movie"" value=""" & strPicUrl & """><param name=""quality"" value=""high""><embed src=""" & strPicUrl & """ pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
            strProductThumb = strProductThumb & "></embed></object>"
        ElseIf FileType = "gif" Or FileType = "jpg" Or FileType = "jpeg" Or FileType = "jpe" Or FileType = "bmp" Or FileType = "png" Then
            strProductThumb = strProductThumb & "<img class=""pic5"" src=""" & strPicUrl & """ "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
            strProductThumb = strProductThumb & " border=""0"">"
        Else
            strProductThumb = strProductThumb & "<img src=""" & strInstallDir & "images/nopic.gif"" "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
            strProductThumb = strProductThumb & " border=""0"">"
        End If
    End If
    GetProductThumb = strProductThumb
End Function

Private Function GetJsProductThumb(ProductThumb, iWidth, iHeight, UrlType)
    Dim strProductThumb, FileType, strPicUrl
    If UrlType = 0 Then
        strPicUrl = ""
    Else
        strPicUrl = "http://" & Trim(Request.ServerVariables("HTTP_HOST"))
    End If
    
    If ProductThumb = "" Then
        strProductThumb = strProductThumb & "<img src=""" & strPicUrl & strInstallDir & "images/nopic.gif"" "
        If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
        If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
        strProductThumb = strProductThumb & " border=""0"">"
    Else
        FileType = LCase(Mid(ProductThumb, InStrRev(ProductThumb, ".") + 1))
        If Left(ProductThumb, 1) <> "/" And InStr(ProductThumb, "://") <= 0 Then
            If Left(ChannelUrl, 1) = "/" Then
                strPicUrl = strPicUrl & ChannelUrl & "/" & UploadDir & "/" & ProductThumb
            Else
                strPicUrl = ChannelUrl & "/" & UploadDir & "/" & ProductThumb
            End If
        Else
            strPicUrl = ProductThumb
        End If
        If FileType = "swf" Then
            strProductThumb = strProductThumb & "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"" "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"""
            strProductThumb = strProductThumb & "><param name=""movie"" value=""" & strPicUrl & """><param name=""quality"" value=""high""><embed src=""" & strPicUrl & """ pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
            strProductThumb = strProductThumb & "></embed></object>"
        ElseIf FileType = "gif" Or FileType = "jpg" Or FileType = "jpeg" Or FileType = "jpe" Or FileType = "bmp" Or FileType = "png" Then
            strProductThumb = strProductThumb & "<a id=""pid"&Productid&"""  title=""����Ŵ�""  href=""" & strPicUrl & """ class=""highslide"" onclick=""return hs.expand(this, {captionId: 'pro"&Productid&"'})""><img class=""pic5"" src=""" & strPicUrl & """ "
            'strProductThumb = strProductThumb & "<img class=""pic5"" src=""" & strPicUrl & """ "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
            strProductThumb = strProductThumb & " border=""0"">"
        Else
            strProductThumb = strProductThumb & "<img src=""" & strInstallDir & "images/nopic.gif"" "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width=""" & iWidth & """"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height=""" & iHeight & """"
            strProductThumb = strProductThumb & " border=""0"">"
        End If
    End If
    GetJsProductThumb = strProductThumb
End Function

Private Function GetSearchResultIDArr()

    Dim sqlSearch, rsSearch
    Dim rsField
    Dim iNum, arrProductID

    If PE_CLng(SearchResultNum) > 0 Then
        sqlSearch = "select top " & PE_CLng(SearchResultNum) & " ProductID "
    Else
        sqlSearch = "select ProductID "
    End If
    sqlSearch = sqlSearch & " from PE_Product where Deleted=" & PE_False & " and EnableSale=" & PE_True & ""
    'If ChannelID > 0 Then
    '    sqlSearch = sqlSearch & " and ChannelID=" & ChannelID & " "
    'End If
    If ClassID > 0 Then
        If Child > 0 Then
            sqlSearch = sqlSearch & " and ClassID in (" & arrChildID & ")"
        Else
            sqlSearch = sqlSearch & " and ClassID=" & ClassID
        End If
    End If
    If SpecialID > 0 Then
        sqlSearch = sqlSearch & " and ProductID in (select ItemID from PE_InfoS where SpecialID=" & SpecialID & ")"
    End If
    If strField <> "" Then  '��ͨ����
        Select Case strField
            Case "Title", "ProductName"
                sqlSearch = sqlSearch & SetSearchString("ProductName")
            Case "Content", "ProductIntro"
                sqlSearch = sqlSearch & SetSearchString("ProductIntro")
            Case "ProducerName"
                sqlSearch = sqlSearch & SetSearchString("ProducerName")
            Case "TrademarkName"
                sqlSearch = sqlSearch & SetSearchString("TrademarkName")
            Case "ProductModel"
                sqlSearch = sqlSearch & SetSearchString("ProductModel")
            Case "ProductStandard"
                sqlSearch = sqlSearch & SetSearchString("ProductStandard")
            Case "ProductNum"
                sqlSearch = sqlSearch & SetSearchString("ProductNum")
            Case "Keywords"
                sqlSearch = sqlSearch & SetSearchString("Keyword")
            Case Else  '�Զ����ֶ�
                Set rsField = Conn.Execute("select Title from PE_Field where (ChannelID=-5 or ChannelID=" & ChannelID & ") and FieldName='" & ReplaceBadChar(strField) & "'")
                If rsField.BOF And rsField.EOF Then
                    sqlSearch = sqlSearch & SetSearchString("ProductName")
                Else
                    sqlSearch = sqlSearch & SetSearchString(ReplaceBadChar(strField))
                End If
                rsField.Close
                Set rsField = Nothing
        End Select
    Else   '�߼�����
        '����߼���������
        Dim ProductNum, ProductName, ProductIntro, ProductExplain, ProducerName, TrademarkName, ProductModel, ProductStandard, LowPrice, HighPrice, BeginDate, EndDate
        ProductName = Trim(Request("ProductName"))
        ProductIntro = Trim(Request("ProductIntro"))
        ProductExplain = Trim(Request("ProductExplain"))
        ProducerName = Trim(Request("ProducerName"))
        TrademarkName = Trim(Request("TrademarkName"))
        ProductModel = Trim(Request("ProductModel"))
        ProductStandard = Trim(Request("ProductStandard"))
        LowPrice = PE_CLng(Request("LowPrice"))
        HighPrice = PE_CLng(Request("HighPrice"))
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        ProductNum = Trim(Request("ProductNum"))
        strFileName = "Search.asp?ModuleName=Shop&ClassID=" & ClassID & "&SpecialID=" & SpecialID
        If ProductNum <> "" Then
            ProductNum = ReplaceBadChar(ProductNum)
            strFileName = strFileName & "&ProductNum=" & ProductNum
            sqlSearch = sqlSearch & " and ProductNum like '%" & ProductNum & "%' "
            
        End If
        If ProductName <> "" Then
            ProductName = ReplaceBadChar(ProductName)
            strFileName = strFileName & "&ProductName=" & ProductName
            sqlSearch = sqlSearch & " and ProductName like '%" & ProductName & "%' "
        End If
        If ProductIntro <> "" Then
            ProductIntro = ReplaceBadChar(ProductIntro)
            strFileName = strFileName & "&ProductIntro=" & ProductIntro
            sqlSearch = sqlSearch & " and ProductIntro like '%" & ProductIntro & "%'"
        End If
        If ProductExplain <> "" Then
            ProductExplain = ReplaceBadChar(ProductExplain)
            strFileName = strFileName & "&ProductExplain=" & ProductExplain
            sqlSearch = sqlSearch & " and ProductExplain like '%" & ProductExplain & "%'"
        End If
        If ProducerName <> "" Then
            ProducerName = ReplaceBadChar(ProducerName)
            strFileName = strFileName & "&ProducerName=" & ProducerName
            sqlSearch = sqlSearch & " and ProducerName like '%" & ProducerName & "%' "
        End If
        If TrademarkName <> "" Then
            TrademarkName = ReplaceBadChar(TrademarkName)
            strFileName = strFileName & "&TrademarkName=" & TrademarkName
            sqlSearch = sqlSearch & " and TrademarkName='" & TrademarkName & "' "
        End If
        If ProductModel <> "" Then
            ProductModel = ReplaceBadChar(ProductModel)
            strFileName = strFileName & "&ProductModel=" & ProductModel
            sqlSearch = sqlSearch & " and ProductModel like '%" & ProductModel & "%' "
        End If
        If ProductStandard <> "" Then
            ProductStandard = ReplaceBadChar(ProductStandard)
            strFileName = strFileName & "&ProductStandard=" & ProductStandard
            sqlSearch = sqlSearch & " and ProductStandard='" & ProductStandard & "' "
        End If
    
        If LowPrice > 0 Then
            strFileName = strFileName & "&LowPrice=" & LowPrice
            sqlSearch = sqlSearch & " and Price >=" & LowPrice
        End If
        If HighPrice > 0 Then
            strFileName = strFileName & "&HighPrice=" & HighPrice
            sqlSearch = sqlSearch & " and Price <=" & HighPrice
        End If

        If IsDate(BeginDate) Then
            strFileName = strFileName & "&BeginDate=" & BeginDate
            If SystemDatabaseType = "SQL" Then
                sqlSearch = sqlSearch & " and BeginDate >= '" & BeginDate & "'"
            Else
                sqlSearch = sqlSearch & " and BeginDate >= #" & BeginDate & "#"
            End If
        End If
        If IsDate(EndDate) Then
            strFileName = strFileName & "&EndDate=" & EndDate
            If SystemDatabaseType = "SQL" Then
                sqlSearch = sqlSearch & " and EndDate <= '" & EndDate & "'"
            Else
                sqlSearch = sqlSearch & " and EndDate <= #" & EndDate & "#"
            End If
        End If

        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-5 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            If Trim(Request(rsField("FieldName"))) <> "" Then
                strFileName = strFileName & "&" & Trim(rsField("FieldName")) & "=" & ReplaceBadChar(Trim(Request(rsField("FieldName"))))
                sqlSearch = sqlSearch & " and " & Trim(rsField("FieldName")) & " like '%" & ReplaceBadChar(Trim(Request(rsField("FieldName")))) & "%' "
            End If
            rsField.MoveNext
        Loop
        Set rsField = Nothing
        
    End If
    
    sqlSearch = sqlSearch & " order by ProductID desc"
    arrProductID = ""
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
        iNum = 0
        Do While Not rsSearch.EOF
            If arrProductID = "" Then
                arrProductID = rsSearch(0)
            Else
                arrProductID = arrProductID & "," & rsSearch(0)
            End If
            iNum = iNum + 1
            If iNum >= MaxPerPage Then Exit Do
            rsSearch.MoveNext
        Loop
    End If
    rsSearch.Close
    Set rsSearch = Nothing

    GetSearchResultIDArr = arrProductID
End Function

'=================================================
'��������GetSearchResult
'��  �ã���ҳ��ʾ�������
'��  ����ImgWidth----ͼƬ���
'        ImgHeight----ͼƬ�߶�
'        Cols ---- ÿ����ʾ����
'=================================================
Private Function GetSearchResult(ImgWidth, ImgHeight, Cols)
    
    If Cols <= 0 Then Cols = 1
    
    Dim sqlSearch, rsSearch, iCount, iNum, arrProductID, strSearchResult, Content, TitleStr, strLink
    strSearchResult = ""
    arrProductID = GetSearchResultIDArr()
    If arrProductID = "" Then
        GetSearchResult = "<p align='center'><br><br>û�л�û���ҵ��κ�" & ChannelShortName & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If

    strSearchResult = "<table width='100%' cellpadding='0' cellspacing='5' border='0' align='center'><tr valign='top'>"
    iNum = 0
    
    sqlSearch = "select P.ProductID,P.ProductName,P.Discount,P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,BeginDate,EndDate,P.UpdateTime,P.ProductThumb,C.ClassID,C.ClassName,C.ParentDir,C.ClassDir from PE_Product P left join PE_Class C on P.ClassID=C.ClassID where ProductID in (" & arrProductID & ") order by ProductID desc"
    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1
    Do While Not rsSearch.EOF
        TitleStr = ReplaceText(rsSearch("ProductName"), 2)
        If UseCreateHTML > 0 Then
            strLink = "<a class='LinkSearchResult' href='" & GetProductUrl(rsSearch("ParentDir"), rsSearch("ClassDir"), rsSearch("UpdateTime"), rsSearch("ProductID")) & "'"
        Else
            strLink = "<a class='LinkSearchResult' href='" & GetProductUrl("", "", "", rsSearch("ProductID")) & "'"
        End If
        strLink = strLink & " target='_blank'>"

        strSearchResult = strSearchResult & "<td><table width='100%' cellspacing='2' border='0'>"
            
        strSearchResult = strSearchResult & "<tr><td align='center' rowspan='2'>" & strLink & GetProductThumb(rsSearch("ProductThumb"), ImgWidth, ImgHeight, 1) & "</a></td>"
        strSearchResult = strSearchResult & "<td align='left'>" & strLink & TitleStr & "</a><br>" & GetProductPrice(0, False, rsSearch("ProductType"), rsSearch("Price_Original"), rsSearch("Price"), rsSearch("Price_Market"), rsSearch("Price_Member"), rsSearch("BeginDate"), rsSearch("EndDate"), rsSearch("Discount")) & "</td></tr>"
        strSearchResult = strSearchResult & "<tr><td align='left' valign='bottom'><a href='" & ChannelUrl_ASPFile & "/ShoppingCart.asp?Action=Add&ProductID=" & rsSearch("ProductID") & "' target='ShoppingCart'><img src='" & ChannelUrl & "/images/ProductBuy.gif' border='0'></a>&nbsp;&nbsp;" & strLink & "<img src='" & ChannelUrl & "/images/ProductContent.gif' border='0'></a></td></tr>"

        strSearchResult = strSearchResult & "</table></td>"
        rsSearch.MoveNext
        iNum = iNum + 1
        If (iNum Mod Cols = 0) And Not rsSearch.EOF Then strSearchResult = strSearchResult & "</tr><tr valign='top'>"
    Loop
    rsSearch.Close
    strSearchResult = strSearchResult & "</tr></table>"

    Set rsSearch = Nothing
    GetSearchResult = strSearchResult
End Function


Private Function GetSearchResult2(strValue)    '�õ��Զ����б�İ�����Ƶ�HTML����
    Dim strCustom, strParameter
    strCustom = strValue
    regEx.Pattern = "��SearchResultList\((.*?)\)��([\s\S]*?)��\/SearchResultList��"
    Set Matches = regEx.Execute(strCustom)
    For Each Match In Matches
        strParameter = Replace(Match.SubMatches(0), Chr(34), " ")
        strCustom = PE_Replace(strCustom, Match.Value, GetSearchResultLabel(strParameter, Match.SubMatches(1)))
    Next
    GetSearchResult2 = strCustom
End Function

Private Function GetSearchResultLabel(strTemp, strList)
    Dim sqlSearch, rsSearch, iCount, arrProductID, Content, TitleStr, strLink
    Dim arrTemp
    Dim strProductThumb, arrPicTemp
    Dim arrClassID, IncludeChild, iSpecialID, ProductType, iNum, IsHot, IsElite, Author, DateNum, OrderType, UsePage, TitleLen, ContentLen
    Dim iCols, iColsHtml, iRows, iRowsHtml, iNumber
    Dim rsField, ArrField, iField
    Dim rsCustom, strCustomList
                
    If strTemp = "" Or strList = "" Then GetSearchResultLabel = "": Exit Function
    iCount = 0
    strCustomList = ""
    
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


    arrProductID = GetSearchResultIDArr()
    If arrProductID = "" Then
        GetSearchResultLabel = "<p align='center'><br><br>û�л�û���ҵ��κ�" & ChannelShortName & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If

    Set rsField = Conn.Execute("select FieldName,LabelName from PE_Field where ChannelID=-5 or ChannelID=" & ChannelID & "")
    If Not (rsField.BOF And rsField.EOF) Then
        ArrField = rsField.getrows(-1)
    End If
    Set rsField = Nothing
    
    sqlSearch = "select P.ProductID,P.ProductName,P.ProductExplain,P.ProductIntro,P.LimitNum,P.Stars,P.Discount,P.ProductType,P.Price,"
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlSearch = sqlSearch & "P." & ArrField(0, iField) & ","
        Next
    End If
    sqlSearch = sqlSearch & "Price_Original,P.Keyword,P.PresentPoint,P.IsHot,P.PresentExp,P.PresentMoney,P.IsElite,P.OnTop,P.ProductNum,P.BarCode,P.Stocks,P.Unit,P.OrderNum,P.Hits,P.Inputer,P.ProductStandard,P.ProductModel,P.TrademarkName,P.ProducerName,P.Price_Market,P.Price_Member,BeginDate,EndDate,P.UpdateTime,P.ProductThumb,C.ClassID,C.ClassName,C.ParentDir,"
    sqlSearch = sqlSearch & "C.ClassDir from PE_Product P left join PE_Class C on P.ClassID=C.ClassID where ProductID in (" & arrProductID & ") order by ProductID desc"
    Set rsCustom = Server.CreateObject("ADODB.Recordset")
    rsCustom.Open sqlSearch, Conn, 1, 1
    Do While Not rsCustom.EOF
        strTemp = strList
        iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1

        strTemp = PE_Replace(strTemp, "{$Number}", iNumber)
        strTemp = PE_Replace(strTemp, "{$ClassID}", rsCustom("ClassID"))
        strTemp = PE_Replace(strTemp, "{$ClassName}", rsCustom("ClassName"))
        strTemp = PE_Replace(strTemp, "{$ParentDir}", rsCustom("ParentDir"))
        strTemp = PE_Replace(strTemp, "{$ClassDir}", rsCustom("ClassDir"))
        If InStr(strTemp, "{$ClassUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ClassUrl}", GetClassUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("ClassID"), 0))

        If InStr(strTemp, "{$ProductUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ProductUrl}", GetProductUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("ProductID")))
        strTemp = PE_Replace(strTemp, "{$ProductID}", rsCustom("ProductID"))
        strTemp = PE_Replace(strTemp, "{$ProductNum}", rsCustom("ProductNum"))
        strTemp = PE_Replace(strTemp, "{$BarCode}", rsCustom("BarCode"))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        strTemp = Replace(strTemp, "{$ProducerName}", rsCustom("ProducerName"))
        strTemp = Replace(strTemp, "{$TrademarkName}", rsCustom("TrademarkName"))
        strTemp = PE_Replace(strTemp, "{$ProductModel}", rsCustom("ProductModel"))
        strTemp = PE_Replace(strTemp, "{$ProductStandard}", rsCustom("ProductStandard"))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$Inputer}", rsCustom("Inputer"))
        strTemp = PE_Replace(strTemp, "{$Unit}", rsCustom("Unit"))
        strTemp = PE_Replace(strTemp, "{$Stocks}", rsCustom("Stocks") - rsCustom("OrderNum"))
        If InStr(strTemp, "{$Keyword}") > 0 Then strTemp = PE_Replace(strTemp, "{$Keyword}", GetKeywords(",", rsCustom("Keyword")))

        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "OnTop")
        ElseIf rsCustom("IsElite") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Elite")
        ElseIf rsCustom("IsHot") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Hot")
        Else
            strTemp = PE_Replace(strTemp, "{$Property}", "Common")
        End If
        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Top}", strTop2)
        Else
            strTemp = PE_Replace(strTemp, "{$Top}", "")
        End If
        If rsCustom("IsElite") = True Then
            strTemp = PE_Replace(strTemp, "{$Elite}", strElite2)
        Else
            strTemp = PE_Replace(strTemp, "{$Elite}", "")
        End If
        If rsCustom("IsHot") = True Then
            strTemp = PE_Replace(strTemp, "{$Hot}", strHot2)
        Else
            strTemp = PE_Replace(strTemp, "{$Hot}", "")
        End If
        
        If TitleLen > 0 Then
            strTemp = PE_Replace(strTemp, "{$ProductName}", GetSubStr(rsCustom("ProductName"), TitleLen, ShowSuspensionPoints))
        Else
            strTemp = PE_Replace(strTemp, "{$ProductName}", rsCustom("ProductName"))
        End If
        strTemp = PE_Replace(strTemp, "{$ProductNameOriginal}", rsCustom("ProductName"))
        strTemp = PE_Replace(strTemp, "{$ProductIntro}", rsCustom("ProductIntro"))
        If ContentLen > 0 Then
            If InStr(strTemp, "{$ProductExplain}") > 0 Then strTemp = PE_Replace(strTemp, "{$ProductExplain}", Left(nohtml(rsCustom("ProductExplain")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$ProductExplain}", "")
        End If

        If InStr(strTemp, "{$ProductThumb}") > 0 Then strTemp = Replace(strTemp, "{$ProductThumb}", GetProductThumb(rsCustom("ProductThumb"), 130, 0, 0))
        If InStr(strTemp, "{$JsProductThumb}") > 0 Then strTemp = Replace(strTemp, "{$JsProductThumb}", GetJsProductThumb(rsCustom("ProductThumb"), 130, 0, 0))
        '�滻��ҳͼƬ
        regEx.Pattern = "\{\$ProductThumb\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strProductThumb = GetProductThumb(Trim(rsCustom("ProductThumb")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)), 0)
            strTemp = Replace(strTemp, Match.Value, strProductThumb)
        Next
        Dim arrJsPicTemp , strJsProductThumb
        regEx.Pattern = "\{\$JsProductThumb\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrJsPicTemp = Split(Match.SubMatches(0), ",")
            strJsProductThumb = GetJsProductThumb(Trim(rsCustom("ProductThumb")), PE_CLng(arrJsPicTemp(0)), PE_CLng(arrJsPicTemp(1)), 0)
            strTemp = Replace(strTemp, Match.Value, strJsProductThumb)
        Next
        
        If InStr(strTemp, "{$Price_Original}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price_Original}", GetPrice_FilterZero_NoSymbol(rsCustom("Price_Original")))
        If rsCustom("Price_Market") > 0 Then
            strTemp = PE_Replace(strTemp, "{$Price_Market}", rsCustom("Price_Market"))
        Else
            If InStr(strTemp, "{$Price_Market}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price_Market}", GetPrice_Market_NoSymbol(rsCustom("Price_Original")))
        End If
        
        Select Case rsCustom("ProductType")
        Case 1
            strTemp = Replace(strTemp, "{$ProductTypeName}", "����������Ʒ")
            If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price_Original")))
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", "")
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        Case 2
            If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price")))
            strTemp = Replace(strTemp, "{$ProductTypeName}", "�Ǽ���Ʒ")
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", "")
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        Case 3
            If rsCustom("BeginDate") <= Date And rsCustom("EndDate") >= Date Then
                strTemp = Replace(strTemp, "{$ProductTypeName}", "�ؼ۴�����Ʒ")
                If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price")))
                strTemp = PE_Replace(strTemp, "{$BeginDate}", rsCustom("BeginDate"))
                strTemp = PE_Replace(strTemp, "{$EndDate}", rsCustom("EndDate"))
                strTemp = PE_Replace(strTemp, "{$Discount}", rsCustom("Discount"))
                strTemp = PE_Replace(strTemp, "{$LimitNum}", rsCustom("LimitNum"))
            Else
                strTemp = Replace(strTemp, "{$ProductTypeName}", "����������Ʒ")
                If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price_Original")))
                strTemp = Replace(strTemp, "{$BeginDate}", "")
                strTemp = Replace(strTemp, "{$EndDate}", "")
                strTemp = Replace(strTemp, "{$Discount}", "")
                strTemp = Replace(strTemp, "{$LimitNum}", "")
            End If
        Case 4
            strTemp = Replace(strTemp, "{$ProductTypeName}", "������Ʒ�����������ۣ�")
            strTemp = PE_Replace(strTemp, "{$Price}", rsCustom("Price"))
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", "")
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        Case 5
            strTemp = Replace(strTemp, "{$ProductTypeName}", "���۴���")
            If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price")))
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", rsCustom("Discount"))
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        End Select
        If InStr(strTemp, "{$Price_Member}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price_Member}", GetPrice_Member_NoSymbol(rsCustom("Price_Member")))
        strTemp = PE_Replace(strTemp, "{$PresentExp}", rsCustom("PresentExp"))
        strTemp = PE_Replace(strTemp, "{$PresentPoint}", rsCustom("PresentPoint"))
        strTemp = PE_Replace(strTemp, "{$PresentMoney}", rsCustom("PresentMoney"))
        strTemp = PE_Replace(strTemp, "{$PointName}", PointName)
        strTemp = PE_Replace(strTemp, "{$PointUnit}", PointUnit)
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
'��  �ã���ʾ��ز�Ʒ
'��  ����ProductNum  ----�����ʾ���ٸ���Ʒ
'        TitleLen   ----��������ַ�����һ������=����Ӣ���ַ�
'        ShowType  ----��ʾ��ʽ
'        ImgWidth ---- ͼƬ���
'        ImgHeight ---- ͼƬ�߶�
'        Cols  -------  ÿ����ʾ������
'=================================================
Private Function GetCorrelative(ProductNum, TitleLen, ShowType, ImgWidth, ImgHeight, Cols)
    Dim rsCorrelative, sqlCorrelative, strCorrelative, strLink
    Dim strKey, arrKey, i, iNum, MaxNum
    If ImgWidth < 0 Then ImgWidth = 100
    If ImgHeight < 0 Then ImgHeight = 100
    If Cols <= 0 Then Cols = 1
    
    If ProductNum > 0 And ProductNum <= 100 Then
        sqlCorrelative = "select top " & ProductNum
    Else
        sqlCorrelative = "Select Top 5 "
    End If
    strKey = Mid(rsProduct("Keyword"), 2, Len(rsProduct("Keyword")) - 2)
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
    sqlCorrelative = sqlCorrelative & " P.ProductID,P.ProductName,P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,Discount,BeginDate,EndDate,P.UpdateTime,P.ProductThumb"
    If UseCreateHTML > 0 Then
        sqlCorrelative = sqlCorrelative & ",C.ParentDir,C.ClassDir from PE_Product P left join PE_Class C on P.ClassID=C.ClassID"
    Else
        sqlCorrelative = sqlCorrelative & " from PE_Product P"
    End If
    sqlCorrelative = sqlCorrelative & " where P.Deleted=" & PE_False & " and P.EnableSale=" & PE_True & ""

    sqlCorrelative = sqlCorrelative & " and " & strKey & " and P.ProductID<>" & ProductID & " Order by P.ProductID desc"
    Set rsCorrelative = Conn.Execute(sqlCorrelative)
    If TitleLen < 0 Or TitleLen > 255 Then TitleLen = 50
    If rsCorrelative.BOF And rsCorrelative.EOF Then
        If ShowType = 1 Then
            strCorrelative = R_XmlText_Class("ShowProduct/NoCorrelative", "û�����{$ChannelShortName}")
        Else
            strCorrelative = "<table align='center'><tr><td align='center' class='tdbg3'><img class='pic5' src='" & strInstallDir & "images/nopic.gif' width='130' height='90' border='0'><br>û���ղ��κ�" & ChannelShortName & "</td></tr></table>"
        End If
    Else
        If ShowType = 1 Then
            Do While Not rsCorrelative.EOF
                If UseCreateHTML > 0 Then
                    strCorrelative = strCorrelative & "<li><a class='LinkProductCorrelative' href='" & GetProductUrl(rsCorrelative("ParentDir"), rsCorrelative("ClassDir"), rsCorrelative("UpdateTime"), rsCorrelative("ProductID")) & "'>" & GetSubStr(rsCorrelative("ProductName"), TitleLen, ShowSuspensionPoints) & "</a></li>"
                Else
                    strCorrelative = strCorrelative & "<li><a class='LinkProductCorrelative' href='" & GetProductUrl("", "", "", rsCorrelative("ProductID")) & "'>" & GetSubStr(rsCorrelative("ProductName"), TitleLen, ShowSuspensionPoints) & "</a></li>"
                End If
                rsCorrelative.MoveNext
            Loop
        Else
            strCorrelative = "<table width='100%' cellpadding='3' cellspacing='1' border='0' align='center' class='tdbg'><tr valign='top'>"
            iNum = 0
            Do While Not rsCorrelative.EOF
                If UseCreateHTML > 0 Then
                    strLink = "<a class='LinkProductCorrelative' href='" & GetProductUrl(rsCorrelative("ParentDir"), rsCorrelative("ClassDir"), rsCorrelative("UpdateTime"), rsCorrelative("ProductID")) & "'"
                Else
                    strLink = "<a class='LinkProductCorrelative' href='" & GetProductUrl("", "", "", rsCorrelative("ProductID")) & "'"
                End If
                strLink = strLink & " target='_blank'>"
    
                strCorrelative = strCorrelative & "<td class='tdbg3'>"
                strCorrelative = strCorrelative & "<table cellspacing='2' border='0'>"
                strCorrelative = strCorrelative & "<tr><td align='center'>" & strLink & GetProductThumb(rsCorrelative("ProductThumb"), ImgWidth, ImgHeight, 1) & "</a></td></tr>"
                strCorrelative = strCorrelative & "<tr><td align='left'>" & strLink & GetSubStr(rsCorrelative("ProductName"), TitleLen, ShowSuspensionPoints) & "</a><br>" & GetProductPrice(0, False, rsCorrelative("ProductType"), rsCorrelative("Price_Original"), rsCorrelative("Price"), rsCorrelative("Price_Market"), rsCorrelative("Price_Member"), rsCorrelative("BeginDate"), rsCorrelative("EndDate"), rsCorrelative("Discount")) & "</td></tr>"
                strCorrelative = strCorrelative & "</table></td>"
                rsCorrelative.MoveNext
                iNum = iNum + 1
                If ((iNum Mod Cols = 0) And (Not rsCorrelative.EOF)) Then strCorrelative = strCorrelative & "</tr><tr valign='top'>"
            Loop
            strCorrelative = strCorrelative & "</tr></table>"
        End If
    End If
    rsCorrelative.Close
    Set rsCorrelative = Nothing
    GetCorrelative = strCorrelative
End Function


Private Function GetStocks()
    If UseCreateHTML > 0 Then
        GetStocks = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetStocks.asp?ProductID=" & ProductID & "'></script>"
    Else
        GetStocks = rsProduct("Stocks") - rsProduct("OrderNum")
    End If
End Function

Private Function GetHits()
    If UseCreateHTML > 0 Then
        GetHits = "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetHits.asp?ProductID=" & ProductID & "'></script>"
    Else
        GetHits = rsProduct("Hits")
    End If
End Function

Private Function GetProductProperty()
    Dim strProperty
    If rsProduct("OnTop") = True Then
        strProperty = strProperty & XmlText_Class("ShowProduct/OnTop", "<font color=blue>��</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsProduct("IsHot") = True Then
        strProperty = strProperty & XmlText_Class("ShowProduct/Hot", "<font color=red>��</font>&nbsp;")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;&nbsp;"
    End If
    If rsProduct("IsElite") = True Then
        strProperty = strProperty & XmlText_Class("ShowProduct/Elite", "<font color=green>��</font>")
    Else
        strProperty = strProperty & "&nbsp;&nbsp;"
    End If
    GetProductProperty = strProperty
End Function

Private Function GetStars(Stars)
    Dim strStars
    strStars = "<font color='" & XmlText_Class("ShowProduct/Star_Color", "#009900") & "'>" & String(Stars, XmlText_Class("ShowProduct/Star", "��")) & "</font>"
    GetStars = strStars
End Function

Private Function GetKeywords(strSplit, strKeyword)
    Dim strTemp
    strTemp = PE_Replace(Mid(strKeyword, 2, Len(strKeyword) - 2), "|", strSplit)
    GetKeywords = strTemp
End Function

Private Function GetProducerInfo(tmpName, iType)
    If IsNull(tmpName) Or IsNull(iType) Then
        GetProducerInfo = tmpName
    Else
        GetProducerInfo = "<a href='" & ChannelUrl_ASPFile & "/Show" & iType & ".asp?ChannelID=1000&" & iType & "Name=" & tmpName & "'>" & tmpName & "</a>"
    End If
End Function


Public Function GetCustomFromTemplate(strValue)   '�õ��Զ����б�İ�����Ƶ�HTML����
    Dim strCustom, strParameter
    strCustom = strValue
    regEx.Pattern = "��ProductList\((.*?)\)��([\s\S]*?)��\/ProductList��"
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
    regEx.Pattern = "\{\$GetProductList\((.*?)\)\}"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        If FoundErr = True Then
            strList = PE_Replace(strList, Match.Value, ErrMsg)
        Else
            strList = PE_Replace(strList, Match.Value, GetListFromLabel(Match.SubMatches(0)))
        End If
    Next
    GetListFromTemplate = strList
End Function

Public Function GetPicFromTemplate(ByVal strValue)
    Dim strPicList
    strPicList = strValue
    regEx.Pattern = "\{\$GetPicProduct\((.*?)\)\}"
    Set Matches = regEx.Execute(strPicList)
    For Each Match In Matches
        If FoundErr = True Then
            strPicList = PE_Replace(strPicList, Match.Value, ErrMsg)
        Else
            strPicList = PE_Replace(strPicList, Match.Value, GetPicFromLabel(Match.SubMatches(0)))
        End If
    Next
    GetPicFromTemplate = strPicList
End Function

Public Function GetSlidePicFromTemplate(ByVal strValue)
    Dim strSlidePic, InitSlideJS
    InitSlideJS = False
    strSlidePic = strValue
    regEx.Pattern = "\{\$GetSlidePicProduct\((.*?)\)\}"
    Set Matches = regEx.Execute(strSlidePic)
    For Each Match In Matches
        If FoundErr = True Then
            strSlidePic = PE_Replace(strSlidePic, Match.Value, ErrMsg)
        Else
            If InitSlideJS = False Then
                strSlidePic = PE_Replace(strSlidePic, Match.Value, JS_SlidePic & GetSlidePicFromLabel(Match.SubMatches(0)))
                InitSlideJS = True
            Else
                strSlidePic = PE_Replace(strSlidePic, Match.Value, GetSlidePicFromLabel(Match.SubMatches(0)))
            End If
        End If
    Next
    GetSlidePicFromTemplate = strSlidePic
End Function

Private Function GetSlidePicFromLabel(ByVal strSource)
    Dim strTemp, arrTemp, tChannelID, arrClassID, tSpecialID
    If strSource = "" Then
        GetSlidePicFromLabel = ""
        Exit Function
    End If
    
    strTemp = Replace(strSource, Chr(34), "")
    arrTemp = Split(strTemp, ",")
    
    Select Case Trim(arrTemp(0))
    Case "arrChildID"
        arrClassID = arrChildID
    Case "ClassID"
        arrClassID = ClassID
    Case Else
        arrClassID = arrTemp(0)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")

    Select Case Trim(arrTemp(2))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(2))
    End Select
    
    Select Case UBound(arrTemp)
    Case 11
        GetSlidePicFromLabel = GetSlidePicProduct(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, PE_CLng(arrTemp(3)), PE_CBool(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), -1)
    Case 12
        GetSlidePicFromLabel = GetSlidePicProduct(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, PE_CLng(arrTemp(3)), PE_CBool(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)))
    Case Else
        GetSlidePicFromLabel = "����ʽ��ǩ��{$GetSlidePicProduct(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
    End Select
End Function

Private Function GetPicFromLabel(ByVal strSource)
    Dim strTemp, arrTemp, tChannelID, arrClassID, tSpecialID
    If strSource = "" Then
        GetPicFromLabel = ""
        Exit Function
    End If
    
    strTemp = Replace(strSource, Chr(34), "")
    arrTemp = Split(strTemp, ",")
    
    If UBound(arrTemp) <> 13 And UBound(arrTemp) <> 18 Then
        GetPicFromLabel = "����ʽ��ǩ��{$GetPicProduct(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If
    
    Select Case Trim(arrTemp(0))
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
        arrClassID = arrTemp(0)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")

    Select Case Trim(arrTemp(2))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(2))
    End Select
    
    Select Case UBound(arrTemp)
    Case 13
        GetPicFromLabel = GetPicProduct(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), 0, 0, False, 4, 0, 1)
    Case 18
        GetPicFromLabel = GetPicProduct(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), 0, PE_CLng(arrTemp(14)), PE_CBool(arrTemp(15)), PE_CLng(arrTemp(16)), PE_CLng(arrTemp(17)), PE_CLng(arrTemp(18)))
    End Select
End Function

Private Function GetListFromLabel(ByVal str1)
    Dim strTemp, arrTemp
    Dim tChannelID, ProductNum, arrClassID, tSpecialID, OrderType, OpenType
    If str1 = "" Then
        GetListFromLabel = ""
        Exit Function
    End If
    
    strTemp = Replace(str1, Chr(34), "")
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) < 18 Or UBound(arrTemp) > 39 Then
        GetListFromLabel = "����ʽ��ǩ��{$GetProductList(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If

    Select Case Trim(arrTemp(0))
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
        arrClassID = arrTemp(0)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")
    Select Case Trim(arrTemp(2))
    Case "SpecialID"
        tSpecialID = SpecialID
    Case Else
        tSpecialID = PE_CLng(arrTemp(2))
    End Select
    
    Select Case Trim(arrTemp(3))
    Case "rsClass_TopNumber"
        ProductNum = 8
    Case "TopNumber"
        ProductNum = 8
    Case Else
        ProductNum = PE_CLng(arrTemp(3))
    End Select
    

    Select Case Trim(arrTemp(8))
    Case "rsClass_ItemListOrderType"
        OrderType = rsClass("ItemListOrderType")
    Case "ItemListOrderType"
        OrderType = ItemListOrderType
    Case Else
        OrderType = PE_CLng(arrTemp(8))
    End Select

    Select Case Trim(arrTemp(18))
    Case "rsClass_ItemOpenType"
        OpenType = rsClass("ItemOpenType")
    Case "ItemOpenType"
        OpenType = ItemOpenType
    Case Else
        OpenType = PE_CLng(arrTemp(18))
    End Select

    Select Case UBound(arrTemp)
    Case 18
        GetListFromLabel = GetProductList(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, ProductNum, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), OrderType, PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CBool(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CBool(arrTemp(17)), OpenType, 0, 0, 0, False, "", False, False, False, 0, False, False, False, False, False, False, 0, 0, "", "", "", "", "")
    Case 19
        GetListFromLabel = GetProductList(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, ProductNum, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), OrderType, PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CBool(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CBool(arrTemp(17)), OpenType, 0, PE_CLng(arrTemp(19)), 0, False, "", False, False, False, 0, False, False, False, False, False, False, 0, 0, "", "", "", "", "")
    Case 20
        GetListFromLabel = GetProductList(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, ProductNum, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), OrderType, PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CBool(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CBool(arrTemp(17)), OpenType, 0, PE_CLng(arrTemp(19)), PE_CLng(arrTemp(20)), False, "", False, False, False, 0, False, False, False, False, False, False, 0, 0, "", "", "", "", "")
    Case 34
        GetListFromLabel = GetProductList(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, ProductNum, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), OrderType, PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CBool(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CBool(arrTemp(17)), OpenType, 0, PE_CLng(arrTemp(19)), PE_CLng(arrTemp(20)), PE_CBool(arrTemp(21)), Trim(arrTemp(22)), PE_CBool(arrTemp(23)), PE_CBool(arrTemp(24)), PE_CBool(arrTemp(25)), PE_CLng(arrTemp(26)), PE_CBool(arrTemp(27)), PE_CBool(arrTemp(28)), PE_CBool(arrTemp(29)), PE_CBool(arrTemp(30)), PE_CBool(arrTemp(31)), PE_CBool(arrTemp(32)), PE_CLng(arrTemp(33)), PE_CLng(arrTemp(34)), "", "", "", "", "")
    Case 39
        GetListFromLabel = GetProductList(arrClassID, PE_CBool(arrTemp(1)), tSpecialID, ProductNum, PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CLng(arrTemp(7)), OrderType, PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CBool(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CBool(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CBool(arrTemp(17)), OpenType, 0, PE_CLng(arrTemp(19)), PE_CLng(arrTemp(20)), PE_CBool(arrTemp(21)), Trim(arrTemp(22)), PE_CBool(arrTemp(23)), PE_CBool(arrTemp(24)), PE_CBool(arrTemp(25)), PE_CLng(arrTemp(26)), PE_CBool(arrTemp(27)), PE_CBool(arrTemp(28)), PE_CBool(arrTemp(29)), PE_CBool(arrTemp(30)), PE_CBool(arrTemp(31)), PE_CBool(arrTemp(32)), PE_CLng(arrTemp(33)), PE_CLng(arrTemp(34)), Trim(arrTemp(35)), Trim(arrTemp(36)), Trim(arrTemp(37)), Trim(arrTemp(38)), Trim(arrTemp(39)))
    Case Else
        GetListFromLabel = "����ʽ��ǩ��{$GetProductList(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
    End Select
    
End Function

Private Function GetCustomFromLabel(strTemp, strList)
    Dim arrTemp
    Dim strProductThumb, arrPicTemp
    Dim arrClassID, IncludeChild, iSpecialID, ProductType, ProductNum, IsHot, IsElite, Author, DateNum, OrderType, UsePage, TitleLen, ContentLen
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
    If UBound(arrTemp) <> 11 and UBound(arrTemp) <> 12 Then
        GetCustomFromLabel = "�Զ����б��ǩ����ProductList(�����б�)���б����ݡ�/ProductList���Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If
        
    Select Case Trim(arrTemp(0))
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
        arrClassID = arrTemp(0)
    End Select
    arrClassID = Replace(Trim(arrClassID), "|", ",")

    IncludeChild = PE_CBool(arrTemp(1))
    
    Select Case Trim(arrTemp(2))
    Case "SpecialID"
        iSpecialID = SpecialID
    Case Else
        iSpecialID = PE_CLng(arrTemp(2))
    End Select
    
    ProductNum = PE_CLng(arrTemp(3))
    ProductType = PE_CLng(arrTemp(4))
    IsHot = PE_CBool(arrTemp(5))
    IsElite = PE_CBool(arrTemp(6))
    DateNum = PE_CLng(arrTemp(7))
    
    Select Case Trim(arrTemp(8))
    Case "rsClass_ItemListOrderType"
        OrderType = rsClass("ItemListOrderType")
    Case "ItemListOrderType"
        OrderType = ItemListOrderType
    Case Else
        OrderType = PE_CLng(arrTemp(8))
    End Select
    
    UsePage = PE_CBool(arrTemp(9))
    
    TitleLen = PE_CLng(arrTemp(10))
    ContentLen = PE_CLng(arrTemp(11))
    If UBound(arrTemp) = 12  then
        IncludePic = PE_CBool(arrTemp(12))
    Else
        IncludePic = False	    
    End If        
    Dim rsField, ArrField, iField
    Set rsField = Conn.Execute("select FieldName,LabelName,FieldType from PE_Field where ChannelID=-5 or ChannelID=" & ChannelID & "")
    If Not (rsField.BOF And rsField.EOF) Then
        ArrField = rsField.getrows(-1)
    End If
    Set rsField = Nothing

    Dim sqlCustom, rsCustom, iCount, strCustomList, strThisClass, strLink
    iCount = 0
    sqlCustom = ""
    strThisClass = ""
    strCustomList = ""
    
    If ProductNum > 0 Then
        sqlCustom = "select top " & ProductNum & " "
    Else
        sqlCustom = "select "
    End If
    If ContentLen > 0 Then
        sqlCustom = sqlCustom & "P.ProductExplain,"
    End If
    If IsArray(ArrField) Then
        For iField = 0 To UBound(ArrField, 2)
            sqlCustom = sqlCustom & "P." & ArrField(0, iField) & ","
        Next
    End If
    sqlCustom = sqlCustom & "P.ProductID,P.ClassID,P.ProductNum,P.LimitNum,P.ProductName,P.UpdateTime,P.ProductThumb,P.ProductIntro,P.Hits"
    sqlCustom = sqlCustom & ",P.IsHot,P.IsElite,P.OnTop,P.ProductModel,P.ProductStandard,P.ProducerName,P.TrademarkName"
    sqlCustom = sqlCustom & ",P.Unit,P.Stocks,P.OrderNum,P.BarCode,P.Stars,P.PresentExp,P.PresentMoney,P.PresentPoint,P.Inputer"
    sqlCustom = sqlCustom & ",P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,P.BeginDate,P.EndDate,P.Discount"
    sqlCustom = sqlCustom & GetSqlStr(arrClassID, IncludeChild, iSpecialID, ProductType, IsHot, IsElite, DateNum, OrderType, True, IncludePic)
    Set rsCustom = Server.CreateObject("ADODB.Recordset")
    rsCustom.Open sqlCustom, Conn, 1, 1
    If rsCustom.BOF And rsCustom.EOF Then
        totalPut = 0
        If IsHot = False And IsElite = False Then
            strCustomList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & ChannelShortName & "</li>"
        ElseIf IsHot = True And IsElite = False Then
            strCustomList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & XmlText_Class("ProductList/t2", "����") & ChannelShortName & "</li>"
        ElseIf IsHot = False And IsElite = True Then
            strCustomList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & XmlText_Class("ProductList/t3", "�Ƽ�") & ChannelShortName & "</li>"
        Else
            strCustomList = "<li>" & strThisClass & XmlText_Class("ProductList/t1", "û��") & XmlText_Class("ProductList/t2", "����") & XmlText_Class("ProductList/t3", "�Ƽ�") & ChannelShortName & "</li>"
        End If
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
        If InStr(strTemp, "{$ClassUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ClassUrl}", GetClassUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("ClassID"), 0))

        If InStr(strTemp, "{$ProductUrl}") > 0 Then strTemp = PE_Replace(strTemp, "{$ProductUrl}", GetProductUrl(rsCustom("ParentDir"), rsCustom("ClassDir"), rsCustom("UpdateTime"), rsCustom("ProductID")))
        strTemp = PE_Replace(strTemp, "{$ProductID}", rsCustom("ProductID"))
        strTemp = PE_Replace(strTemp, "{$ProductNum}", rsCustom("ProductNum"))
        strTemp = PE_Replace(strTemp, "{$BarCode}", rsCustom("BarCode"))
        If InStr(strTemp, "{$UpdateDate}") > 0 Then strTemp = PE_Replace(strTemp, "{$UpdateDate}", FormatDateTime(rsCustom("UpdateTime"), 2))
        strTemp = PE_Replace(strTemp, "{$UpdateTime}", rsCustom("UpdateTime"))
        strTemp = PE_Replace(strTemp, "{$Stars}", GetStars(rsCustom("Stars")))
        strTemp = PE_Replace(strTemp, "{$ProducerName}", rsCustom("ProducerName"))
        strTemp = PE_Replace(strTemp, "{$TrademarkName}", rsCustom("TrademarkName"))
        strTemp = PE_Replace(strTemp, "{$ProductModel}", rsCustom("ProductModel"))
        strTemp = PE_Replace(strTemp, "{$ProductStandard}", rsCustom("ProductStandard"))
        strTemp = PE_Replace(strTemp, "{$Hits}", rsCustom("Hits"))
        strTemp = PE_Replace(strTemp, "{$Inputer}", rsCustom("Inputer"))
        strTemp = PE_Replace(strTemp, "{$Unit}", rsCustom("Unit"))
        strTemp = PE_Replace(strTemp, "{$Stocks}", rsCustom("Stocks") - rsCustom("OrderNum"))

        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "OnTop")
        ElseIf rsCustom("IsElite") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Elite")
        ElseIf rsCustom("IsHot") = True Then
            strTemp = PE_Replace(strTemp, "{$Property}", "Hot")
        Else
            strTemp = PE_Replace(strTemp, "{$Property}", "Common")
        End If
        If rsCustom("OnTop") = True Then
            strTemp = PE_Replace(strTemp, "{$Top}", strTop2)
        Else
            strTemp = PE_Replace(strTemp, "{$Top}", "")
        End If
        If rsCustom("IsElite") = True Then
            strTemp = PE_Replace(strTemp, "{$Elite}", strElite2)
        Else
            strTemp = PE_Replace(strTemp, "{$Elite}", "")
        End If
        If rsCustom("IsHot") = True Then
            strTemp = PE_Replace(strTemp, "{$Hot}", strHot2)
        Else
            strTemp = PE_Replace(strTemp, "{$Hot}", "")
        End If
        
        If TitleLen > 0 Then
            strTemp = PE_Replace(strTemp, "{$ProductName}", GetSubStr(rsCustom("ProductName"), TitleLen, ShowSuspensionPoints))
        Else
            strTemp = PE_Replace(strTemp, "{$ProductName}", rsCustom("ProductName"))
        End If
        strTemp = PE_Replace(strTemp, "{$ProductNameOriginal}", rsCustom("ProductName"))
        strTemp = PE_Replace(strTemp, "{$ProductIntro}", rsCustom("ProductIntro"))
        If ContentLen > 0 Then
            If InStr(strTemp, "{$ProductExplain}") > 0 Then strTemp = PE_Replace(strTemp, "{$ProductExplain}", Left(nohtml(rsCustom("ProductExplain")), ContentLen))
        Else
            strTemp = PE_Replace(strTemp, "{$ProductExplain}", "")
        End If

        If InStr(strTemp, "{$ProductThumb}") > 0 Then strTemp = Replace(strTemp, "{$ProductThumb}", GetProductThumb(rsCustom("ProductThumb"), 130, 0, 0))
        If InStr(strTemp, "{$JsProductThumb}") > 0 Then strTemp = Replace(strTemp, "{$JsProductThumb}", GetJsProductThumb(rsCustom("ProductThumb"), 130, 0, 0))
        '�滻��ҳͼƬ
        regEx.Pattern = "\{\$ProductThumb\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrPicTemp = Split(Match.SubMatches(0), ",")
            strProductThumb = GetProductThumb(Trim(rsCustom("ProductThumb")), PE_CLng(arrPicTemp(0)), PE_CLng(arrPicTemp(1)), 0)
            strTemp = Replace(strTemp, Match.Value, strProductThumb)
        Next
        Dim arrJsPicTemp , strJsProductThumb
        regEx.Pattern = "\{\$JsProductThumb\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrJsPicTemp = Split(Match.SubMatches(0), ",")
            strJsProductThumb = GetJsProductThumb(Trim(rsCustom("ProductThumb")), PE_CLng(arrJsPicTemp(0)), PE_CLng(arrJsPicTemp(1)), 0)
            strTemp = Replace(strTemp, Match.Value, strJsProductThumb)
        Next
        If InStr(strTemp, "{$Price_Original}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price_Original}", GetPrice_FilterZero_NoSymbol(rsCustom("Price_Original")))
        If rsCustom("Price_Market") > 0 Then
            strTemp = PE_Replace(strTemp, "{$Price_Market}", rsCustom("Price_Market"))
        Else
            If InStr(strTemp, "{$Price_Market}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price_Market}", GetPrice_Market_NoSymbol(rsCustom("Price_Original")))
        End If
        
        Select Case rsCustom("ProductType")
        Case 1
            strTemp = Replace(strTemp, "{$ProductTypeName}", "����������Ʒ")
            If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price_Original")))
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", "")
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        Case 2
            If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price")))
            strTemp = Replace(strTemp, "{$ProductTypeName}", "�Ǽ���Ʒ")
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", "")
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        Case 3
            If rsCustom("BeginDate") <= Date And rsCustom("EndDate") >= Date Then
                strTemp = Replace(strTemp, "{$ProductTypeName}", "�ؼ۴�����Ʒ")
                If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price")))
                strTemp = PE_Replace(strTemp, "{$BeginDate}", rsCustom("BeginDate"))
                strTemp = PE_Replace(strTemp, "{$EndDate}", rsCustom("EndDate"))
                strTemp = PE_Replace(strTemp, "{$Discount}", rsCustom("Discount"))
                strTemp = PE_Replace(strTemp, "{$LimitNum}", rsCustom("LimitNum"))
            Else
                strTemp = Replace(strTemp, "{$ProductTypeName}", "����������Ʒ")
                If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price_Original")))
                strTemp = Replace(strTemp, "{$BeginDate}", "")
                strTemp = Replace(strTemp, "{$EndDate}", "")
                strTemp = Replace(strTemp, "{$Discount}", "")
                strTemp = Replace(strTemp, "{$LimitNum}", "")
            End If
        Case 4
            strTemp = Replace(strTemp, "{$ProductTypeName}", "������Ʒ�����������ۣ�")
            strTemp = PE_Replace(strTemp, "{$Price}", rsCustom("Price"))
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", "")
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        Case 5
            strTemp = Replace(strTemp, "{$ProductTypeName}", "���۴���")
            If InStr(strTemp, "{$Price}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price}", GetPrice_FilterZero_NoSymbol(rsCustom("Price")))
            strTemp = Replace(strTemp, "{$BeginDate}", "")
            strTemp = Replace(strTemp, "{$EndDate}", "")
            strTemp = Replace(strTemp, "{$Discount}", rsCustom("Discount"))
            strTemp = Replace(strTemp, "{$LimitNum}", "")
        End Select
        If InStr(strTemp, "{$Price_Member}") > 0 Then strTemp = PE_Replace(strTemp, "{$Price_Member}", GetPrice_Member_NoSymbol(rsCustom("Price_Member")))
        strTemp = PE_Replace(strTemp, "{$PresentExp}", rsCustom("PresentExp"))
        strTemp = PE_Replace(strTemp, "{$PresentPoint}", rsCustom("PresentPoint"))
        strTemp = PE_Replace(strTemp, "{$PresentMoney}", rsCustom("PresentMoney"))
        strTemp = PE_Replace(strTemp, "{$PointName}", PointName)
        strTemp = PE_Replace(strTemp, "{$PointUnit}", PointUnit)
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


Private Function GetPrice()
    Dim dblTruePrice, dblTempPrice
    dblTruePrice = 0
    UserLogined = CheckUserLogined()
    If UserLogined = True Then GetUser (username)
    Dim rs
    Set rs = Conn.Execute("select ProductType,Price,Price_Original,Price_Member,Price_Agent,BeginDate,EndDate from PE_Product where ProductID=" & ProductID & "")
    If Not (rs.BOF And rs.EOF) Then
        Select Case GroupType
        Case 0, 1 'δ��¼
            Select Case rs("ProductType")
            Case 1, 2, 4, 5
                dblTruePrice = rs("Price")
            Case 3
                If Date < rs("BeginDate") Or Date > rs("EndDate") Then
                    dblTruePrice = rs("Price_Original")
                Else
                    dblTruePrice = rs("Price")
                End If
            End Select
        Case 2, 3   'ע���Ա
            Select Case rs("ProductType")
            Case 1, 2
                If rs("Price_Member") > 0 Then '���ָ���˻�Ա��
                    dblTruePrice = rs("Price_Member")
                Else
                    dblTruePrice = rs("Price") * Discount_Member / 100
                End If
            Case 3
                If Date < rs("BeginDate") Or Date > rs("EndDate") Then
                    dblTempPrice = rs("Price_Original")
                Else
                    dblTempPrice = rs("Price")
                End If
                If rs("Price_Member") > 0 Then '���ָ���˻�Ա��
                    If rs("Price_Member") <= dblTempPrice Then
                        dblTruePrice = rs("Price_Member")
                    Else
                        dblTruePrice = dblTempPrice
                    End If
                Else
                    If PE_CLng(UserSetting(12)) = 1 Then '����������������Ż�
                        dblTruePrice = dblTempPrice * Discount_Member / 100
                    Else
                        If rs("Price_Original") * Discount_Member / 100 >= dblTempPrice Then
                            dblTruePrice = dblTempPrice
                        Else
                            dblTruePrice = rs("Price_Original") * Discount_Member / 100
                        End If
                    End If
                End If
            Case 4
                dblTruePrice = rs("Price")
            Case 5
                dblTempPrice = rs("Price")
                If rs("Price_Member") > 0 Then '���ָ���˻�Ա��
                    If rs("Price_Member") <= dblTempPrice Then
                        dblTruePrice = rs("Price_Member")
                    Else
                        dblTruePrice = dblTempPrice
                    End If
                Else
                    If PE_CLng(UserSetting(12)) = 1 Then '����������������Ż�
                        dblTruePrice = dblTempPrice * Discount_Member / 100
                    Else
                        If rs("Price_Original") * Discount_Member / 100 >= dblTempPrice Then
                            dblTruePrice = dblTempPrice
                        Else
                            dblTruePrice = rs("Price_Original") * Discount_Member / 100
                        End If
                    End If
                End If
            End Select
        Case 4  '������
            dblTempPrice = rs("Price")
            If rs("Price_Agent") > 0 Then '���ָ���˴����
                dblTruePrice = rs("Price_Agent")
            Else
                If Discount_Member = 100 Then
                    dblTruePrice = dblTempPrice
                Else
                    If PE_CLng(UserSetting(12)) = 1 Then '����������������Ż�
                        dblTruePrice = dblTempPrice * Discount_Member / 100
                    Else
                        If rs("Price_Original") * Discount_Member / 100 <= dblTempPrice Then
                            dblTruePrice = rs("Price_Original") * Discount_Member / 100
                        Else
                            dblTruePrice = dblTempPrice
                        End If
                    End If
                End If
            End If
        End Select
    End If
    GetPrice = dblTruePrice
End Function

Public Sub GetHtml_Index()
    Dim strTemp, strTopUser, strFriendSite, arrTemp, strAnnounce, strPopAnnouce, iCols, iClassID
    Dim ProductList_ChildClass, ProductList_ChildClass2

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
    regEx.Pattern = "��ProductList_ChildClass��([\s\S]*?)��\/ProductList_ChildClass��"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        ProductList_ChildClass = Match.SubMatches(0)
        strHtml = regEx.Replace(strHtml, "{$ProductList_ChildClass}")
                
        '�õ�ÿ����ʾ������
        iCols = 1
        regEx.Pattern = "��Cols=(\d{1,2})��"
        Set Matches2 = regEx.Execute(ProductList_ChildClass)
        ProductList_ChildClass = regEx.Replace(ProductList_ChildClass, "")
        For Each Match2 In Matches2
            If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
        Next
        
        ProductList_ChildClass2 = ""
        
        '��ʼѭ�����õ���������Ŀ�б��HTML����
        iClassID = 0
        Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=0 and ShowOnIndex=" & PE_True & " order by RootID")
        Do While Not rsClass.EOF
            strTemp = ProductList_ChildClass
            
            strTemp = GetCustomFromTemplate(strTemp)
            strTemp = GetListFromTemplate(strTemp)
            strTemp = GetPicFromTemplate(strTemp)
            
            strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), 0))
            strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", rsClass("Readme"))
            strTemp = PE_Replace(strTemp, "{$rsClass_Tips}", rsClass("Tips"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", rsClass("ClassID"))
            strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", rsClass("ClassName"))
            strTemp = Replace(strTemp, "{$ShowClassAD}", "")
            
            rsClass.MoveNext
            iClassID = iClassID + 1
            If iClassID Mod iCols = 0 And Not rsClass.EOF Then
                ProductList_ChildClass2 = ProductList_ChildClass2 & strTemp
                If iCols > 1 Then ProductList_ChildClass2 = ProductList_ChildClass2 & "</tr><tr>"
            Else
                ProductList_ChildClass2 = ProductList_ChildClass2 & strTemp
                If iCols > 1 Then ProductList_ChildClass2 = ProductList_ChildClass2 & "<td width='1'></td>"
            End If
        Loop
        rsClass.Close
        Set rsClass = Nothing

        strHtml = Replace(strHtml, "{$ProductList_ChildClass}", ProductList_ChildClass2)

    Next
    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
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
    
    Dim ProductList_CurrentClass, ProductList_CurrentClass2, ProductList_ChildClass, ProductList_ChildClass2
    If Child > 0 And ClassShowType <> 2 Then    '�����ǰ��Ŀ������Ŀ
        ItemCount = PE_CLng(Conn.Execute("select Count(*) from PE_Product where ClassID=" & ClassID & "")(0))
        If ItemCount <= 0 Then     '�����ǰ��Ŀû�в�Ʒ
            regEx.Pattern = "��ProductList_CurrentClass��([\s\S]*?)��\/ProductList_CurrentClass��"
            strHtml = regEx.Replace(strHtml, "") '��ȥ����ʾ��ǰ��Ŀ��ֻ���ڱ���Ŀ�Ĳ�Ʒ�б�����
        Else      '�����ǰ��Ŀ������Ŀ���ҵ�ǰ��Ŀ�в�Ʒ������Ҫ��ʾ������
            regEx.Pattern = "��ProductList_CurrentClass��([\s\S]*?)��\/ProductList_CurrentClass��"
            Set Matches = regEx.Execute(strHtml)
            For Each Match In Matches
                ProductList_CurrentClass = Match.SubMatches(0)
                strHtml = regEx.Replace(strHtml, "{$ProductList_CurrentClass}")
                
                strTemp = ProductList_CurrentClass
                strTemp = GetCustomFromTemplate(strTemp)
                strTemp = GetListFromTemplate(strTemp)
                strTemp = GetPicFromTemplate(strTemp)
                
                strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, 0))
                strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", ReadMe)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", ClassName)
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", ClassID)
                
                strHtml = Replace(strHtml, "{$ProductList_CurrentClass}", strTemp)
            Next
        End If
        
        '�õ�����Ŀ�б�İ�����Ƶ�HTML����
        regEx.Pattern = "��ProductList_ChildClass��([\s\S]*?)��\/ProductList_ChildClass��"
        Set Matches = regEx.Execute(strHtml)
        For Each Match In Matches
            ProductList_ChildClass = Match.SubMatches(0)
            strHtml = regEx.Replace(strHtml, "{$ProductList_ChildClass}")
            
            '�õ�ÿ����ʾ������
            iCols = 1
            regEx.Pattern = "��Cols=(\d{1,2})��"
            Set Matches2 = regEx.Execute(ProductList_ChildClass)
            ProductList_ChildClass = regEx.Replace(ProductList_ChildClass, "")
            For Each Match2 In Matches2
                If Match2.SubMatches(0) > 1 Then iCols = Match2.SubMatches(0)
            Next
            
            '��ʼѭ�����õ���������Ŀ�б��HTML����
            iClassID = 0
            Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and ClassType=1 and ParentID=" & ClassID & " and IsElite=" & PE_True & " and ClassType=1 order by RootID,OrderID")
            Do While Not rsClass.EOF
                strTemp = ProductList_ChildClass
                
                strTemp = GetCustomFromTemplate(strTemp)
                strTemp = GetListFromTemplate(strTemp)
                strTemp = GetPicFromTemplate(strTemp)
                
                strTemp = Replace(strTemp, "{$rsClass_ClassUrl}", GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), 0))
                strTemp = PE_Replace(strTemp, "{$rsClass_Readme}", rsClass("Readme"))
                strTemp = PE_Replace(strTemp, "{$rsClass_Tips}", rsClass("Tips"))
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassName}", rsClass("ClassName"))
                strTemp = PE_Replace(strTemp, "{$rsClass_ClassID}", rsClass("ClassID"))
                strTemp = Replace(strTemp, "{$ShowClassAD}", "")
            
                rsClass.MoveNext
                iClassID = iClassID + 1
                If iClassID Mod iCols = 0 And Not rsClass.EOF Then
                    ProductList_ChildClass2 = ProductList_ChildClass2 & strTemp
                    If iCols > 1 Then ProductList_ChildClass2 = ProductList_ChildClass2 & "</tr><tr>"
                Else
                    ProductList_ChildClass2 = ProductList_ChildClass2 & strTemp
                    If iCols > 1 Then ProductList_ChildClass2 = ProductList_ChildClass2 & "<td width='1'></td>"
                End If
            Loop
            rsClass.Close
            Set rsClass = Nothing

            strHtml = Replace(strHtml, "{$ProductList_ChildClass}", ProductList_ChildClass2)

        Next
    End If

    strHtml = GetCustomFromTemplate(strHtml)
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)
    
    Dim strPath
    strPath = ChannelUrl & GetListPath(StructureType, ListFileType, ParentDir, ClassDir)
    
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$ClassPicUrl}", ClassPicUrl)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = Replace(strHtml, "{$ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, 0))
    strHtml = Replace(strHtml, "{$ClassListUrl}", GetClass_1Url(ParentDir, ClassDir, ClassID, 0))

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

End Sub


Public Sub GetHtml_Product()
    Dim arrTemp
    strHtml = GetCustomFromTemplate(strHtml)  '�����Ƚ����Զ����б��ǩ

    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ProductID}", ProductID)
    Call ReplaceCommonLabel   '����ͨ�ñ�ǩ�������Զ����ǩ
    strHtml = GetCustomFromTemplate(strHtml)  '�����Ƚ����Զ����б��ǩ
    strHtml = GetListFromTemplate(strHtml)
    strHtml = GetPicFromTemplate(strHtml)
    strHtml = GetSlidePicFromTemplate(strHtml)

    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ProductID}", ProductID)
    strHtml = Replace(strHtml, "{$PageTitle}", ReplaceText(ProductName, 2))
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    
    If InStr(strHtml, "{$MY_") > 0 Then
        Dim rsField
        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-5 or ChannelID=" & ChannelID & "")
        Do While Not rsField.EOF
            strHtml = PE_Replace(strHtml, rsField("LabelName"), PE_HTMLEncode(rsProduct(Trim(rsField("FieldName")))))
            rsField.MoveNext
        Loop
        Set rsField = Nothing
    End If
    

    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    If InStr(strHtml, "{$ClassUrl}") > 0 Then strHtml = Replace(strHtml, "{$ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, 0))
    strHtml = CustomContent("Class", Custom_Content_Class, strHtml)
    
    If InStr(strHtml, "{$ProductName}") > 0 Then strHtml = PE_Replace(strHtml, "{$ProductName}", ReplaceText(ProductName, 2))
    strHtml = PE_Replace(strHtml, "{$ProductNum}", rsProduct("ProductNum"))
    strHtml = PE_Replace(strHtml, "{$BarCode}", rsProduct("BarCode"))
    strHtml = PE_Replace(strHtml, "{$ProductModel}", rsProduct("ProductModel"))
    strHtml = PE_Replace(strHtml, "{$ProductStandard}", rsProduct("ProductStandard"))
    If InStr(strHtml, "{$ProducerName}") > 0 Then strHtml = Replace(strHtml, "{$ProducerName}", GetProducerInfo(rsProduct("ProducerName"), "Producer"))
    If InStr(strHtml, "{$TrademarkName}") > 0 Then strHtml = Replace(strHtml, "{$TrademarkName}", GetProducerInfo(rsProduct("TrademarkName"), "Trademark"))
    strHtml = PE_Replace(strHtml, "{$PresentExp}", rsProduct("PresentExp"))
    strHtml = PE_Replace(strHtml, "{$PresentPoint}", rsProduct("PresentPoint"))
    strHtml = PE_Replace(strHtml, "{$PresentMoney}", rsProduct("PresentMoney"))
    strHtml = PE_Replace(strHtml, "{$PointName}", PointName)
    strHtml = PE_Replace(strHtml, "{$PointUnit}", PointUnit)
    If InStr(strHtml, "{$Stocks}") > 0 Then strHtml = Replace(strHtml, "{$Stocks}", GetStocks())
    strHtml = PE_Replace(strHtml, "{$Unit}", rsProduct("Unit"))
    
    If InStr(strHtml, "{$Price_Original}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price_Original}", GetPrice_FilterZero_NoSymbol(rsProduct("Price_Original")))
    If rsProduct("Price_Market") > 0 Then
        strHtml = PE_Replace(strHtml, "{$Price_Market}", rsProduct("Price_Market"))
    Else
        If InStr(strHtml, "{$Price_Market}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price_Market}", GetPrice_Market_NoSymbol(rsProduct("Price_Original")))
    End If
    
    Select Case rsProduct("ProductType")
    Case 1
        strHtml = Replace(strHtml, "{$ProductTypeName}", "����������Ʒ")
        If InStr(strHtml, "{$Price}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price}", GetPrice_FilterZero_NoSymbol(rsProduct("Price_Original")))
        strHtml = Replace(strHtml, "{$BeginDate}", "")
        strHtml = Replace(strHtml, "{$EndDate}", "")
        strHtml = Replace(strHtml, "{$Discount}", "")
        strHtml = Replace(strHtml, "{$LimitNum}", "")
    Case 2
        If InStr(strHtml, "{$Price}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price}", GetPrice_FilterZero_NoSymbol(rsProduct("Price")))
        strHtml = Replace(strHtml, "{$ProductTypeName}", "�Ǽ���Ʒ")
        strHtml = Replace(strHtml, "{$BeginDate}", "")
        strHtml = Replace(strHtml, "{$EndDate}", "")
        strHtml = Replace(strHtml, "{$Discount}", "")
        strHtml = Replace(strHtml, "{$LimitNum}", "")
    Case 3
        If rsProduct("BeginDate") <= Date And rsProduct("EndDate") >= Date Then
            strHtml = Replace(strHtml, "{$ProductTypeName}", "�ؼ۴�����Ʒ")
            If InStr(strHtml, "{$Price}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price}", GetPrice_FilterZero_NoSymbol(rsProduct("Price")))
            strHtml = PE_Replace(strHtml, "{$BeginDate}", rsProduct("BeginDate"))
            strHtml = PE_Replace(strHtml, "{$EndDate}", rsProduct("EndDate"))
            strHtml = PE_Replace(strHtml, "{$Discount}", rsProduct("Discount"))
            strHtml = PE_Replace(strHtml, "{$LimitNum}", rsProduct("LimitNum"))
        Else
            strHtml = Replace(strHtml, "{$ProductTypeName}", "����������Ʒ")
            If InStr(strHtml, "{$Price}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price}", GetPrice_FilterZero_NoSymbol(rsProduct("Price_Original")))
            strHtml = Replace(strHtml, "{$BeginDate}", "")
            strHtml = Replace(strHtml, "{$EndDate}", "")
            strHtml = Replace(strHtml, "{$Discount}", "")
            strHtml = Replace(strHtml, "{$LimitNum}", "")
        End If
    Case 4
        strHtml = Replace(strHtml, "{$ProductTypeName}", "������Ʒ�����������ۣ�")
        strHtml = PE_Replace(strHtml, "{$Price}", rsProduct("Price"))
        strHtml = Replace(strHtml, "{$BeginDate}", "")
        strHtml = Replace(strHtml, "{$EndDate}", "")
        strHtml = Replace(strHtml, "{$Discount}", "")
        strHtml = Replace(strHtml, "{$LimitNum}", "")
    Case 5
        strHtml = Replace(strHtml, "{$ProductTypeName}", "���۴���")
        If InStr(strHtml, "{$Price}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price}", GetPrice_FilterZero_NoSymbol(rsProduct("Price")))
        strHtml = Replace(strHtml, "{$BeginDate}", "")
        strHtml = Replace(strHtml, "{$EndDate}", "")
        strHtml = Replace(strHtml, "{$Discount}", rsProduct("Discount"))
        strHtml = Replace(strHtml, "{$LimitNum}", "")
    End Select
    If InStr(strHtml, "{$Price_Member}") > 0 Then strHtml = PE_Replace(strHtml, "{$Price_Member}", GetPrice_Member_NoSymbol(rsProduct("Price_Member")))
    
    If UseCreateHTML > 0 Then
        strHtml = Replace(strHtml, "{$Price_Your}", "<script language='javascript' src='" & ChannelUrl_ASPFile & "/GetPrice.asp?ProductID=" & ProductID & "'></script>")
    Else
        If InStr(strHtml, "{$Price_Your}") > 0 Then strHtml = Replace(strHtml, "{$Price_Your}", GetPrice())
    End If
    
    If rsProduct("IsHot") = True Then
        strHtml = Replace(strHtml, "{$Hot}", "<img src='" & ChannelUrl & "/images/P_Hot.gif' alt='������Ʒ'>")
    Else
        strHtml = Replace(strHtml, "{$Hot}", "")
    End If
    If rsProduct("IsElite") = True Then
        strHtml = Replace(strHtml, "{$Elite}", "<img src='" & ChannelUrl & "/images/P_Elite.gif' alt='�Ƽ���Ʒ'>")
    Else
        strHtml = Replace(strHtml, "{$Elite}", "")
    End If
    If rsProduct("OnTop") = True Then
        strHtml = Replace(strHtml, "{$OnTop}", "<img src='" & ChannelUrl & "/images/P_OnTop.gif' alt='�̶���Ʒ'>")
    Else
        strHtml = Replace(strHtml, "{$OnTop}", "")
    End If
    If InStr(strHtml, "{$SalePromotion}") > 0 Then
        Dim strSalePromotion
        Select Case rsProduct("SalePromotionType")
        Case 0
            strSalePromotion = "������"
        Case 1
            strSalePromotion = "�� <b>" & rsProduct("MinNumber") & "</b> " & rsProduct("Unit") & "�� <b>" & rsProduct("PresentNumber") & "</b> " & rsProduct("Unit") & "ͬ����Ʒ"
        Case 2
            Dim rsPresent
            Set rsPresent = Conn.Execute("select ProductID,ProductName,Unit,Price_Original,Price from PE_Product where ProductNum='" & rsProduct("PresentID") & "' and ProductType=4")
            If Not (rsPresent.BOF And rsPresent.EOF) Then
                If rsPresent("Price") > 0 Then
                    strSalePromotion = "�� <b>" & rsProduct("MinNumber") & "</b> " & rsProduct("Unit") & "���Գ�ֵ���� <b>" & rsProduct("PresentNumber") & "</b> " & rsPresent("Unit") & "<a href='" & ChannelUrl_ASPFile & "/ShowProduct.asp?ProductID=" & rsPresent("ProductID") & "' target='_blank'>" & rsPresent("ProductName") & "��ԭ�ۣ�<STRIKE>��" & rsPresent("Price_Original") & "</STRIKE>�������ۣ���" & rsPresent("Price") & "��</a>"
                Else
                    strSalePromotion = "�� <b>" & rsProduct("MinNumber") & "</b> " & rsProduct("Unit") & "�� <b>" & rsProduct("PresentNumber") & "</b> " & rsPresent("Unit") & rsPresent("ProductName")
                End If
            Else
                strSalePromotion = "������"
            End If
            Set rsPresent = Nothing
        Case 3
            strSalePromotion = "����� <b>" & rsProduct("PresentNumber") & "</b> " & rsProduct("Unit") & "ͬ����Ʒ"
        Case 4
            Dim rsPresent1
            Set rsPresent1 = Conn.Execute("select ProductID,ProductName,Unit,Price_Original,Price from PE_Product where ProductNum='" & rsProduct("PresentID") & "' and ProductType=4")
            If Not (rsPresent1.BOF And rsPresent1.EOF) Then
                If rsPresent1("Price") > 0 Then
                    strSalePromotion = "�����Ʒ���Գ�ֵ���� <b>" & rsProduct("PresentNumber") & "</b> " & rsPresent1("Unit") & "<a href='" & ChannelUrl_ASPFile & "/ShowProduct.asp?ProductID=" & rsPresent1("ProductID") & "' target='_blank'>" & rsPresent1("ProductName") & "��ԭ�ۣ�<STRIKE>��" & rsPresent1("Price_Original") & "</STRIKE>�������ۣ���" & rsPresent1("Price") & "��</a>"
                Else
                    strSalePromotion = "�����Ʒ�� <b>" & rsProduct("PresentNumber") & "</b> " & rsPresent1("Unit") & rsPresent1("ProductName")
                End If
            Else
                strSalePromotion = "������"
            End If
            Set rsPresent1 = Nothing
        End Select
        strHtml = Replace(strHtml, "{$SalePromotion}", strSalePromotion)
    End If
    If InStr(strHtml, "{$Hits}") > 0 Then strHtml = Replace(strHtml, "{$Hits}", GetHits())
    If InStr(strHtml, "{$ProductProperty}") > 0 Then strHtml = Replace(strHtml, "{$ProductProperty}", GetProductProperty())
    If InStr(strHtml, "{$Stars}") > 0 Then strHtml = Replace(strHtml, "{$Stars}", GetStars(rsProduct("Stars")))
    If InStr(strHtml, "{$ProductThumb}") > 0 Then strHtml = Replace(strHtml, "{$ProductThumb}", GetProductThumb(rsProduct("ProductThumb"), 130, 0, 0))
    If InStr(strHtml, "{$JsProductThumb}") > 0 Then strHtml = Replace(strHtml, "{$JsProductThumb}", GetJsProductThumb(rsProduct("ProductThumb"), 130, 0, 0))
    If InStr(strHtml, "{$UpdateDate}") > 0 Then strHtml = Replace(strHtml, "{$UpdateDate}", FormatDateTime(rsProduct("UpdateTime"), 2))
    strHtml = Replace(strHtml, "{$UpdateTime}", rsProduct("UpdateTime"))
    
    If InStr(strHtml, "{$ProductIntro}") > 0 Then strHtml = Replace(strHtml, "{$ProductIntro}", PE_HTMLEncode(ReplaceText(rsProduct("ProductIntro"), 1)))
    Dim strProductIntro
    regEx.Pattern = "\{\$ProductIntro\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strProductIntro= "����ʽ��ǩ��{$strProductIntro(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"

        Else
            Select Case PE_Clng(arrTemp(0))
            Case 1
                    strProductIntro = PE_HTMLEncode(ReplaceText(rsProduct("ProductIntro"), 1))
            Case 2
                If PE_Clng(arrTemp(1))>0 then
                    strProductIntro = GetSubStr(nohtml(rsProduct("ProductIntro")),PE_Clng(arrTemp(1)),False)
                Else
                    strProductIntro = nohtml(rsProduct("ProductIntro"))
                End IF
            End Select
        End If
        strHtml = Replace(strHtml, Match.Value, strProductIntro)
	Next
	
    If InStr(strHtml, "{$ProductExplain}") > 0 Then strHtml = Replace(strHtml, "{$ProductExplain}", ReplaceKeyLink(ReplaceText(Replace(Replace(rsProduct("ProductExplain"), "[InstallDir_ChannelDir]", ChannelUrl & "/"), "{$UploadDir}", UploadDir), 1)))
    
    strHtml = Replace(strHtml, "{$ProductProtect}", "")
    strHtml = Replace(strHtml, "{$ShowAD}", "")
    If InStr(strHtml, "{$Keyword}") > 0 Then strHtml = PE_Replace(strHtml, "{$Keyword}", GetKeywords(",", rsProduct("Keyword")))
    If InStr(strHtml, "{$Vote}") > 0 Then strHtml = Replace(strHtml, "{$Vote}", GetVoteOfContent(ProductID)) 'ͶƱ��ǩ
    Select Case rsProduct("ServiceTerm")
    Case -1
        strHtml = Replace(strHtml, "{$ServiceTerm}", "����")
    Case 0
        strHtml = Replace(strHtml, "{$ServiceTerm}", "����������")
    Case Else
        strHtml = Replace(strHtml, "{$ServiceTerm}", "<b>" & rsProduct("ServiceTerm") & "</b> ��")
    End Select
    
    If InStr(strHtml, "{$CorrelativeProduct}") > 0 Then strHtml = Replace(strHtml, "{$CorrelativeProduct}", GetCorrelative(5, 50, 2, 130, 90, 4))
    
    
    Dim strProductThumb
    '�滻��Ʒ����ͼ
    regEx.Pattern = "\{\$ProductThumb\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strProductThumb = GetProductThumb(Trim(rsProduct("ProductThumb")), PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), 0)
        strHtml = Replace(strHtml, Match.Value, strProductThumb)
    Next
    
  Dim strJsProductThumb
    regEx.Pattern = "\{\$JsProductThumb\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strJsProductThumb = GetJsProductThumb(Trim(rsProduct("ProductThumb")), PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), 0)
        strHtml = Replace(strHtml, Match.Value, strJsProductThumb)
    Next

    Dim strCorrelativeProduct
    regEx.Pattern = "\{\$CorrelativeProduct\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) = 2 Then
            strCorrelativeProduct = GetCorrelative(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), 130, 90, 4)
        ElseIf UBound(arrTemp) = 5 Then
            strCorrelativeProduct = GetCorrelative(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)))
        Else
            strCorrelativeProduct = "����ʽ��ǩ��{$CorrelativeProduct(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        End If
        strHtml = Replace(strHtml, Match.Value, strCorrelativeProduct)
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
    
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
End Sub

Public Sub GetHtml_Search()
    Select Case strField
    Case "Title"
        strField = "ProductName"
    Case "Content"
        strField = "ProductExplain"
    End Select

    strHtml = GetTemplate(ChannelID, 5, 0)
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
    If InStr(strHtml, "{$SearchResult}") > 0 Then strHtml = Replace(strHtml, "{$SearchResult}", GetSearchResult(130, 90, 2))
    strHtml = GetSearchResult2(strHtml)
    
    Dim strSearchResult
    Dim arrTemp
    regEx.Pattern = "\{\$SearchResult\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strSearchResult = GetSearchResult(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)))
        strHtml = Replace(strHtml, Match.Value, strSearchResult)
    Next

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

    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage_SearchResult, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage_SearchResult, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
End Sub

Public Sub ShowFavorite()
    Response.Write "<table width='100%' cellpadding='3' cellspacing='1' border='0' align='center' class='border'><tr valign='top' class='tdbg'>"
    
    Dim sqlFavorite, rsFavorite, iNum, strLink
    
    sqlFavorite = sqlFavorite & "select P.ProductID,P.ProductName,P.Discount,P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,P.BeginDate,P.EndDate,P.UpdateTime,P.ProductThumb"
    If UseCreateHTML > 0 Then
        sqlFavorite = sqlFavorite & ",C.ParentDir,C.ClassDir from PE_Product P left join PE_Class C on P.ClassID=C.ClassID"
    Else
        sqlFavorite = sqlFavorite & " from PE_Product P"
    End If
    sqlFavorite = sqlFavorite & " where P.Deleted=" & PE_False & " and P.EnableSale=" & PE_True & ""
    sqlFavorite = sqlFavorite & " and ProductID in (select InfoID from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & ")"
    sqlFavorite = sqlFavorite & " order by P.ProductID desc"

    Set rsFavorite = Server.CreateObject("ADODB.Recordset")
    rsFavorite.Open sqlFavorite, Conn, 1, 1
    If rsFavorite.BOF And rsFavorite.EOF Then
        totalPut = 0
        Response.Write "<td align='center' class='tdbg3'><img class='pic5' src='" & strInstallDir & "images/nopic.gif' width='130' height='90' border='0'><br>û���ղ��κ�" & ChannelShortName & "</td>"
    Else
        iNum = 0
        totalPut = rsFavorite.RecordCount
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
                    rsFavorite.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
            End If
        End If
        Do While Not rsFavorite.EOF
            If UseCreateHTML > 0 Then
                strLink = "<a href='" & GetProductUrl(rsFavorite("ParentDir"), rsFavorite("ClassDir"), rsFavorite("UpdateTime"), rsFavorite("ProductID")) & "'"
            Else
                strLink = "<a href='" & GetProductUrl("", "", "", rsFavorite("ProductID")) & "'"
            End If
            strLink = strLink & " target='_blank'>"

            Response.Write "<td class='tdbg3'>"
            Response.Write "<table width='100%' cellspacing='2' border='0'>"
            Response.Write "<tr><td align='center' rowspan='2'><input type='checkbox' name='InfoID' value='" & rsFavorite("ProductID") & "'>" & strLink & GetProductThumb(rsFavorite("ProductThumb"), 130, 90, 0) & "</a></td>"
            Response.Write "<td align='left'>" & strLink & rsFavorite("ProductName") & "</a><br>" & GetProductPrice(0, False, rsFavorite("ProductType"), rsFavorite("Price_Original"), rsFavorite("Price"), rsFavorite("Price_Market"), rsFavorite("Price_Member"), rsFavorite("BeginDate"), rsFavorite("EndDate"), rsFavorite("Discount")) & "</td></tr>"
            Response.Write "<tr><td align='left' valign='bottom'><a href='" & ChannelUrl_ASPFile & "/ShoppingCart.asp?Action=Add&ProductID=" & rsFavorite("ProductID") & "' target='ShoppingCart'><img src='" & ChannelUrl & "/images/ProductBuy.gif' border='0'></a>&nbsp;&nbsp;" & strLink & "<img src='" & ChannelUrl & "/images/ProductContent.gif' border='0'></a>&nbsp;&nbsp;<a href='User_Favorite.asp?Action=Remove&ChannelID=" & ChannelID & "&InfoID=" & rsFavorite("ProductID") & "'><img src='images/fav2.gif' border='0'></a></td></tr>"
            Response.Write "</table></td>"
            rsFavorite.MoveNext
            iNum = iNum + 1
            If iNum >= MaxPerPage Then Exit Do
            If ((iNum Mod 2 = 0) And (Not rsFavorite.EOF)) Then Response.Write "</tr><tr valign='top' class='tdbg'>"
        Loop
    End If
    Response.Write "</tr></table><br>"
    rsFavorite.Close
    Set rsFavorite = Nothing
    
    Response.Write ShowPage("User_Favorite.asp?ChannelID=" & ChannelID & "", totalPut, 20, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False)
End Sub
Function XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    XmlText_Class = XmlText("Product", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Product", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>

