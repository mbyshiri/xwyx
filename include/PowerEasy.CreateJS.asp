<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Write "<html><head><title>更新JS文件</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Call GetChannel(ChannelID)
Dim rsJs
Select Case Action
Case "CreateAllJs"
    Call CreateAllJS
Case "CreateJs"
    Call CreateJS
End Select
If Trim(Request("ShowBack")) = "Yes" Then
    Response.Write "<p align='center'><a href='" & ComeUrl & "'>【返回】</a></p>"
End If
Response.Write "</body></html>"
Call CloseConn

Sub CreateAllJS()
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Response.Write "<li>开始更新所有JS文件……</li>"
    Set rsJs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID)
    Do While Not rsJs.EOF
        Call CreateJS_CommonSub
        rsJs.MoveNext
    Loop
    rsJs.Close
    Set rsJs = Nothing
    Response.Write "<li>更新所有JS文件成功！</li>"
End Sub

Sub CreateJS()
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    Dim ID
    ID = PE_CLng(Trim(Request("ID")))
    If ID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    Set rsJs = Conn.Execute("select * from PE_JsFile where ID=" & ID)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If
    
    Call CreateJS_CommonSub

    rsJs.Close
    Set rsJs = Nothing
End Sub

Sub CreateJS_CommonSub()
    If rsJs("JsType") = 0 Then
        Select Case ModuleType
        Case 1
            Call CreateJS_ArticleList(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        Case 2
            Call CreateJS_SoftList(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        Case 3
            Call CreateJS_PhotoList(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        Case 5
            Call CreateJS_ProductList(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        End Select
    Else
        Select Case ModuleType
        Case 1
            Call CreateJS_PicArticle(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        Case 2
            Call CreateJS_PicSoft(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        Case 3
            Call CreateJS_PicPhoto(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        Case 5
            Call CreateJS_PicProduct(rsJs("JsFileName"), rsJs("Config"), rsJs("ContentType"))
        End Select
    End If
    Response.Write "<li>更新“" & rsJs("JsName") & "”JS文件成功！</li>"
End Sub


Sub CreateJS_ArticleList(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, ArticleNum, IsHot, IsElite, DateNum, OrderType, ShowType, TitleLen, ContentLen
    Dim ShowClassName, ShowIncludePic, ShowAuthor, ShowDateType, ShowHits, ShowHotSign
    Dim ShowNewSign, ShowTips, ShowCommentLink, OpenType, UrlType, ShowPropertyType
    Dim InputerName, Cols, CssNameA, CssName1, CssName2
    
    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    JsConfig = Split(arrConfig, "|")
    ShowType = PE_CLng(JsConfig(0))
    ArticleNum = PE_CLng(JsConfig(1))
    ClassID = PE_CLng(JsConfig(2))
    IncludeChild = CBool(JsConfig(3))
    IsHot = CBool(JsConfig(4))
    IsElite = CBool(JsConfig(5))
    DateNum = PE_CLng(JsConfig(6))
    OrderType = PE_CLng(JsConfig(7))
    TitleLen = PE_CLng(JsConfig(8))
    ContentLen = PE_CLng(JsConfig(9))
    ShowClassName = CBool(JsConfig(10))
    ShowIncludePic = CBool(JsConfig(11))
    ShowAuthor = CBool(JsConfig(12))
    ShowDateType = PE_CLng(JsConfig(13))
    ShowHits = CBool(JsConfig(14))
    ShowHotSign = CBool(JsConfig(15))
    ShowNewSign = CBool(JsConfig(16))
    ShowTips = CBool(JsConfig(17))
    ShowCommentLink = CBool(JsConfig(18))
    OpenType = PE_CLng(JsConfig(19))
    SpecialID = PE_CLng(JsConfig(20))
    UrlType = PE_CLng(JsConfig(21))
    ShowPropertyType = PE_CLng(JsConfig(22))
    InputerName = ZeroToEmpty(JsConfig(23))
    Cols = PE_CLng1(JsConfig(24))
    CssNameA = ZeroToEmpty(JsConfig(25))
    CssName1 = ZeroToEmpty(JsConfig(26))
    CssName2 = ZeroToEmpty(JsConfig(27))

    Dim PE_Article
    Set PE_Article = New Article
    Call PE_Article.Init
    strJS = PE_Article.GetArticleList(ChannelID, ClassID, IncludeChild, SpecialID, UrlType, ArticleNum, IsHot, IsElite, InputerName, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowIncludePic, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, ShowCommentLink, False, OpenType, Cols, CssNameA, CssName1, CssName2)
    Set PE_Article = Nothing
    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)

End Sub

Sub CreateJS_PicArticle(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, ArticleNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType
    Dim InputerName, CssNameA, CssName1, CssName2
    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"

    JsConfig = Split(arrConfig, "|")
    ClassID = PE_CLng(JsConfig(0))
    IncludeChild = CBool(JsConfig(1))
    ArticleNum = PE_CLng(JsConfig(2))
    IsHot = CBool(JsConfig(3))
    IsElite = CBool(JsConfig(4))
    DateNum = PE_CLng(JsConfig(5))
    OrderType = PE_CLng(JsConfig(6))
    ShowType = PE_CLng(JsConfig(7))
    ImgWidth = PE_CLng(JsConfig(8))
    ImgHeight = PE_CLng(JsConfig(9))
    TitleLen = PE_CLng(JsConfig(10))
    ContentLen = PE_CLng(JsConfig(11))
    ShowTips = CBool(JsConfig(12))
    Cols = PE_CLng1(JsConfig(13))
    SpecialID = PE_CLng(JsConfig(14))
    UrlType = PE_CLng(JsConfig(15))
    
    Dim PE_Article
    Set PE_Article = New Article
    Call PE_Article.Init
    strJS = PE_Article.GetPicArticle(ChannelID, ClassID, IncludeChild, SpecialID, ArticleNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType)
    Set PE_Article = Nothing

    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)
    
End Sub


Sub CreateJS_PhotoList(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, PhotoNum, IsHot, IsElite, DateNum, OrderType, ShowType, TitleLen, ContentLen
    Dim ShowClassName, ShowAuthor, ShowDateType, ShowHits, ShowHotSign
    Dim ShowNewSign, ShowTips, OpenType, UrlType, ShowPropertyType
    Dim InputerName, Cols, CssNameA, CssName1, CssName2
    
    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    JsConfig = Split(arrConfig, "|")
    ShowType = PE_CLng(JsConfig(0))
    PhotoNum = PE_CLng(JsConfig(1))
    ClassID = PE_CLng(JsConfig(2))
    IncludeChild = CBool(JsConfig(3))
    IsHot = CBool(JsConfig(4))
    IsElite = CBool(JsConfig(5))
    DateNum = PE_CLng(JsConfig(6))
    OrderType = PE_CLng(JsConfig(7))
    TitleLen = PE_CLng(JsConfig(8))
    ContentLen = PE_CLng(JsConfig(9))
    ShowClassName = CBool(JsConfig(10))
    ShowAuthor = CBool(JsConfig(11))
    ShowDateType = PE_CLng(JsConfig(12))
    ShowHits = CBool(JsConfig(13))
    ShowHotSign = CBool(JsConfig(14))
    ShowNewSign = CBool(JsConfig(15))
    ShowTips = CBool(JsConfig(16))
    OpenType = PE_CLng(JsConfig(17))
    UrlType = PE_CLng(JsConfig(18))
    ShowPropertyType = PE_CLng(JsConfig(19))
    InputerName = ZeroToEmpty(JsConfig(20))
    SpecialID = PE_CLng(JsConfig(21))
    Cols = PE_CLng1(JsConfig(22))
    CssNameA = ZeroToEmpty(JsConfig(23))
    CssName1 = ZeroToEmpty(JsConfig(24))
    CssName2 = ZeroToEmpty(JsConfig(25))
    
    Dim PE_Photo
    Set PE_Photo = New Photo
    Call PE_Photo.Init
    strJS = PE_Photo.GetPhotoList(ChannelID, ClassID, IncludeChild, SpecialID, UrlType, PhotoNum, IsHot, IsElite, InputerName, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, False, OpenType, Cols, CssNameA, CssName1, CssName2)
    Set PE_Photo = Nothing

    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)

End Sub

Sub CreateJS_PicPhoto(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, PhotoNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType

    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"

    JsConfig = Split(arrConfig, "|")
    ClassID = PE_CLng(JsConfig(0))
    IncludeChild = CBool(JsConfig(1))
    PhotoNum = PE_CLng(JsConfig(2))
    IsHot = CBool(JsConfig(3))
    IsElite = CBool(JsConfig(4))
    DateNum = PE_CLng(JsConfig(5))
    OrderType = PE_CLng(JsConfig(6))
    ShowType = PE_CLng(JsConfig(7))
    ImgWidth = PE_CLng(JsConfig(8))
    ImgHeight = PE_CLng(JsConfig(9))
    TitleLen = PE_CLng(JsConfig(10))
    ContentLen = PE_CLng(JsConfig(11))
    ShowTips = CBool(JsConfig(12))
    Cols = PE_CLng1(JsConfig(13))
    UrlType = PE_CLng(JsConfig(14))
    SpecialID = PE_CLng(JsConfig(15))
    
    Dim PE_Photo
    Set PE_Photo = New Photo
    Call PE_Photo.Init
    strJS = PE_Photo.GetPicPhoto(ChannelID, ClassID, IncludeChild, SpecialID, PhotoNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType)
    Set PE_Photo = Nothing

    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)
    
End Sub

Sub CreateJS_SoftList(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, SoftNum, IsHot, IsElite, DateNum, OrderType, ShowType, TitleLen, ContentLen
    Dim ShowClassName, ShowAuthor, ShowDateType, ShowHits, ShowHotSign
    Dim ShowNewSign, ShowTips, OpenType, UrlType, ShowPropertyType
    Dim InputerName, Cols, CssNameA, CssName1, CssName2
    
    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    JsConfig = Split(arrConfig, "|")
    ShowType = PE_CLng(JsConfig(0))
    SoftNum = PE_CLng(JsConfig(1))
    ClassID = PE_CLng(JsConfig(2))
    IncludeChild = CBool(JsConfig(3))
    IsHot = CBool(JsConfig(4))
    IsElite = CBool(JsConfig(5))
    DateNum = PE_CLng(JsConfig(6))
    OrderType = PE_CLng(JsConfig(7))
    TitleLen = PE_CLng(JsConfig(8))
    ContentLen = PE_CLng(JsConfig(9))
    ShowClassName = CBool(JsConfig(10))
    ShowAuthor = CBool(JsConfig(11))
    ShowDateType = PE_CLng(JsConfig(12))
    ShowHits = CBool(JsConfig(13))
    ShowHotSign = CBool(JsConfig(14))
    ShowNewSign = CBool(JsConfig(15))
    ShowTips = CBool(JsConfig(16))
    OpenType = PE_CLng(JsConfig(17))
    UrlType = PE_CLng(JsConfig(18))
    ShowPropertyType = PE_CLng(JsConfig(19))
    SpecialID = PE_CLng(JsConfig(20))
    InputerName = ZeroToEmpty(JsConfig(21))
    Cols = PE_CLng1(JsConfig(22))
    CssNameA = ZeroToEmpty(JsConfig(23))
    CssName1 = ZeroToEmpty(JsConfig(24))
    CssName2 = ZeroToEmpty(JsConfig(25))
    Dim PE_Soft
    Set PE_Soft = New Soft
    Call PE_Soft.Init
    strJS = PE_Soft.GetSoftList(ChannelID, ClassID, IncludeChild, SpecialID, UrlType, SoftNum, IsHot, IsElite, InputerName, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, False, OpenType, Cols, CssNameA, CssName1, CssName2)
    Set PE_Soft = Nothing

    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)
    
End Sub

Sub CreateJS_PicSoft(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, SoftNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType

    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"

    JsConfig = Split(arrConfig, "|")
    ClassID = PE_CLng(JsConfig(0))
    IncludeChild = CBool(JsConfig(1))
    SoftNum = PE_CLng(JsConfig(2))
    IsHot = CBool(JsConfig(3))
    IsElite = CBool(JsConfig(4))
    DateNum = PE_CLng(JsConfig(5))
    OrderType = PE_CLng(JsConfig(6))
    ShowType = PE_CLng(JsConfig(7))
    ImgWidth = PE_CLng(JsConfig(8))
    ImgHeight = PE_CLng(JsConfig(9))
    TitleLen = PE_CLng(JsConfig(10))
    ContentLen = PE_CLng(JsConfig(11))
    ShowTips = CBool(JsConfig(12))
    Cols = PE_CLng1(JsConfig(13))
    UrlType = PE_CLng(JsConfig(14))
    SpecialID = PE_CLng(JsConfig(15))
    
    Dim PE_Soft
    Set PE_Soft = New Soft
    Call PE_Soft.Init
    strJS = PE_Soft.GetPicSoft(ChannelID, ClassID, IncludeChild, SpecialID, SoftNum, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, ContentLen, ShowTips, Cols, UrlType)
    Set PE_Soft = Nothing

    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)
    
End Sub

Sub CreateJS_ProductList(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, ProductNum, ProductType, IsHot, IsElite, DateNum, OrderType, ShowType
    Dim TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowDateType, ShowHotSign, ShowNewSign, UsePage, OpenType, UrlType
    
    Dim IntervalLines, Cols, ShowTableTitle, TableTitleStr, ShowProductModel, ShowProductStandard
    Dim ShowUnit, ShowStocksType, ShowWeight, ShowPrice_Market, ShowPrice_Original, ShowPrice
    Dim ShowPrice_Member, ShowDiscount, ShowButtonType, ButtonStyle
    Dim CssNameTable, CssNameTitle, CssNameA, CssName1, CssName2

    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    JsConfig = Split(arrConfig, "|")
    ClassID = PE_CLng(JsConfig(0))
    IncludeChild = CBool(JsConfig(1))
    SpecialID = PE_CLng(JsConfig(2))
    ProductNum = PE_CLng(JsConfig(3))
    ProductType = PE_CLng(JsConfig(4))
    IsHot = CBool(JsConfig(5))
    IsElite = CBool(JsConfig(6))
    DateNum = PE_CLng(JsConfig(7))
    OrderType = PE_CLng(JsConfig(8))
    ShowType = PE_CLng(JsConfig(9))
    TitleLen = PE_CLng(JsConfig(10))
    ContentLen = PE_CLng(JsConfig(11))
    ShowClassName = CBool(JsConfig(12))
    ShowPropertyType = PE_CLng(JsConfig(13))
    ShowDateType = PE_CLng(JsConfig(14))
    ShowHotSign = CBool(JsConfig(15))
    ShowNewSign = CBool(JsConfig(16))
    UsePage = CBool(JsConfig(17))
    OpenType = PE_CLng(JsConfig(18))
    UrlType = PE_CLng(JsConfig(19))

    IntervalLines = PE_CLng(JsConfig(20))
    Cols = PE_CLng1(JsConfig(21))
    ShowTableTitle = CBool(JsConfig(22))
    TableTitleStr = Trim(Replace(JsConfig(23), "{$}", "|"))
    ShowProductModel = CBool(JsConfig(24))
    ShowProductStandard = CBool(JsConfig(25))
    ShowUnit = CBool(JsConfig(26))
    ShowStocksType = CBool(JsConfig(27))
    ShowWeight = CBool(JsConfig(28))
    ShowPrice_Market = CBool(JsConfig(29))
    ShowPrice_Original = CBool(JsConfig(30))
    ShowPrice = CBool(JsConfig(31))
    ShowPrice_Member = CBool(JsConfig(32))
    ShowDiscount = CBool(JsConfig(33))

    ShowButtonType = PE_CLng(JsConfig(34))
    ButtonStyle = PE_CLng(JsConfig(35))
    CssNameTable = ZeroToEmpty(JsConfig(36))
    CssNameTitle = ZeroToEmpty(JsConfig(37))
    CssNameA = ZeroToEmpty(JsConfig(38))
    CssName1 = ZeroToEmpty(JsConfig(39))
    CssName2 = ZeroToEmpty(JsConfig(40))
    
    Dim PE_Product

    Set PE_Product = New Product
    Call PE_Product.Init

    strJS = PE_Product.GetProductList(ClassID, IncludeChild, SpecialID, ProductNum, ProductType, IsHot, IsElite, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowDateType, ShowHotSign, ShowNewSign, UsePage, OpenType, UrlType, IntervalLines, Cols, ShowTableTitle, TableTitleStr, ShowProductModel, ShowProductStandard, ShowUnit, ShowStocksType, ShowWeight, ShowPrice_Market, ShowPrice_Original, ShowPrice, ShowPrice_Member, ShowDiscount, ShowButtonType, ButtonStyle, CssNameTable, CssNameTitle, CssNameA, CssName1, CssName2)
    Set PE_Product = Nothing

    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)
    
End Sub

Sub CreateJS_PicProduct(JsFileName, ByVal arrConfig, ByVal ContentType)
    Dim JsConfig, hf, strJS
    Dim ClassID, IncludeChild, SpecialID, ProductNum, ProductType, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, Cols, UrlType
    
    arrConfig = arrConfig & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    JsConfig = Split(arrConfig, "|")
    ClassID = PE_CLng(JsConfig(0))
    IncludeChild = CBool(JsConfig(1))
    SpecialID = PE_CLng(JsConfig(2))
    ProductNum = PE_CLng(JsConfig(3))
    ProductType = PE_CLng(JsConfig(4))
    IsHot = CBool(JsConfig(5))
    IsElite = CBool(JsConfig(6))
    DateNum = PE_CLng(JsConfig(7))
    OrderType = PE_CLng(JsConfig(8))
    ShowType = PE_CLng(JsConfig(9))
    ImgWidth = PE_CLng(JsConfig(10))
    ImgHeight = PE_CLng(JsConfig(11))
    TitleLen = PE_CLng(JsConfig(12))
    Cols = PE_CLng1(JsConfig(13))
    UrlType = PE_CLng(JsConfig(14))

    Dim PE_Product
    Set PE_Product = New Product
    Call PE_Product.Init
    strJS = PE_Product.GetPicProduct(ClassID, IncludeChild, SpecialID, ProductNum, ProductType, IsHot, IsElite, DateNum, OrderType, ShowType, ImgWidth, ImgHeight, TitleLen, Cols, UrlType, False, False, False, 0, 1)
    Set PE_Product = Nothing
    
    Call SaveJsFile(ContentType, InstallDir & ChannelDir, JsFileName, strJS, ComeUrl)
    
End Sub


'=================================================
'过程名：SaveJsFile
'作  用：保存JS文件
'=================================================
Sub SaveJsFile(ByVal ContentType, ByVal SaveFilePath, ByVal JsFileName, ByVal strJS, ByVal ComeUrl)
    'On Error Resume Next
    If PE_CLng(ContentType) = 0 Then
        strJS = Replace(Replace(strJS, vbCrLf, "\n"), """", "\""")
        strJS = "document.write(" & Chr(34) & strJS & Chr(34) & ");"
    End If
    If CreateMultiFolder(SaveFilePath & "/js") Then
        Call WriteToFile(SaveFilePath & "/js/" & JsFileName, strJS)
    End If
End Sub

%>
