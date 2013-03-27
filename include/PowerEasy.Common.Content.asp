<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'**************************************************
'函数名：GetListPath
'作  用：获得列表路径
'参  数：iStructureType ---- 目录结构方式
'        iListFileType ---- 列表文件方式
'        sParentDir ---- 父栏目目录
'        sClassDir ---- 当前栏目目录
'返回值：列表路径
'**************************************************
Public Function GetListPath(iStructureType, iListFileType, sParentDir, sClassDir)
    Select Case iListFileType
    Case 0
        Select Case iStructureType
            Case 0, 1, 2
                GetListPath = sParentDir & sClassDir & "/"
            Case 3, 4, 5
                GetListPath = "/" & sClassDir & "/"
            Case Else
                GetListPath = sParentDir & sClassDir & "/"
        End Select
    Case 1
        GetListPath = "/List/"
    Case 2
        GetListPath = "/"
    End Select
    GetListPath = Replace(GetListPath, "//", "/")
End Function

'**************************************************
'函数名：GetListFileName
'作  用：获得列表文件名称(栏目有分页时)
'参  数：iListFileType ----列表文件方式
'        iClassID ---- 栏目ID
'        iCurrentPage ---- 列表当前页数
'返回值：列表文件名称
'**************************************************
Public Function GetListFileName(iListFileType, iClassID, iCurrentPage, iPages)
    Select Case iListFileType
    Case 0
        If iCurrentPage = 1 Then
            GetListFileName = "Index"
        Else
            GetListFileName = "List_" & iPages - iCurrentPage + 1
        End If
    Case 1, 2
        If iCurrentPage = 1 Then
            GetListFileName = "List_" & iClassID
        Else
            GetListFileName = "List_" & iClassID & "_" & iPages - iCurrentPage + 1
        End If
    End Select
End Function

'**************************************************
'函数名：GetList_1FileName
'作  用：获得列表文件名
'参  数：iListFileType ---- 列表文件方式
'        iClassID ---- 栏目ID
'返回值：列表文件名
'**************************************************
Public Function GetList_1FileName(iListFileType, iClassID)
    Select Case iListFileType
    Case 0
        GetList_1FileName = "List_0"
    Case 1
        GetList_1FileName = "List_" & iClassID & "_0"
    Case 2
        GetList_1FileName = "List_" & iClassID & "_0"
    End Select
End Function

'**************************************************
'函数名：GetItemPath
'作  用：获得项目路径
'参  数：iStructureType ---- 目录结构方式
'        sParentDir ---- 父栏目目录
'        sClassDir ---- 当前栏目目录
'        UpdateTime ---- 栏目目录
'返回值：获得项目路径
'**************************************************
Public Function GetItemPath(iStructureType, sParentDir, sClassDir, UpdateTime)
    Select Case iStructureType
    Case 0      '频道/大类/小类/月份/文件（栏目分级，再按月份保存）
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 1      '频道/大类/小类/日期/文件（栏目分级，再按日期分，每天一个目录）
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 2      '频道/大类/小类/文件（栏目分级，不再按月份）
        GetItemPath = sParentDir & sClassDir & "/"
    Case 3      '频道/栏目/月份/文件（栏目平级，再按月份保存）
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 4      '频道/栏目/日期/文件（栏目平级，再按日期分，每天一个目录）
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 5      '频道/栏目/文件（栏目平级，不再按月份）
        GetItemPath = "/" & sClassDir & "/"
    Case 6      '频道/文件（直接放在频道目录中）
        GetItemPath = "/"
    Case 7      '频道/HTML/文件（直接放在指定的“HTML”文件夹中）
        GetItemPath = "/HTML/"
    Case 8      '频道/年份/文件（直接按年份保存，每年一个目录）
        GetItemPath = "/" & Year(UpdateTime) & "/"
    Case 9      '频道/月份/文件（直接按月份保存，每月一个目录）
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 10     '频道/日期/文件（直接按日期保存，每天一个目录）
        GetItemPath = "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 11     '频道/年份/月份/文件（先按年份，再按月份保存，每月一个目录）
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 12     '频道/年份/日期/文件（先按年份，再按日期分，每天一个目录）
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 13     '频道/月份/日期/文件（先按月份，再按日期分，每天一个目录）
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 14     '频道/年份/月份/日期/文件（先按年份，再按日期分，每天一个目录）
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    End Select
    GetItemPath = Replace(GetItemPath, "//", "/")
End Function

'**************************************************
'函数名：GetItemFileName
'作  用：获得项目名称
'参  数：iFileNameType ---- 文件名称类型
'        sChannelDir ---- 当前频道目录
'        UpdateTime ---- 更新时间
'        ItemID ---- 内容ID（ArticleID/SoftID/PhotoID)
'返回值：获得项目名称
'**************************************************
Public Function GetItemFileName(iFileNameType, sChannelDir, UpdateTime, ItemID)
    Select Case iFileNameType
    Case 0
        GetItemFileName = ItemID
    Case 1
        GetItemFileName = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2)
    Case 2
        GetItemFileName = sChannelDir & "_" & ItemID
    Case 3
        GetItemFileName = sChannelDir & "_" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2)
    Case 4
        GetItemFileName = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2) & "_" & ItemID
    Case 5
        GetItemFileName = sChannelDir & "_" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2) & "_" & ItemID
    End Select
End Function



Sub CreateAllJS()
    Dim NeedCreateJS
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        If Status = 3 Then
            NeedCreateJS = True
        Else
            NeedCreateJS = False
        End If
    Else
        NeedCreateJS = True
    End If
    If NeedCreateJS = True then
        Response.Write "<br><iframe id='CreateJS' width='100%' height='100' frameborder='0' src='Admin_CreateJS.asp?ChannelID=" & ChannelID & "&Action=CreateAllJs'></iframe>"
    End If
End Sub

Sub CreateAllJS_User()
    If Status = 3 Then
        Response.Write "<br><iframe id='CreateJS' width='100%' height='100' frameborder='0' src='User_CreateJS.asp?ChannelID=" & ChannelID & "&Action=CreateAllJs'></iframe>"
    End If
End Sub

Function GetArticleUrl(ByVal tParentDir, ByVal tClassDir, ByVal tUpdateTime, ByVal tArticleID, ByVal tClassPurview, ByVal tInfoPurview, ByVal tInfoPoint)
    If IsNull(tParentDir) Then tParentDir = ""
    If IsNull(tClassDir) Then tClassDir = ""
    If IsNull(tClassPurview) Then tClassPurview = 0
    If IsNull(tInfoPurview) Then tInfoPurview = 0    
    If UseCreateHTML > 0 And tClassPurview = 0 And tInfoPoint = 0 And tInfoPurview = 0 Then
        GetArticleUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tArticleID) & FileExt_Item
    Else
        GetArticleUrl = ChannelUrl_ASPFile & "/ShowArticle.asp?ArticleID=" & tArticleID
    End If
End Function

Function GetPhotoUrl(ByVal tParentDir, ByVal tClassDir, ByVal tUpdateTime, ByVal tPhotoID, ByVal tClassPurview, ByVal tInfoPurview, ByVal tInfoPoint)
    If IsNull(tParentDir) Then tParentDir = ""
    If IsNull(tClassDir) Then tClassDir = ""
    If IsNull(tClassPurview) Then tClassPurview = 0
    If IsNull(tInfoPurview) Then tInfoPurview = 0
    
    If UseCreateHTML > 0 And tClassPurview = 0 And tInfoPoint = 0 And tInfoPurview = 0 Then
        GetPhotoUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tPhotoID) & FileExt_Item
    Else
        GetPhotoUrl = ChannelUrl_ASPFile & "/ShowPhoto.asp?PhotoID=" & tPhotoID
    End If
End Function

Function GetSoftUrl(ByVal tParentDir, ByVal tClassDir, ByVal tUpdateTime, ByVal tSoftID)
    If IsNull(tParentDir) Then tParentDir = ""
    If IsNull(tClassDir) Then tClassDir = ""
    
    If UseCreateHTML > 0 Then
        GetSoftUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tSoftID) & FileExt_Item
    Else
        GetSoftUrl = ChannelUrl_ASPFile & "/ShowSoft.asp?SoftID=" & tSoftID
    End If
End Function

Function GetProductUrl(ByVal tParentDir, ByVal tClassDir, ByVal tUpdateTime, ByVal tProductID)
    If IsNull(tParentDir) Then tParentDir = ""
    If IsNull(tClassDir) Then tClassDir = ""
    
    If UseCreateHTML > 0 Then
        GetProductUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tProductID) & FileExt_Item
    Else
        GetProductUrl = ChannelUrl_ASPFile & "/ShowProduct.asp?ProductID=" & tProductID
    End If
End Function

'**************************************************
'函数名：ReplaceKeyLink
'作  用：替换站内链接
'参  数：iText-----输入字符串
'返回值：替换后字符串
'**************************************************
Function ReplaceKeyLink(iText)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol, LinkUrl
    If PE_Cache.GetValue("Site_KeyList") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=0 order by Priority"
        rsKey.Open sqlKey, Conn, 1, 1
        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_KeyList", rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceKeyLink = iText
            Exit Function
        End If
    End If
    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_KeyList"), "@@@")
    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        If UBound(Keycol) >= 3 Then
            If Keycol(2) = 0 Then
                LinkUrl = "<a class=""channel_keylink"" href=""" & Keycol(1) & """>" & Keycol(0) & "</a>"
            Else
                LinkUrl = "<a class=""channel_keylink"" href=""" & Keycol(1) & """ target=""_blank"">" & Keycol(0) & "</a>"
            End If
            rText = PE_Replace_keylink(rText, Keycol(0), LinkUrl, Keycol(3))
        End If
    Next
    ReplaceKeyLink = rText
End Function


'**************************************************
'函数名：PE_Replace_keylink
'作  用：使用正则替换将HTML代码中的非HTML标签内容进行替换
'参  数：expression ---- 主数据
'        find ---- 被替换的字符
'        replacewith ---- 替换后的字符
'        replacenum  ---- 替换次数
'返回值：容错后的替换字符串,如果 replacewith 空字符,被替换的字符 替换成空
'**************************************************
Function PE_Replace_keylink(ByVal expression, ByVal find, ByVal replacewith, ByVal replacenum)
    If IsNull(expression) Or IsNull(find) Or IsNull(replacewith) Then
        PE_Replace_keylink = ""
        Exit Function
    End If

    Dim newStr
    If PE_Clng(replacenum) > 0 Then
        PE_Replace_keylink = Replace(expression, find, replacewith, 1, replacenum)
    Else
        regEx.Pattern = "([][$( )*+.?\\^{|])"  '正则表达式中的特殊字符，要进行转义
        find = regEx.Replace(find, "\$1")
        replacewith = Replace(replacewith, "$", "&#36;") '对$进行处理，特殊情况
        regEx.Pattern = "(>[^><]*)" & find & "([^><]*<)(?!/a)"
        newStr = regEx.Replace(">" & expression & "<", "$1" & replacewith & "$2")
        PE_Replace_keylink = Mid(newStr, 2, Len(newStr) - 2)
    End If
End Function

'**************************************************
'函数名：GetClassFild
'作  用：得到栏目属性
'**************************************************
Function GetClassFild(iClassID, iType)
    Dim rsClass
    If IsNull(iClassID) Then
        GetClassFild = 0
        Exit Function
    End If
    
    If iClassID <> PriClassID Or ClassField(1) = "" Then
        Set rsClass = Conn.Execute("select top 1 ClassID,ClassName,ClassPurview,ClassDir,ParentDir from PE_Class where ClassID=" & iClassID)
        If Not (rsClass.BOF Or rsClass.EOF) Then
            ClassField(0) = iClassID
            ClassField(1) = rsClass("ClassName")
            ClassField(2) = rsClass("ClassPurview")
            ClassField(3) = rsClass("ClassDir")
            ClassField(4) = rsClass("ParentDir")
            PriClassID = iClassID
        Else
            ClassField(0) = 0
            ClassField(1) = "不属于任何栏目"
            ClassField(2) = 0
            ClassField(3) = ""
            ClassField(4) = ""
        End If
        Set rsClass = Nothing
        
    End If
    GetClassFild = ClassField(iType)
End Function

Private Function GetAuthorInfo(tmpAuthorName, iChannelID)
    Dim i, tempauthor, authorarry, temprs, temparr
    If IsNull(tmpAuthorName) Or tmpAuthorName = "未知" Or tmpAuthorName = "佚名" Then
        GetAuthorInfo = tmpAuthorName
    Else
        authorarry = Split(tmpAuthorName, "|")
        For i = 0 To UBound(authorarry)
            tempauthor = tempauthor & "<a href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & iChannelID & "&AuthorName=" & authorarry(i) & "' title='" & authorarry(i) & "'>" & GetSubStr(authorarry(i), AuthorInfoLen, True) & "</a>"
            If i <> UBound(authorarry) Then tempauthor = tempauthor & "|"
        Next
        GetAuthorInfo = tempauthor
    End If
End Function

Private Function GetCopyFromInfo(tmpCopyFrom, iChannelID)
    Dim temprs, temparr
    If IsNull(tmpCopyFrom) Or tmpCopyFrom = "本站原创" Then
        GetCopyFromInfo = "本站原创"
    Else
        GetCopyFromInfo = "<a href='" & strInstallDir & "ShowCopyFrom.asp?ChannelID=" & iChannelID & "&SourceName=" & tmpCopyFrom & "'>" & tmpCopyFrom & "</a>"
    End If
End Function

Private Function GetInfoPoint(InfoPoint)
    If InfoPoint = 9999 Then
        GetInfoPoint = "0"
    Else
        GetInfoPoint = InfoPoint
    End If
End Function

Private Function GetKeywords(strSplit, strKeyword)
    GetKeywords = PE_Replace(Mid(strKeyword, 2, Len(strKeyword) - 2), "|", strSplit)
End Function


%>
