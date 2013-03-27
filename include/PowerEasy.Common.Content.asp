<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'**************************************************
'��������GetListPath
'��  �ã�����б�·��
'��  ����iStructureType ---- Ŀ¼�ṹ��ʽ
'        iListFileType ---- �б��ļ���ʽ
'        sParentDir ---- ����ĿĿ¼
'        sClassDir ---- ��ǰ��ĿĿ¼
'����ֵ���б�·��
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
'��������GetListFileName
'��  �ã�����б��ļ�����(��Ŀ�з�ҳʱ)
'��  ����iListFileType ----�б��ļ���ʽ
'        iClassID ---- ��ĿID
'        iCurrentPage ---- �б�ǰҳ��
'����ֵ���б��ļ�����
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
'��������GetList_1FileName
'��  �ã�����б��ļ���
'��  ����iListFileType ---- �б��ļ���ʽ
'        iClassID ---- ��ĿID
'����ֵ���б��ļ���
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
'��������GetItemPath
'��  �ã������Ŀ·��
'��  ����iStructureType ---- Ŀ¼�ṹ��ʽ
'        sParentDir ---- ����ĿĿ¼
'        sClassDir ---- ��ǰ��ĿĿ¼
'        UpdateTime ---- ��ĿĿ¼
'����ֵ�������Ŀ·��
'**************************************************
Public Function GetItemPath(iStructureType, sParentDir, sClassDir, UpdateTime)
    Select Case iStructureType
    Case 0      'Ƶ��/����/С��/�·�/�ļ�����Ŀ�ּ����ٰ��·ݱ��棩
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 1      'Ƶ��/����/С��/����/�ļ�����Ŀ�ּ����ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 2      'Ƶ��/����/С��/�ļ�����Ŀ�ּ������ٰ��·ݣ�
        GetItemPath = sParentDir & sClassDir & "/"
    Case 3      'Ƶ��/��Ŀ/�·�/�ļ�����Ŀƽ�����ٰ��·ݱ��棩
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 4      'Ƶ��/��Ŀ/����/�ļ�����Ŀƽ�����ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 5      'Ƶ��/��Ŀ/�ļ�����Ŀƽ�������ٰ��·ݣ�
        GetItemPath = "/" & sClassDir & "/"
    Case 6      'Ƶ��/�ļ���ֱ�ӷ���Ƶ��Ŀ¼�У�
        GetItemPath = "/"
    Case 7      'Ƶ��/HTML/�ļ���ֱ�ӷ���ָ���ġ�HTML���ļ����У�
        GetItemPath = "/HTML/"
    Case 8      'Ƶ��/���/�ļ���ֱ�Ӱ���ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/"
    Case 9      'Ƶ��/�·�/�ļ���ֱ�Ӱ��·ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 10     'Ƶ��/����/�ļ���ֱ�Ӱ����ڱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 11     'Ƶ��/���/�·�/�ļ����Ȱ���ݣ��ٰ��·ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 12     'Ƶ��/���/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 13     'Ƶ��/�·�/����/�ļ����Ȱ��·ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 14     'Ƶ��/���/�·�/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    End Select
    GetItemPath = Replace(GetItemPath, "//", "/")
End Function

'**************************************************
'��������GetItemFileName
'��  �ã������Ŀ����
'��  ����iFileNameType ---- �ļ���������
'        sChannelDir ---- ��ǰƵ��Ŀ¼
'        UpdateTime ---- ����ʱ��
'        ItemID ---- ����ID��ArticleID/SoftID/PhotoID)
'����ֵ�������Ŀ����
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
'��������ReplaceKeyLink
'��  �ã��滻վ������
'��  ����iText-----�����ַ���
'����ֵ���滻���ַ���
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
'��������PE_Replace_keylink
'��  �ã�ʹ�������滻��HTML�����еķ�HTML��ǩ���ݽ����滻
'��  ����expression ---- ������
'        find ---- ���滻���ַ�
'        replacewith ---- �滻����ַ�
'        replacenum  ---- �滻����
'����ֵ���ݴ����滻�ַ���,��� replacewith ���ַ�,���滻���ַ� �滻�ɿ�
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
        regEx.Pattern = "([][$( )*+.?\\^{|])"  '������ʽ�е������ַ���Ҫ����ת��
        find = regEx.Replace(find, "\$1")
        replacewith = Replace(replacewith, "$", "&#36;") '��$���д����������
        regEx.Pattern = "(>[^><]*)" & find & "([^><]*<)(?!/a)"
        newStr = regEx.Replace(">" & expression & "<", "$1" & replacewith & "$2")
        PE_Replace_keylink = Mid(newStr, 2, Len(newStr) - 2)
    End If
End Function

'**************************************************
'��������GetClassFild
'��  �ã��õ���Ŀ����
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
            ClassField(1) = "�������κ���Ŀ"
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
    If IsNull(tmpAuthorName) Or tmpAuthorName = "δ֪" Or tmpAuthorName = "����" Then
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
    If IsNull(tmpCopyFrom) Or tmpCopyFrom = "��վԭ��" Then
        GetCopyFromInfo = "��վԭ��"
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
