<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'定义频道设置相关的变量
Dim ChannelID, ChannelName, ChannelShortName, ChannelItemUnit, ChannelDir, ChannelPicUrl, ChannelUrl, ChannelUrl_ASPFile, UploadDir
Dim Meta_Keywords_Channel, Meta_Description_Channel, Custom_Content_Channel, ChannelPurview, ChannelArrGroupID, AuthorInfoLen
Dim ShowChannelName, ShowNameOnPath, ShowClassTreeGuide, ShowSuspensionPoints, DaysOfNew, HitsOfHot, Template_Index, DefaultSkinID
Dim MaxPerPage_Index, MaxPerPage_New, MaxPerPage_Hot, MaxPerPage_Elite, MaxPerPage_SpecialList
Dim UseCreateHTML, AutoCreateType, StructureType, ListFileType, FileNameType, FileExt_Index, FileExt_List, FileExt_Item, UpdatePages
Dim ItemCount_Channel, ItemChecked_Channel, CommentCount_Channel, SpecialCount_Channel
Dim ModuleType, ModuleName, SheetName
Dim arrEnabledTabs, MoneyPerKw
Dim TopMenuType, ClassGuideType, CheckLevel
Dim EmailOfReject, EmailOfPassed
Dim JS_SpecialNum
Dim MaxFileSize, Fields_Options
Dim FileName_Index
Dim MaxPerLine
Dim CommandChannelPoint

Sub GetChannel(tChannelID)
    Dim sqlChannel, rsChannel
    ModuleType = 0
    ChannelItemUnit = ""
    ChannelDir = ""
    ChannelPicUrl = ""
    Meta_Keywords_Channel = ""
    Meta_Description_Channel = ""
    Custom_Content_Channel = ""
    ChannelPurview = 0
    ChannelArrGroupID = ""
    AuthorInfoLen = 8
    MaxPerPage_Index = 20
    MaxPerPage_New = 20
    MaxPerPage_Hot = 20
    MaxPerPage_Elite = 20
    MaxPerPage_SpecialList = 20
    ShowClassTreeGuide = False
    HitsOfHot = SiteHitsOfHot
    DaysOfNew = 10
    DefaultSkinID = 0
    ShowSuspensionPoints = False
    UseCreateHTML = 0
    AutoCreateType = 0
    StructureType = 0
    ListFileType = 0
    FileNameType = 0
    FileExt_Index = arrFileExt(0)
    FileExt_List = arrFileExt(0)
    FileExt_Item = arrFileExt(0)
    UpdatePages = 3
    Template_Index = 0
    ItemCount_Channel = 0
    ItemChecked_Channel = 0
    CommentCount_Channel = 0
    SpecialCount_Channel = 0
    MaxPerLine = 10
    ChannelUrl = ""
    If tChannelID > 0 Then
        sqlChannel = "select * from PE_Channel where ChannelID=" & tChannelID
        Set rsChannel = Conn.Execute(sqlChannel)
        If rsChannel.BOF And rsChannel.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "找不到指定的频道"
        Else
            If rsChannel("Disabled") = True Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>此频道已经被管理员禁用！</li>"
            End If
            ChannelName = rsChannel("ChannelName")
            ChannelDir = rsChannel("ChannelDir")
            ChannelPicUrl = rsChannel("ChannelPicUrl")
            ChannelShortName = rsChannel("ChannelShortName")
            ChannelItemUnit = rsChannel("ChannelItemUnit")
            ShowChannelName = rsChannel("ShowName")
            ShowNameOnPath = rsChannel("ShowNameOnPath")
            UploadDir = rsChannel("UploadDir")
            Meta_Keywords_Channel = rsChannel("Meta_Keywords")
            Meta_Description_Channel = rsChannel("Meta_Description")
            Custom_Content_Channel = rsChannel("Custom_Content")
            ChannelPurview = rsChannel("ChannelPurview")
            ChannelArrGroupID = rsChannel("arrGroupID")
            AuthorInfoLen = rsChannel("AuthorInfoLen")
            MaxPerPage_Index = rsChannel("MaxPerPage_Index")
            MaxPerPage_SearchResult = rsChannel("MaxPerPage_SearchResult")
            MaxPerPage_New = rsChannel("MaxPerPage_New")
            MaxPerPage_Hot = rsChannel("MaxPerPage_Hot")
            MaxPerPage_Elite = rsChannel("MaxPerPage_Elite")
            MaxPerPage_SpecialList = rsChannel("MaxPerPage_SpecialList")
            ShowClassTreeGuide = rsChannel("ShowClassTreeGuide")
            HitsOfHot = rsChannel("HitsOfHot")
            DaysOfNew = rsChannel("DaysOfNew")
            DefaultSkinID = rsChannel("DefaultSkinID")
            CheckLevel = rsChannel("CheckLevel")
            ShowSuspensionPoints = rsChannel("ShowSuspensionPoints")
            UseCreateHTML = rsChannel("UseCreateHTML")
            AutoCreateType = rsChannel("AutoCreateType")
            StructureType = rsChannel("StructureType")
            ListFileType = rsChannel("ListFileType")
            FileNameType = rsChannel("FileNameType")
            FileExt_Index = arrFileExt(rsChannel("FileExt_Index"))
            FileExt_List = arrFileExt(rsChannel("FileExt_List"))
            FileExt_Item = arrFileExt(rsChannel("FileExt_Item"))
            UpdatePages = PE_CLng1(rsChannel("UpdatePages"))
            TopMenuType = rsChannel("TopMenuType")
            ClassGuideType = rsChannel("ClassGuideType")
            arrEnabledTabs = rsChannel("arrEnabledTabs")
            MoneyPerKw = rsChannel("MoneyPerKw")
            EmailOfReject = Replace(rsChannel("EmailOfReject") & "", vbCrLf, "\n")
            EmailOfPassed = Replace(rsChannel("EmailOfPassed") & "", vbCrLf, "\n")
            CommandChannelPoint = PE_Clng(rsChannel("CommandChannelPoint"))
            ModuleType = rsChannel("ModuleType")
            MaxPerLine = rsChannel("MaxPerLine")

            Template_Index = rsChannel("Template_Index")
            ItemCount_Channel = rsChannel("ItemCount")
            ItemChecked_Channel = rsChannel("ItemChecked")
            CommentCount_Channel = rsChannel("CommentCount")
            SpecialCount_Channel = rsChannel("SpecialCount")
            If IsNull(ItemCount_Channel) Then ItemCount_Channel = 0
            If IsNull(ItemChecked_Channel) Then ItemChecked_Channel = 0
            If IsNull(CommentCount_Channel) Then CommentCount_Channel = 0
            If IsNull(SpecialCount_Channel) Then SpecialCount_Channel = 0

            JS_SpecialNum = rsChannel("JS_SpecialNum")
            MaxFileSize = rsChannel("MaxFileSize")
            Fields_Options = rsChannel("Fields_Options")

            '只使用绝对地址时，才使用频道子域名
            If IsNull(rsChannel("LinkUrl")) Or Trim(rsChannel("LinkUrl")) = "" Or Left(strInstallDir, 7) <> "http://" Then
                ChannelUrl = strInstallDir & ChannelDir
            Else
                ChannelUrl = rsChannel("LinkUrl")
            End If
            If Right(ChannelUrl, 1) = "/" Then
                ChannelUrl = Left(ChannelUrl, Len(ChannelUrl) - 1)
            End If
            'If SystemDatabaseType = "SQL" Then
                ChannelUrl_ASPFile = ChannelUrl
            'Else
            '    ChannelUrl_ASPFile = strInstallDir & ChannelDir
            'End If
            If ChannelPurview > 0 Then UseCreateHTML = 0
            Select Case ModuleType
            Case 1
                ModuleName = "Article"
                SheetName = "PE_Article"
            Case 2
                ModuleName = "Soft"
                SheetName = "PE_Soft"
            Case 3
                ModuleName = "Photo"
                SheetName = "PE_Photo"
            Case 5
                ModuleName = "Product"
                SheetName = "PE_Product"
            Case 6
                ModuleName = "Supply"
                SheetName = "PE_Supply"
            End Select
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If
End Sub
%>
