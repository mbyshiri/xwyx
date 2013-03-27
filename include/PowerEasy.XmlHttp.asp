<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'==================================================
'函数名：GetHttpPage
'作  用：获取网页源码
'参  数：HttpUrl ------要获取源码的网页地址
'      ：Coding  ------编码， 1 GB 2 UTF
'==================================================
Function GetHttpPage(HttpUrl, Coding)
    On Error Resume Next
    If IsNull(HttpUrl) = True Or Len(HttpUrl) < 18 Or HttpUrl = "" Then
        GetHttpPage = ""
        Exit Function
    End If
    Dim Http
    Set Http = Server.CreateObject("MSXML2.XMLHTTP")
    Http.Open "GET", HttpUrl, False
    Http.Send
    If Http.Readystate <> 4 Then
        GetHttpPage = ""
        Exit Function
    End If
    If Coding = 1 Then
        GetHttpPage = BytesToBstr(Http.ResponseBody, "UTF-8")
    ElseIf Coding = 2 Then
        GetHttpPage = BytesToBstr(Http.ResponseBody, "Big5")
    Else
        GetHttpPage = BytesToBstr(Http.ResponseBody, "GB2312")
    End If
    
    Set Http = Nothing
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

'==================================================
'函数名：PostHttpPage
'作  用：登录
'==================================================
Function PostHttpPage(RefererUrl, PostUrl, PostData, Coding)
    On Error Resume Next
    Dim xmlHttp
    Dim RetStr
    Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
    xmlHttp.Open "POST", PostUrl, False
    xmlHttp.setRequestHeader "Content-Length", Len(PostData)
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlHttp.setRequestHeader "Referer", RefererUrl
    xmlHttp.Send PostData
    If Err Then
        Set xmlHttp = Nothing
        PostHttpPage = "$False$"
        Exit Function
    End If
    If Coding = 1 Then
        PostHttpPage = BytesToBstr(xmlHttp.ResponseBody, "UTF-8")
    ElseIf Coding = 2 Then
        PostHttpPage = BytesToBstr(xmlHttp.ResponseBody, "Big5")
    Else
        PostHttpPage = BytesToBstr(xmlHttp.ResponseBody, "GB2312")
    End If
    
    Set xmlHttp = Nothing
End Function

'==================================================
'函数名：BytesToBstr
'作  用：将获取的源码转换为中文
'参  数：Body ------要转换的变量
'参  数：Cset ------要转换的类型
'==================================================
Function BytesToBstr(Body, Cset)
    Dim Objstream
    Set Objstream = Server.CreateObject("adodb.stream")
    Objstream.Type = 1
    Objstream.Mode = 3
    Objstream.Open
    Objstream.Write Body
    Objstream.Position = 0
    Objstream.Type = 2
    Objstream.Charset = Cset
    BytesToBstr = Objstream.ReadText
    Objstream.Close
    Set Objstream = Nothing
End Function

'==================================================
'函数名：GetBody
'作  用：截取字符串
'参  数：ConStr ------将要截取的字符串
'参  数：StartStr ------开始字符串
'参  数：OverStr ------结束字符串
'参  数：IncluL ------是否包含StartStr
'参  数：IncluR ------是否包含OverStr
'==================================================
Function GetBody(ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or IsNull(ConStr) = True Or StartStr = "" Or IsNull(StartStr) = True Or OverStr = "" Or IsNull(OverStr) = True Then
        GetBody = "$False$"
        Exit Function
    End If
    Dim Start, Over

    Start = InStrB(1, ConStr, StartStr, vbBinaryCompare)

    If Start <= 0 Then
        Start = InStrB(1, ConStr, Replace(StartStr, vbCrLf, Chr(10)), vbBinaryCompare)
        If Start <= 0 Then
            Start = InStrB(1, ConStr, Replace(StartStr, vbCrLf, Chr(13)), vbBinaryCompare)
            If Start <= 0 Then
                GetBody = "$False$"
                Exit Function
            Else
                If IncluL = False Then
                    Start = Start + LenB(StartStr)
                End If
            End If
        Else
            If IncluL = False Then
                Start = Start + LenB(StartStr)
            End If
        End If
    Else
        If IncluL = False Then
            Start = Start + LenB(StartStr)
        End If
    End If

    Over = InStrB(Start, ConStr, OverStr, vbBinaryCompare)
    If Over <= 0 Or Over <= Start Then
        Over = InStrB(Start, ConStr, Replace(OverStr, vbCrLf, Chr(10)), vbBinaryCompare)
        If Over <= 0 Or Over <= Start Then
            Over = InStrB(Start, ConStr, Replace(OverStr, vbCrLf, Chr(13)), vbBinaryCompare)
            If Over <= 0 Or Over <= Start Then
                GetBody = "$False$"
                Exit Function
            Else
                If IncluR = True Then
                    Over = Over + LenB(OverStr)
                End If
            End If
        Else
            If IncluR = True Then
                Over = Over + LenB(OverStr)
            End If
        End If
    Else
        If IncluR = True Then
            Over = Over + LenB(OverStr)
        End If
    End If

    GetBody = MidB(ConStr, Start, Over - Start)
End Function


'==================================================
'函数名：ReplaceRemoteUrl
'作  用：替换字符串中的远程文件为本地文件并保存远程文件
'参  数：strContent ------ 要替换的字符串
'==================================================
Function ReplaceRemoteUrl(ByVal strContent)
    If IsObjInstalled("Microsoft.XMLHTTP") = False Or ObjInstalled_FSO = False Then
        ReplaceRemoteUrl = strContent
        Exit Function
    End If
    Dim RemoteFiles, RemoteFile, RemoteFileUrl, SaveFilePath, SavePath, SavePath2, SaveFileName, ThumbFileName, SaveFileType, arrSaveFileName, ranNum, dtNow, FileCount, SavedFiles
    Dim temptime, FilesArray, tempi
    If fso.FolderExists(Server.MapPath(InstallDir)) = False Then fso.CreateFolder Server.MapPath(InstallDir)
    If fso.FolderExists(Server.MapPath(InstallDir & ChannelDir)) = False Then fso.CreateFolder Server.MapPath(InstallDir & ChannelDir)
    SavePath = InstallDir & ChannelDir & "/" & UploadDir        '文件保存的本地路径
    If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)
    SavePath = SavePath & "/"
    
    FileCount = 0
    SavedFiles = "|"
    tempi = 0
    regEx.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}([\w\-]+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|jpeg|jpe|bmp|png)))"
    Set RemoteFiles = regEx.Execute(strContent)

    Dim AddWatermark, AddThumb, IsThumb
    If Trim(Request.Form("AddWatermark")) = "Yes" Then
        AddWatermark = True
    Else
        AddWatermark = False
    End If
    If Trim(Request.Form("AddThumb")) = "Yes" Then
        AddThumb = True
    Else
        AddThumb = False
    End If
    

    For Each RemoteFile In RemoteFiles
        IsThumb = False
        RemoteFileUrl = RemoteFile.value
        If InStr(SavedFiles, "|" & RemoteFileUrl & "|") > 0 Then
            '如果已经保存则不进行处理
        Else
            If FileCount = 0 Then
                Response.Write "<b>正在保存远程文件……请稍候！<font color='red'>在此过程中请勿刷新页面！</font></b> "
                Response.Flush
            End If

            SavedFiles = SavedFiles & RemoteFileUrl & "|"
            dtNow = Now()
            arrSaveFileName = Split(RemoteFileUrl, ".")
            SaveFileType = arrSaveFileName(UBound(arrSaveFileName))
            SavePath2 = Year(dtNow) & Right("0" & Month(dtNow), 2)
            If fso.FolderExists(Server.MapPath(SavePath & SavePath2)) = False Then fso.CreateFolder Server.MapPath(SavePath & SavePath2)
            SavePath2 = SavePath2 & "/"
            SaveFilePath = SavePath & SavePath2
            
            Randomize
            ranNum = Int(900 * Rnd) + 100
            temptime = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & Right("0" & Second(dtNow), 2) & ranNum
            SaveFileName = temptime & "." & SaveFileType
            ThumbFileName = temptime & "_S." & SaveFileType
            If SaveRemoteFile(RemoteFileUrl, SaveFilePath & SaveFileName) = True Then
                strContent = Replace(strContent, RemoteFileUrl, "[InstallDir_ChannelDir]{$UploadDir}/" & SavePath2 & SaveFileName)
                If PhotoObject = 1 Then
                    Dim PE_Thumb
                    Set PE_Thumb = New CreateThumb
                    If tempi = 0 And AddThumb = True Then
                        If PE_Thumb.CreateThumb(SaveFilePath & SaveFileName, SaveFilePath & ThumbFileName, 0, 0) = True Then
                            IsThumb = True
                        End If
                    End If
                    If AddWatermark = True Then
                        Call PE_Thumb.AddWatermark(SaveFilePath & SaveFileName)
                    End If
                    Set PE_Thumb = Nothing
                End If

                If IsThumb = True Then
                    UploadFiles = SavePath2 & ThumbFileName & "|" & SavePath2 & SaveFileName
                Else
                    If UploadFiles = "" Then
                        UploadFiles = SavePath2 & SaveFileName
                    Else
                        UploadFiles = UploadFiles & "|" & SavePath2 & SaveFileName
                    End If
                End If
                If PE_CLng(Trim(Request.Form("IncludePic"))) = 0 Then
                    If FileCount > 0 Then
                        IncludePic = 2
                    Else
                        IncludePic = 1
                    End If
                Else
                    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
                End If

                If InStr(UploadFiles, "|") = 0 Then
                    DefaultPicUrl = UploadFiles
                Else
                    FilesArray = Split(UploadFiles, "|")
                    DefaultPicUrl = FilesArray(0)
                End If
                FileCount = FileCount + 1
            End If
            tempi = tempi + 1
            Response.Write "·"
            Response.Flush
        End If
    Next
    If FileCount > 0 Then Response.Write " <b><font color='blue'>共成功保存了 " & FileCount & " 张远程图片！</font></b><br>"
    ReplaceRemoteUrl = strContent
End Function

'==================================================
'函数名：SaveRemoteFile
'作  用：保存远程的文件到本地
'参  数：LocalFileName ------ 本地文件名
'        RemoteFileUrl ------ 远程文件URL
'返回值：True ----- 保存成功
'       False ----- 保存失败
'==================================================
Function SaveRemoteFile(RemoteFileUrl, LocalFileName)
    On Error Resume Next

    Dim Ads, Retrieval, GetRemoteData
    Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
    With Retrieval
        .Open "Get", RemoteFileUrl, False, "", ""
        .Send
        GetRemoteData = .ResponseBody
    End With
    If Err.Number <> 0 Then
        Err.Clear
        Response.Write "<br>" & RemoteFileUrl & " Get Failed"
        SaveRemoteFile = False
        Exit Function
    End If
    Set Retrieval = Nothing
    Set Ads = Server.CreateObject("Adodb.Stream")
    With Ads
        .Type = 1
        .Open
        .Write GetRemoteData
        .SaveToFile Server.MapPath(LocalFileName), 2
        .Cancel
        .Close
    End With
    Set Ads = Nothing
    If Err.Number <> 0 Then
        Err.Clear
        Response.Write "<br>" & LocalFileName & " Save Failed"
        SaveRemoteFile = False
    Else
        SaveRemoteFile = True
    End If
End Function

'=================================================
'方法名：ReplaceStringPath()
'作  用：区域采集内容连接替换
'=================================================
Function ReplaceStringPath(ByVal AreaCode, ByVal AreaUrl, ByVal UpFileType)
    If IsNull(AreaCode) = True Then
        ReplaceStringPath = ""
    End If
    Dim strTemp, strTemp2, strTemp3
    
    regEx.Pattern = "(value|src|href)(\s*=)(.[^\<]*)(\.)(" & UpFileType & ")"
    Set Matches = regEx.Execute(AreaCode)
    For Each Match In Matches
        regEx.Pattern = "(value|src|href)(\s*=)"
        Set strTemp = regEx.Execute(Match.value)
        For Each Match2 In strTemp
            strTemp2 = Match2.value
        Next
        regEx.Pattern = "(value|src|href)(\s*=)"
        strTemp = regEx.Replace(Match.value, "")
    
        If Left(strTemp, 1) = "'" Then
            strTemp3 = "'"
        ElseIf Left(strTemp, 1) = """" Then
            strTemp3 = """"
        End If
        strTemp = regEx.Replace(strTemp, "")
        strTemp = Replace(strTemp, """", "")
        strTemp = Replace(strTemp, "'", "")
        AreaCode = Replace(AreaCode, Match.value, strTemp2 & strTemp3 & DefiniteUrl(strTemp, AreaUrl))
    Next
    ReplaceStringPath = AreaCode
End Function

'==================================================
'函数名：DefiniteUrl
'作  用：将相对地址转换为绝对地址
'参  数：PrimitiveUrl ------要转换的相对地址
'参  数：ConsultUrl ------当前网页地址
'==================================================
Function DefiniteUrl(ByVal PrimitiveUrl, ByVal ConsultUrl)
    Dim ConTemp, PriTemp, Pi, Ci, PriArray, ConArray
    If PrimitiveUrl = "" Or ConsultUrl = "" Or PrimitiveUrl = "$False$" Or ConsultUrl = "$False$" Then
        DefiniteUrl = "$False$"
        Exit Function
    End If
    If Left(LCase(ConsultUrl), 7) <> "http://" Then
        ConsultUrl = "http://" & ConsultUrl
    End If
    ConsultUrl = Replace(ConsultUrl, "\", "/")
    ConsultUrl = Replace(ConsultUrl, "://", ":\\")
    PrimitiveUrl = Replace(PrimitiveUrl, "\", "/")
   
    If Right(ConsultUrl, 1) <> "/" Then
        If InStr(ConsultUrl, "/") > 0 Then
            If InStr(Right(ConsultUrl, Len(ConsultUrl) - InStrRev(ConsultUrl, "/")), ".") > 0 Then
            Else
                ConsultUrl = ConsultUrl & "/"
            End If
        Else
            ConsultUrl = ConsultUrl & "/"
        End If
    End If
    ConArray = Split(ConsultUrl, "/")

    If Left(LCase(PrimitiveUrl), 7) = "http://" Then
        DefiniteUrl = Replace(PrimitiveUrl, "://", ":\\")
    ElseIf Left(PrimitiveUrl, 1) = "/" Then
        DefiniteUrl = ConArray(0) & PrimitiveUrl
    ElseIf Left(PrimitiveUrl, 2) = "./" Then
        PrimitiveUrl = Right(PrimitiveUrl, Len(PrimitiveUrl) - 2)
        If Right(ConsultUrl, 1) = "/" Then
            DefiniteUrl = ConsultUrl & PrimitiveUrl
        Else
            DefiniteUrl = Left(ConsultUrl, InStrRev(ConsultUrl, "/")) & PrimitiveUrl
        End If
    ElseIf Left(PrimitiveUrl, 3) = "../" Then
        Do While Left(PrimitiveUrl, 3) = "../"
            PrimitiveUrl = Right(PrimitiveUrl, Len(PrimitiveUrl) - 3)
            Pi = Pi + 1
        Loop
        For Ci = 0 To (UBound(ConArray) - 1 - Pi)
            If DefiniteUrl <> "" Then
                DefiniteUrl = DefiniteUrl & "/" & ConArray(Ci)
            Else
                DefiniteUrl = ConArray(Ci)
            End If
        Next
        DefiniteUrl = DefiniteUrl & "/" & PrimitiveUrl
    Else
        If InStr(PrimitiveUrl, "/") > 0 Then
            PriArray = Split(PrimitiveUrl, "/")
            If InStr(PriArray(0), ".") > 0 Then
                If Right(PrimitiveUrl, 1) = "/" Then
                    DefiniteUrl = "http:\\" & PrimitiveUrl
                Else
                    If InStr(PriArray(UBound(PriArray) - 1), ".") > 0 Then
                        DefiniteUrl = "http:\\" & PrimitiveUrl
                    Else
                        DefiniteUrl = "http:\\" & PrimitiveUrl & "/"
                    End If
                End If
            Else
                If Right(ConsultUrl, 1) = "/" Then
                    DefiniteUrl = ConsultUrl & PrimitiveUrl
                Else
                    DefiniteUrl = Left(ConsultUrl, InStrRev(ConsultUrl, "/")) & PrimitiveUrl
                End If
            End If
        Else
            If InStr(PrimitiveUrl, ".") > 0 Then
                If Right(ConsultUrl, 1) = "/" Then
                    If Right(LCase(PrimitiveUrl), 3) = ".cn" Or Right(LCase(PrimitiveUrl), 3) = "com" Or Right(LCase(PrimitiveUrl), 3) = "net" Or Right(LCase(PrimitiveUrl), 3) = "org" Then
                        DefiniteUrl = "http:\\" & PrimitiveUrl & "/"
                    Else
                        DefiniteUrl = ConsultUrl & PrimitiveUrl
                    End If
                Else
                    If Right(LCase(PrimitiveUrl), 3) = ".cn" Or Right(LCase(PrimitiveUrl), 3) = "com" Or Right(LCase(PrimitiveUrl), 3) = "net" Or Right(LCase(PrimitiveUrl), 3) = "org" Then
                        DefiniteUrl = "http:\\" & PrimitiveUrl & "/"
                    Else
                        DefiniteUrl = Left(ConsultUrl, InStrRev(ConsultUrl, "/")) & "/" & PrimitiveUrl
                    End If
                End If
            Else
                If Right(ConsultUrl, 1) = "/" Then
                    DefiniteUrl = ConsultUrl & PrimitiveUrl & "/"
                Else
                    DefiniteUrl = Left(ConsultUrl, InStrRev(ConsultUrl, "/")) & "/" & PrimitiveUrl & "/"
                End If
            End If
        End If
    End If
    If Left(DefiniteUrl, 1) = "/" Then
        DefiniteUrl = Right(DefiniteUrl, Len(DefiniteUrl) - 1)
    End If
    If DefiniteUrl <> "" Then
        DefiniteUrl = Replace(DefiniteUrl, "//", "/")
        DefiniteUrl = Replace(DefiniteUrl, ":\\", "://")
    Else
        DefiniteUrl = "$False$"
    End If
End Function

%>
