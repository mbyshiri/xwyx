<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const MaxTotalSize = 104857600    '上传数据限制，最大上传100M
Const NoAllowExt = "asa|asax|ascs|ashx|asmx|asp|aspx|axd|cdx|cer|config|cs|csproj|idc|licx|rem|resources|resx|shtm|shtml|soap|stm|vb|vbproj|vsdisco|webinfo"    '不允许上传类型(黑名单)
Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip"
Dim uEnableUpload, uMaxFileSize, AdminLogined
Dim IsUploadAnonymous, Anonymous
Sub Execute()
    If ObjInstalled_FSO = False Then
        Response.Write "您的服务器不支持FSO，或者FSO已经改名，所以不能上传！"
        Exit Sub
    End If	    
    If CheckLogin() = False and ShowAnonymous = False Then
        Response.Write "请先登录！"
        Exit Sub
    End If
    
    Dim Forms, Files
    Dim oUpFilestream   '上传的数据流
    
    '********************************************
    '以下代码是对提交的数据进行分析
    '********************************************
    Dim RequestBinDate, sSpace, bCrLf, sInfo, iInfoStart, iInfoEnd, tStream, iStart
    Dim sFormValue, sFileName
    Dim iFindStart, iFindEnd
    Dim iFormstart, iFormEnd, sFormName
    Dim FileInfo(6)
        '代码开始
    If Request.TotalBytes < 1 Then  '如果没有数据上传
        FoundErr = True
        ErrMsg = "没有数据上传"
        Exit Sub
    End If
        
    If Request.TotalBytes > MaxTotalSize Then  '如果上传的数据超出限制大小
        FoundErr = True
        ErrMsg = "上传的数据超出限制大小"
        Exit Sub
    End If
        
    Set Forms = Server.CreateObject("Scripting.Dictionary")
    Forms.CompareMode = 1
    Set Files = Server.CreateObject("Scripting.Dictionary")
    Files.CompareMode = 1
    Set tStream = Server.CreateObject("ADODB.Stream")
    Set oUpFilestream = Server.CreateObject("ADODB.Stream")
    oUpFilestream.Type = 1
    oUpFilestream.Mode = 3
    oUpFilestream.Open
    oUpFilestream.Write Request.BinaryRead(Request.TotalBytes)
    oUpFilestream.Position = 0
    RequestBinDate = oUpFilestream.Read
    iFormEnd = oUpFilestream.size
    bCrLf = ChrB(13) & ChrB(10)
    '取得每个项目之间的分隔符
    sSpace = LeftB(RequestBinDate, InStrB(1, RequestBinDate, bCrLf) - 1)
    iStart = LenB(sSpace)
    iFormstart = iStart + 2
    '分解项目
    
    Do
        iInfoEnd = InStrB(iFormstart, RequestBinDate, bCrLf & bCrLf) + 3
        tStream.Type = 1
        tStream.Mode = 3
        tStream.Open
        oUpFilestream.Position = iFormstart
        oUpFilestream.CopyTo tStream, iInfoEnd - iFormstart
        tStream.Position = 0
        tStream.Type = 2
        tStream.Charset = "gb2312"
        sInfo = tStream.ReadText
                
        '取得表单项目名称
        iFindStart = InStr(22, sInfo, "name=""", 1) + 6
        iFindEnd = InStr(iFindStart, sInfo, """", 1)
        sFormName = Mid(sInfo, iFindStart, iFindEnd - iFindStart)
            
        iFormstart = InStrB(iInfoEnd, RequestBinDate, sSpace) - 1
        If InStr(45, sInfo, "filename=""", 1) > 0 Then   '如果是文件
            '取得文件属性
            iFindStart = InStr(iFindEnd, sInfo, "filename=""", 1) + 10
            iFindEnd = InStr(iFindStart, sInfo, """" & vbCrLf, 1)
            sFileName = Mid(sInfo, iFindStart, iFindEnd - iFindStart)
            FileInfo(0) = sFormName
            FileInfo(1) = GetFileName(sFileName)
            FileInfo(2) = GetFilePath(sFileName)
            FileInfo(3) = GetFileExt(sFileName)
            iFindStart = InStr(iFindEnd, sInfo, "Content-Type: ", 1) + 14
            iFindEnd = InStr(iFindStart, sInfo, vbCr)
            FileInfo(4) = Mid(sInfo, iFindStart, iFindEnd - iFindStart)
            FileInfo(5) = iInfoEnd
            FileInfo(6) = iFormstart - iInfoEnd - 2
            Files.Add sFormName, FileInfo
        Else    '如果是表单项目
            tStream.Close
            tStream.Type = 1
            tStream.Mode = 3
            tStream.Open
            oUpFilestream.Position = iInfoEnd
            oUpFilestream.CopyTo tStream, iFormstart - iInfoEnd - 2
            tStream.Position = 0
            tStream.Type = 2
            tStream.Charset = "gb2312"
            sFormValue = tStream.ReadText
            If Forms.Exists(sFormName) Then
                Forms(sFormName) = Forms(sFormName) & ", " & sFormValue
            Else
                Forms.Add sFormName, sFormValue
            End If
        End If
        tStream.Close
        iFormstart = iFormstart + iStart + 2
        '如果到文件尾了就退出
    Loop Until (iFormstart + 2) >= iFormEnd
    RequestBinDate = ""
    Set tStream = Nothing
    '********************************************
    '数据分析结束
    '********************************************
    
    Dim EnableUploadFile, MaxFileSize, UpFileType, SavePath, dirMonth, tmpPath
    If fso.FolderExists(Server.MapPath(InstallDir)) = False Then fso.CreateFolder Server.MapPath(InstallDir)
        
    Dim FileType, Uname, checkuserrs, MaxSpaceSize,FieldName
    FileType = LCase(Trim(Forms("FileType")))
    FieldName = Trim(Forms("FieldName"))	
    MaxSpaceSize = PE_CLng(Trim(Forms("size")))
    Anonymous = PE_CLng(Trim(Forms("Anonymous")))
    If Anonymous = 1  Then
        If ShowAnonymous = True Then
            Dim  rsGroup
            Set rsGroup = Conn.Execute("select * from PE_UserGroup where GroupID=-1")
            arrClass_Input = Trim(rsGroup("arrClass_Input"))
            UserSetting = Split(Trim(rsGroup("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
            rsGroup.Close
            Set rsGroup = Nothing	
            uMaxFileSize = PE_CLng(UserSetting(10))		
            If PE_CBool(PE_CLng(UserSetting(9))) = True Then
                IsUploadAnonymous = True
            Else
                IsUploadAnonymous = False	
            End If
        Else
            IsUploadAnonymous = False			
        End If 
    End If
	
    If CheckLogin() = False Then
        If Anonymous = 1 Then
            If IsUploadAnonymous = False Then 
                Response.Write "您没有权限上传！"
                Exit Sub			
            End If
        Else		
            Response.Write "请先登录！"
            Exit Sub
        End If
    End If	
	
    Dim ChannelID, sqlChannel, rsChannel, UploadDir, ModuleType, IsThumb
    
    ChannelID = PE_CLng(Trim(Forms("ChannelID")))
    If ChannelID = 0 Then
        EnableUploadFile = True
        Select Case FileType
        Case "authorpic", "copyfrompic"
            If AdminLogined <> True Then
                Response.Write "请先登录！"
                Exit Sub
            End If
            UploadDir = FileType & "/"
            SavePath = InstallDir & UploadDir
            UpFileType = "gif|jpg|jpeg|jpe|bmp|png|swf"
            MaxFileSize = 2048
            If fso.FolderExists(Server.MapPath(InstallDir & FileType)) = False Then fso.CreateFolder Server.MapPath(InstallDir & FileType)
        Case "producerpic", "trademarkpic"
            If AdminLogined <> True Then
                Response.Write "请先登录！"
                Exit Sub
            End If
            UploadDir = "Shop/" & FileType & "/"
            SavePath = InstallDir & UploadDir
            UpFileType = "gif|jpg|jpeg|jpe|bmp|png|swf"
            MaxFileSize = 2048
            If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)
        Case "adminblogpic"
            If AdminLogined <> True Then
                Response.Write "请先登录！"
                Exit Sub
            End If
            Uname = ReplaceBadChar(Trim(Forms("Uname")))
            Set checkuserrs = Conn.Execute("select UserID,UserName from PE_User where UserName='" & Uname & "'")
            If checkuserrs.BOF And checkuserrs.EOF Then
                Response.Write "用户验证错"
                checkuserrs.Close
                Set checkuserrs = Nothing
                Exit Sub
            Else
                UploadDir = "Space/" & Uname & checkuserrs("UserID") & "/"
                checkuserrs.Close
                Set checkuserrs = Nothing
                SavePath = InstallDir & UploadDir
                UpFileType = "gif|jpg|jpeg|jpe|bmp|png|swf"
                MaxFileSize = 2048
                If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)
            End If
        Case "userblogpic"
            uEnableUpload = True
            Set checkuserrs = Conn.Execute("select UserID,UserName from PE_User where UserName='" & UserName & "'")
            If checkuserrs.BOF And checkuserrs.EOF Then
                Response.Write "用户验证错"
                checkuserrs.Close
                Set checkuserrs = Nothing
                Exit Sub
            Else
                UploadDir = "Space/" & UserName & checkuserrs("UserID") & "/"
                checkuserrs.Close
                Set checkuserrs = Nothing
                SavePath = InstallDir & UploadDir
                UpFileType = "gif|jpg|jpeg|jpe|bmp|png"
                MaxFileSize = 2048
                If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)
                Dim ft, foldersize, realsize
                Set ft = fso.GetFolder(Server.MapPath(SavePath))
                foldersize = ft.size
                If foldersize = 0 Then foldersize = 1
                realsize = foldersize / 1048576
                If realsize > MaxSpaceSize Then
                    Response.Write "您的空间已满,请清理后再上传！"
                    Exit Sub
            End If
                Set ft = Nothing
            End If
        Case "Intervieweepic" '上传简历相片2006-1-13
            uEnableUpload = True
            UploadDir = "UploadPhoto/"
            SavePath = InstallDir & UploadDir
            UpFileType = "gif|jpg|jpeg|jpe|bmp|png"
            MaxFileSize = 2048
            If fso.FolderExists(Server.MapPath(InstallDir & "UploadPhoto")) = False Then fso.CreateFolder Server.MapPath(InstallDir & "UploadPhoto")
        Case "adpic"
            SavePath = InstallDir & ADDir & "/UploadADPic/"
            UpFileType = "gif|jpg|jpeg|jpe|bmp|png|swf"
            MaxFileSize = 2048
            If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)				
        Case Else
            Response.Write "频道参数丢失！"
            Exit Sub
        End Select
    ElseIf ChannelID < 0 Then
        Dim tempFileType
        tempFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")		
        Select Case FileType
        Case "pic", "batchpic", "softpic", "intervieweepic", "fieldpic"
            UpFileType = Trim(tempFileType(0))
        Case "photo", "photos"
            UpFileType = Trim(tempFileType(0)) & "|" & Trim(tempFileType(1))
        Case "flash"
            UpFileType = Trim(tempFileType(1))
        Case "media"
            UpFileType = Trim(tempFileType(2))
        Case "real"
            UpFileType = Trim(tempFileType(3))
        Case "fujian","fieldsoft"
            UpFileType = Trim(tempFileType(4))
        Case "soft"
            UpFileType = Trim(tempFileType(1)) & "|" & Trim(tempFileType(2)) & "|" & Trim(tempFileType(3)) & "|" & Trim(tempFileType(4))
        Case Else
            UpFileType = ""
        End Select	
        Select Case ChannelID			
        Case -1
            EnableUploadFile = True
            SavePath = InstallDir &"Others/Announce/"
            MaxFileSize = 2048
            If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)		
        Case -2
            EnableUploadFile = True		
            SavePath = InstallDir &"Others/Email/"
            MaxFileSize = 20480
            If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)		
        Case -3 
            EnableUploadFile = True		
            SavePath = InstallDir &"Others/Message/"
            MaxFileSize = 2048
            If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)		
        Case Else
            EnableUpload = False	
        End Select
    Else	
        sqlChannel = "select ChannelID,ChannelName,ChannelDir,ModuleType,Disabled,EnableUploadFile,UploadDir,MaxFileSize,UpFileType from PE_Channel where ChannelID=" & ChannelID
        Set rsChannel = Server.CreateObject("adodb.recordset")
        rsChannel.Open sqlChannel, Conn, 1, 1
        If rsChannel.BOF And rsChannel.EOF Then
            Response.Write "找不到此频道"
            FoundErr = True
            rsChannel.Close
            Set rsChannel = Nothing
            Exit Sub
        End If
        If rsChannel("Disabled") = True Then
            Response.Write "此频道已经被禁用！"
            FoundErr = True
        Else
            EnableUploadFile = rsChannel("EnableUploadFile")
            MaxFileSize = rsChannel("MaxFileSize")
            UpFileType = rsChannel("UpFileType")
            Dim arrFileType
            If UpFileType = "" Then
                arrFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")
            Else
                arrFileType = Split(UpFileType, "$")
                If UBound(arrFileType) < 4 Then
                    arrFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")
                End If
            End If
            ModuleType = rsChannel("ModuleType")
            Select Case ModuleType
            Case 0, 1, 3, 4, 5, 6, 7, 8 '2006-1-13
                UploadDir = rsChannel("UploadDir") & "/"
            Case 2
                If FileType = "softpic" Or FileType = "pic" Then '软件简介图片上传
                    UploadDir = "UploadSoftPic/"
                Else
                    UploadDir = rsChannel("UploadDir") & "/"
                End If
            End Select

            SavePath = InstallDir & rsChannel("ChannelDir") & "/" & UploadDir

            Select Case FileType
            Case "pic", "batchpic", "softpic", "intervieweepic", "fieldpic"
                UpFileType = Trim(arrFileType(0))
            Case "photo", "photos"
                UpFileType = Trim(arrFileType(0)) & "|" & Trim(arrFileType(1))
            Case "flash"
                UpFileType = Trim(arrFileType(1))
            Case "media"
                UpFileType = Trim(arrFileType(2))
            Case "real"
                UpFileType = Trim(arrFileType(3))
            Case "fujian","fieldsoft"
                UpFileType = Trim(arrFileType(4))
            Case "soft"
                UpFileType = Trim(arrFileType(1)) & "|" & Trim(arrFileType(2)) & "|" & Trim(arrFileType(3)) & "|" & Trim(arrFileType(4))
            Case Else
                UpFileType = ""
            End Select
            If fso.FolderExists(Server.MapPath(InstallDir & rsChannel("ChannelDir"))) = False Then fso.CreateFolder Server.MapPath(InstallDir & rsChannel("ChannelDir"))
            If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)
            
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    If IsUploadAnonymous = True and Anonymous = 1 and ShowAnonymous = True Then uEnableUpload = True
    If uEnableUpload = False Then EnableUploadFile = False
    If MaxFileSize > uMaxFileSize Then MaxFileSize = uMaxFileSize
    If EnableUploadFile = False Then
        Response.Write "本频道未开放文件上传功能"
        FoundErr = True
    End If
    
    If FoundErr = True Then Exit Sub
    
    Response.Write "<html>" & vbCrLf
    Response.Write "<head>" & vbCrLf
    Response.Write "<title>上传文件结果</title>" & vbCrLf
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link rel='stylesheet' type='text/css' href='../Editor/editor_dialog.css'>" & vbCrLf
    Response.Write "</head>" & vbCrLf
    Response.Write "<body leftmargin='5' topmargin='0'>" & vbCrLf
    
    
    Dim EnableUpload, msg
    Dim AddWatermark, EnableCreateThumb, ThumbWidth, ThumbHeight, LinkUrl
    Dim strJS, dtNow, ranNum, strFileName, strThumbPath, i
    Dim strTemp, FileCount, strUploadPics
    Dim FormNames
    Dim oFileInfo
    Dim cFileName, cFilePath, cFileExt, cFileMIME, cFileStart, cFileSize
    Dim oFilestream
    Dim PE_Thumb
    Set PE_Thumb = New CreateThumb

    FileCount = 0
    FormNames = Files.Keys
    LinkUrl = LCase(Trim(Forms("LinkUrl")))
    
    IsThumb = Trim(Forms("IsThumb"))
    If IsNumeric(IsThumb) Then
       IsThumb = CLng(IsThumb)
    Else
        IsThumb = 0
    End If
    
    For i = 0 To Files.Count - 1
        On Error Resume Next
        EnableUpload = False        
        dtNow = Now()
        oFileInfo = Files.Item(FormNames(i))
        cFileName = oFileInfo(1)
        cFilePath = oFileInfo(2)
        cFileExt = oFileInfo(3)
        cFileMIME = oFileInfo(4)
        cFileStart = oFileInfo(5)
        cFileSize = oFileInfo(6)
                
        If cFileSize < 10 Then
            FoundErr = True
            If Not (FileType = "batchpic" Or FileType = "photos") Then
                msg = "请先选择你要上传的文件！"
            End If
        Else
            If cFileSize > (MaxFileSize * 1024) Then
                If FileType = "batchpic" Then
                    Response.Write "<li>第 " & i + 1 & " 个文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！</li>"
                Else
                    msg = "文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！"
                End If
                FoundErr = True
            Else
                If CheckFileExt(UpFileType, cFileExt) = False Or CheckFileExt(NoAllowExt, cFileExt) = True Or IsValidStr(cFileExt) = False Then
                    FoundErr = True
                    If cFileName <> "" Then
                        If (FileType = "batchpic" Or FileType = "photos") Then
                            Response.Write "<li>第 " & i + 1 & " 个文件不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType & "</li>"
                        Else
                            msg = "这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
                        End If
                    End If
                Else
                    If Left(LCase(cFileMIME), 5) = "text/" And CheckFileExt(NeedCheckFileMimeExt, cFileExt) = True Then
                        FoundErr = True
                        If (FileType = "batchpic" Or FileType = "photos") Then
                            Response.Write "<li>第 " & i + 1 & " 个文件是用文本文件伪造的图片文件或压缩文件，为了系统安全，不允许上传这种类型的文件！</li>"
                        Else
                            msg = "为了系统安全，不允许上传用文本文件伪造的图片文件！"
                        End If
                    Else
                        EnableUpload = True
                    End If
                End If
            End If
        End If
        
        If EnableUpload = True Then
            dirMonth = Year(dtNow) & Right("0" & Month(dtNow), 2) & "/"
            tmpPath = SavePath & dirMonth
            If FileType = "adminblogpic" Or FileType = "userblogpic" Then
                If fso.FolderExists(Server.MapPath(SavePath)) = False Then fso.CreateFolder Server.MapPath(SavePath)
            End If
            If fso.FolderExists(Server.MapPath(tmpPath)) = False Then fso.CreateFolder Server.MapPath(tmpPath)
            
            Randomize
            strFileName = GetNumString()
            tmpPath = tmpPath & strFileName & "." & cFileExt
            
            Set oFilestream = Server.CreateObject("ADODB.Stream")
            oFilestream.Type = 1
            oFilestream.Mode = 3
            oFilestream.Open
            oUpFilestream.Position = cFileStart
            oUpFilestream.CopyTo oFilestream, cFileSize
            oFilestream.SaveToFile Server.MapPath(tmpPath)   '保存文件
            oFilestream.Close
            Set oFilestream = Nothing
            
            FileCount = FileCount + 1
            
            Select Case FileType
            Case "batchpic"
                Response.Write "<li>第 " & i + 1 & " 张图片上传成功！"

                If LinkUrl <> "" And LinkUrl <> "http://" Then strTemp = strTemp & "<a href='" & LinkUrl & "' target='_blank'>"
                strTemp = strTemp & "<img src='" & tmpPath & "'alt='" & Trim(Forms("alttext" & i)) & "'"
                If Trim(Forms("width" & i)) <> "" Then strTemp = strTemp & "width='" & PE_CLng(Trim(Forms("width" & i))) & "'"
                If Trim(Forms("height" & i)) <> "" Then strTemp = strTemp & "height='" & PE_CLng(Trim(Forms("height" & i))) & "'"
                strTemp = strTemp & " border='" & PE_CLng(Trim(Forms("border" & i))) & "'"
                strTemp = strTemp & " style='BORDER-COLOR:" & Trim(Forms("bordercolor" & i)) & "'"
                strTemp = strTemp & " align='" & Trim(Forms("aligntype" & i)) & "'"
                If Trim(Forms("vspace" & i)) <> "" Then strTemp = strTemp & " vspace='" & PE_CLng(Trim(Forms("vspace" & i))) & "'"
                If Trim(Forms("hspace" & i)) <> "" Then strTemp = strTemp & " hspace='" & PE_CLng(Trim(Forms("hspace" & i))) & "'"
                If Trim(Forms("styletype" & i)) <> "" Then strTemp = strTemp & " style='filter:" & Trim(Forms("styletype" & i)) & "'"
                
                If Trim(Forms("zoom" & i)) = "Yes" Then
                    strTemp = strTemp & " onload='resizepic(this)' onmousewheel='return bbimg(this)'"
                End If
                strTemp = strTemp & ">"
                If LinkUrl <> "" And LinkUrl <> "http://" Then strTemp = strTemp & "</a>"
                strTemp = strTemp & "<BR><BR>"
                
                strUploadPics = strUploadPics & "$$$" & dirMonth & strFileName & "." & cFileExt
                If Trim(Forms("AddWatermark" & i)) = "Yes" Then
                    AddWatermark = True
                Else
                    AddWatermark = False
                End If
                If Trim(Forms("CreateThumb" & i)) = "Yes" Then
                    EnableCreateThumb = True
                Else
                    EnableCreateThumb = False
                End If
                ThumbWidth = PE_CLng(Trim(Forms("ThumbWidth" & i)))
                ThumbHeight = PE_CLng(Trim(Forms("ThumbHeight" & i)))
                
                If PhotoObject > 0 And EnableCreateThumb = True Then
                    strThumbPath = SavePath & dirMonth & strFileName & "_S." & cFileExt
                    If PE_Thumb.CreateThumb(tmpPath, strThumbPath, ThumbWidth, ThumbHeight) = True Then
                        FileCount = FileCount + 1
                        strUploadPics = strUploadPics & "$$$" & dirMonth & strFileName & "_S." & cFileExt
                        Response.Write " <FONT color='green'> 创建缩略图成功！</FONT> "
                    End If
                End If
                If PhotoObject > 0 And AddWatermark = True Then
                    If PE_Thumb.AddWatermark(tmpPath) = True Then
                        Response.Write " <FONT color='blue'>生成水印成功！</font> "
                    End If
                End If
                Response.Write "</li>"
            Case "pic"
                strUploadPics = dirMonth & strFileName & "." & cFileExt
                If PhotoObject > 0 Then
                    strThumbPath = SavePath & dirMonth & strFileName & "_S." & cFileExt
                    If PE_Thumb.CreateThumb(tmpPath, strThumbPath, 0, 0) = True Then
                        strUploadPics = strUploadPics & "$$$" & dirMonth & strFileName & "_S." & cFileExt
                        FileCount = FileCount + 1
                    End If
                    Call PE_Thumb.AddWatermark(tmpPath)
                End If
                Response.Write "图片上传成功！ <a href='upload.asp?DialogType=" & FileType & "&ChannelID=" & ChannelID & "&PhotoUpfileType=" & PE_CLng(Trim(Forms("PhotoUpfileType"))) & "'>继续上传</a>" & vbCrLf
                strJS = strJS & "parent.url.value='" & tmpPath & "';" & vbCrLf
                strJS = strJS & "parent.frmPreview.img.src='" & tmpPath & "';" & vbCrLf
                strJS = strJS & "parent.frmPreview.img2.src='" & tmpPath & "';" & vbCrLf
                strJS = strJS & "parent.upfilename.value='" & FileCount & "$$$" & strUploadPics & "';" & vbCrLf
                Exit For
            Case "flash", "media", "real", "fujian"
                Response.Write "文件上传成功！ <a href='upload.asp?DialogType=" & FileType & "&ChannelID=" & ChannelID & "&PhotoUpfileType=" & PE_CLng(Trim(Forms("PhotoUpfileType"))) & "'>继续上传</a>" & vbCrLf
                strJS = strJS & "parent.document.form1.url.value='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                strJS = strJS & "parent.document.form1.UpFileName.value='" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                Exit For
            Case "photos"
                Response.Write "<li>第 " & i + 1 & " 张图片上传成功！</li>" & vbCrLf
                
                strJS = strJS & "  var url" & i & "='图片地址'+(parent.document.myform.PhotoUrl.length+1)+'|" & dirMonth & strFileName & "." & cFileExt & "'; " & vbCrLf
                strJS = strJS & "parent.document.myform.PhotoUrl.options[parent.document.myform.PhotoUrl.length]=new Option(url" & i & ",url" & i & ");" & vbCrLf
                If PhotoObject > 0 Then
                    If IsThumb = i Then
                        strThumbPath = SavePath & dirMonth & strFileName & "_S." & cFileExt
                        If PE_Thumb.CreateThumb(tmpPath, strThumbPath, Thumb_DefaultWidth, Thumb_DefaultHeight) = True Then
                            strJS = strJS & "parent.document.myform.PhotoThumb.value='" & dirMonth & strFileName & "_S." & cFileExt & "'; " & vbCrLf
                        Else
                            strJS = strJS & "parent.document.myform.PhotoThumb.value='" & dirMonth & strFileName & "." & cFileExt & "'; " & vbCrLf
                        End If
                    End If
                    Call PE_Thumb.AddWatermark(tmpPath)
                Else
                    If IsThumb = i Then
                        strJS = strJS & "parent.document.myform.PhotoThumb.value='" & dirMonth & strFileName & "." & cFileExt & "'; " & vbCrLf
                    End If
                End If
            Case "softpic"
                Response.Write "图片上传成功！ <a href='javascript:history.go(-1)'>继续上传</a>"
                strJS = strJS & "parent.document.myform.SoftPicUrl.value='UploadSoftPic/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                If PhotoObject > 0 Then
                    If PE_Thumb.CreateThumb(tmpPath, tmpPath, 0, 0) = True Then
                        FileCount = FileCount + 1
                    End If
                    Call PE_Thumb.AddWatermark(tmpPath)
                End If
                Exit For
            Case "soft"
                Response.Write "文件上传成功！ <a href='javascript:history.go(-1)'>继续上传</a>"
                strJS = strJS & "var url='下载地址'+(parent.document.myform.DownloadUrl.length+1)+'|" & dirMonth & strFileName & "." & cFileExt & "'; " & vbCrLf
                strJS = strJS & "parent.document.myform.DownloadUrl.options[parent.document.myform.DownloadUrl.length]=new Option(url,url);" & vbCrLf
                strJS = strJS & "parent.document.myform.SoftSize.value='" & CStr(Round(cFileSize / 1024)) & "';" & vbCrLf
                Exit For
            Case "authorpic", "copyfrompic"
                Response.Write "文件上传成功！ <a href='javascript:history.go(-1)'>继续上传</a>"
                strJS = strJS & "parent.document.myform.Photo.value='" & InstallDir & FileType & "/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                strJS = strJS & "parent.document.myform.showphoto.src='" & InstallDir & FileType & "/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                Exit For
            Case "producerpic", "trademarkpic"
                Response.Write "文件上传成功！ <a href='javascript:history.go(-1)'>继续上传</a>"
                strJS = strJS & "parent.document.myform.Photo.value='" & InstallDir & "Shop/" & FileType & "/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                strJS = strJS & "parent.document.myform.showphoto.src='" & InstallDir & "Shop/" & FileType & "/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                Exit For
            Case "adminblogpic"
                Response.Write "文件上传成功！ <a href='javascript:history.go(-1)'>继续上传</a>"
                strJS = strJS & "parent.document.myform.Photo.value='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                strJS = strJS & "parent.document.myform.showphoto.src='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                Exit For
            Case "userblogpic"
                If cFileExt = "rar" Or cFileExt = "zip" Or cFileExt = "ace" Then
                    Response.Write "附件上传成功！"
                    strJS = strJS & "parent.document.myform.img.src='../images/rar.gif';" & vbCrLf
                ElseIf cFileExt = "swf" Then
                    Response.Write "FLASH上传成功！"
                    strJS = strJS & "parent.document.myform.img.src='../images/swf.gif';" & vbCrLf
                Else
                    Response.Write "图片上传成功！"
                    strJS = strJS & "parent.document.myform.img.src='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                End If
                strJS = strJS & "parent.document.myform.url.value='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                Exit For
            Case "fieldpic" 
                strJS = strJS & "parent.document.myform."&FieldName&".value='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf	
                strJS = strJS & "history.go(-1);" & vbCrLf							
                Exit For	
            Case "fieldsoft" 
                strJS = strJS & "parent.document.myform."&FieldName&".value='" & SavePath & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf	
                strJS = strJS & "history.go(-1);" & vbCrLf							
                Exit For							
            Case "intervieweepic" '2006-1-13
                Response.Write "照片上传成功！ <a href='javascript:history.go(-1)'>继续上传</a>"
                strJS = strJS & "parent.document.myform.MyPhoto.value='UploadPhotos" & "/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                Exit For
            Case "adpic"
                'Response.Write "文件上传成功！"
                If cFileExt = "swf" Then
                    strJS = strJS & "parent.document.myform.FlashUrl.value='" & InstallDir & ADDir & "/UploadADPic/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                    strJS = strJS & "parent.ADTypeChecked(1);" & vbCrLf
                Else
                    strJS = strJS & "parent.document.myform.ImgUrl.value='" & InstallDir & ADDir & "/UploadADPic/" & dirMonth & strFileName & "." & cFileExt & "';" & vbCrLf
                    strJS = strJS & "parent.ADTypeChecked(0);" & vbCrLf
                End If
                strJS = strJS & "history.go(-1);" & vbCrLf
                Exit For
            End Select
        End If
    Next
    
    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    If FileType = "batchpic" Then
        strTemp = strTemp & "$$$" & FileCount & strUploadPics
        Response.Write "window.returnValue=""" & strTemp & """" & vbCrLf
        Response.Write "setTimeout(""window.close()"", 1000);" & vbCrLf
    Else
        Response.Write strJS
        If FoundErr = True Then
            If msg <> "" Then Response.Write "alert('" & msg & "');" & vbCrLf
            If FileType = "pic" Or FileType = "flash" Or FileType = "media" Or FileType = "real" Or FileType = "fujian" Then
                Response.Write "window.location='upload.asp?DialogType=" & FileType & "&ChannelID=" & ChannelID
                If Anonymous = 1 Then Response.Write "&Anonymous=1"				
                Response.Write "';" & vbCrLf
            Else
                Response.Write "history.go(-1);" & vbCrLf
            End If
        Else
            If FileType = "photos" Then Response.Write "history.go(-1);" & vbCrLf
        End If
    End If
    Response.Write "</script>"
    Response.Write "</body></html>"
    

    Set PE_Thumb = Nothing

    '清除变量及对像
    Forms.RemoveAll
    Set Forms = Nothing
    Files.RemoveAll
    Set Files = Nothing
    oUpFilestream.Close
    Set oUpFilestream = Nothing
    
    Call ClearAspFile(SavePath & dirMonth)
    
End Sub

Sub ShowUploadForm()
    Anonymous = PE_Clng(Request("Anonymous"))
    If Anonymous = 1  Then
        If ShowAnonymous = True Then
            Dim  rsGroup
            Set rsGroup = Conn.Execute("select * from PE_UserGroup where GroupID=-1")
            arrClass_Input = Trim(rsGroup("arrClass_Input"))
            UserSetting = Split(Trim(rsGroup("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
            rsGroup.Close
            Set rsGroup = Nothing	
            uMaxFileSize = PE_CLng(UserSetting(10))		
            If PE_CBool(PE_CLng(UserSetting(9))) = True Then
                IsUploadAnonymous = True
         	Else
 	            IsUploadAnonymous = False	
            End If
       Else
            IsUploadAnonymous = False			
        End If 
    End If
    If CheckLogin() = False Then
        If Anonymous = 1 Then
            If IsUploadAnonymous = False Then 
                Response.Write "您没有上传权限！"
                Exit Sub			
            End If			
        Else
            Response.Write "请先登录！"
            Exit Sub
        End If
	End If
    
    Response.Write "<html>" & vbCrLf
    Response.Write "<head>" & vbCrLf
    Response.Write "<title>上传文件</title>" & vbCrLf
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link rel='stylesheet' type='text/css' href='../editor/editor_dialog.css'>" & vbCrLf
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function check() " & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    var strFileName=document.form1.FileName.value;" & vbCrLf
    Response.Write "    if (strFileName=='')" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert('请选择要上传的文件');" & vbCrLf
    Response.Write "        document.form1.FileName.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
    Response.Write "</head>" & vbCrLf
    Response.Write "<body class='Filebg' leftmargin='0' topmargin='0'><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td>" & vbCrLf
    
    
    Dim ChannelID, sqlChannel, rsChannel, FileType, ModuleType, i, Uname, USpace, FieldName
    ChannelID = PE_CLng(Trim(Request("ChannelID")))
    FileType = LCase(Trim(Request("DialogType")))
    FieldName = Trim(Request("FieldName"))
    If FileType = "userblogpic" Then
        USpace = LCase(Trim(Request("size")))
    End If
    If ChannelID = 0 Then
        Select Case FileType
        Case "authorpic", "copyfrompic", "producerpic", "trademarkpic", "adpic"
            If AdminLogined <> True Then
                Response.Write "请先登录后台！"
                FoundErr = True
            End If
        Case "adminblogpic"
            If AdminLogined <> True Then
                Response.Write "请先登录后台！"
                FoundErr = True
            End If
            Uname = LCase(Trim(Request("Uname")))
        Case "userblogpic"
        Case "intervieweepic"   '2006-1-13
        Case Else
            Response.Write "频道参数丢失！"
            FoundErr = True
        End Select
    Elseif ChannelID > 0 then
        sqlChannel = "select ChannelDir,Disabled,EnableUploadFile,ModuleType from PE_Channel where ChannelID=" & ChannelID
        Set rsChannel = Server.CreateObject("adodb.recordset")
        rsChannel.Open sqlChannel, Conn, 1, 1
        If rsChannel.BOF And rsChannel.EOF Then
            Response.Write "找不到此频道1"
            FoundErr = True
        Else
            If rsChannel("Disabled") = True Then
                Response.Write "此频道已经被禁用！"
                FoundErr = True
            Else
                If rsChannel("EnableUploadFile") = False Then
                    Response.Write "对不起，本频道不允许上传文件！"
                    FoundErr = True
                End If
                If PE_CLng(Trim(Request("Anonymous"))) = 1 Then
                    arrClass_Input = Conn.Execute("SELECT arrClass_Input from PE_UserGroup where GroupID=-1")(0)
                    If FoundInArr(arrClass_Input, rsChannel("ChannelDir") & "none", ",") = True or ShowAnonymous = False Then
                        Response.Write "对不起，本频道没有开启匿名访问功能！"							   
                        FoundErr = True
                    End If
                End If
                ModuleType = rsChannel("ModuleType")
            End If
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    
    If FoundErr <> True Then
        If ModuleType = 3 Then
            Response.Write "<form action='Upfile.asp' method='post' name='form1' enctype='multipart/form-data'>" & vbCrLf
            
            If PE_CLng(Trim(Request("PhotoUpfileType"))) = 0 Then
                For i = 0 To 9
                    Response.Write "  <input name='IsThumb' type='radio' value='" & i & "'"
                    If i = 0 Then Response.Write " checked"
                    Response.Write ">"
                    Response.Write "  <input name='FileName" & i & "' type='FILE' class='FileButton' size='28'>" & vbCrLf
                    If (i + 1) Mod 2 = 0 Then Response.Write "  <br>" & vbCrLf
                Next
                Response.Write "<font style='font-size:9pt'>若选中文件名前的单选框，则表示将此图片设为缩略图。</font>&nbsp;&nbsp;" & vbCrLf
                Response.Write "&nbsp;&nbsp;<input type='submit' name='Submit' value='开始上传'>" & vbCrLf
                Response.Write "  <input name='FileType' type='hidden' id='FileType' value='photos'>" & vbCrLf
            Else
                Response.Write "  <input name='FileName' type='FILE' class='FileButton' size='28'>" & vbCrLf
                Response.Write "  <input type='submit' name='Submit' value='上传'>" & vbCrLf
                Response.Write "  <input name='FileType' type='hidden' id='FileType' value='" & FileType & "'>" & vbCrLf
                Response.Write "  <input name='FieldName' type='hidden' id='FieldName' value='" & FieldName & "'>" & vbCrLf				
            End If
            Response.Write "  <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>" & vbCrLf
            Response.Write "  <input name='PhotoUpfileType' type='hidden' id='PhotoUpfileType' value='" & PE_CLng(Trim(Request("PhotoUpfileType"))) & "'>" & vbCrLf
            Response.Write "</form>" & vbCrLf
        Else
            Response.Write "<form action='Upfile.asp' method='post' name='form1' onSubmit='return check()' enctype='multipart/form-data'>" & vbCrLf
            
            If FileType = "authorpic" Or FileType = "copyfrompic" Or FileType = "producerpic" Or FileType = "trademarkpic" Then
                Response.Write "  <input name='FileName' type='FILE' class='FileButton' size='20'>" & vbCrLf
            ElseIf FileType = "adminblogpic" Then
                Response.Write "  <input name='FileName' type='FILE' class='FileButton' size='20'>" & vbCrLf
                Response.Write "  <input name='Uname' type='hidden' id='Uname' value='" & Uname & "'>" & vbCrLf
            ElseIf FileType = "userblogpic" Then
                Response.Write "  <input name='FileName' type='FILE' class='FileButton' size='20'>" & vbCrLf
                Response.Write "  <input name='size' type='hidden' id='size' value='" & USpace & "'>" & vbCrLf
            Else
                Response.Write "  <input name='FileName' type='FILE' class='FileButton' size='35'>" & vbCrLf
            End If
            Response.Write "  <input type='submit' name='Submit' value='上传'>" & vbCrLf
            If ShowAnonymous = True and Anonymous = 1 Then
                Response.Write "  <input name='Anonymous' type='hidden' id='Anonymous' value='" & Anonymous & "'>" & vbCrLf           
            End If            						 			
            Response.Write "  <input name='FileType' type='hidden' id='FileType' value='" & FileType & "'>" & vbCrLf
            Response.Write "  <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>" & vbCrLf
            Response.Write "  <input name='FieldName' type='hidden' id='FieldName' value='" & FieldName & "'>" & vbCrLf			
            Response.Write "</form>" & vbCrLf
        End If
    End If
    Response.Write "</td></tr></table>" & vbCrLf
    Response.Write "</body>" & vbCrLf
    Response.Write "</html>" & vbCrLf
End Sub


Function CheckLogin()
    Dim AdminName, AdminPassword, RndPassword
    Dim UserPassword, LastPassword, UserSetting
    Dim rsUser, sqlUser
    
    AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
    AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
    RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
    
    If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Then
        CheckLogin = False
    Else
        '验证管理员帐号及密码并检测是否为多人同时使用
        sqlUser = "select * from PE_Admin where AdminName='" & AdminName & "' and Password='" & AdminPassword & "'"
        Set rsUser = Conn.Execute(sqlUser)
        If rsUser.BOF And rsUser.EOF Then
            AdminLogined = False
        Else
            If rsUser("EnableMultiLogin") <> True And Trim(rsUser("RndPassword")) <> RndPassword Then
                AdminLogined = False
            Else
                AdminLogined = True
            End If
        End If
        rsUser.Close
        Set rsUser = Nothing
    End If

    UserName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserName")))
    If AdminLogined = True Then
        uEnableUpload = True
        uMaxFileSize = 99999999
        CheckLogin = True
        Exit Function
    End If
    
    UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
    LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))
    If (UserName = "" Or UserPassword = "" Or LastPassword = "") Then
        CheckLogin = False
        Exit Function
    End If
    
    
    sqlUser = "SELECT U.UserID,U.SpecialPermission,U.UserSetting,G.GroupSetting FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
    sqlUser = sqlUser & " UserName='" & UserName & "' AND UserPassword='" & UserPassword & "' AND LastPassword='" & LastPassword & "' and IsLocked=" & PE_False & ""
    Set rsUser = Conn.Execute(sqlUser)
    If rsUser.BOF And rsUser.EOF Then
        CheckLogin = False
    Else
        CheckLogin = True
        If rsUser("SpecialPermission") = True Then
            UserSetting = Split(Trim(rsUser("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        Else
            UserSetting = Split(Trim(rsUser("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        End If
        uEnableUpload = CBool(PE_CLng(UserSetting(9)))
        uMaxFileSize = PE_CLng(UserSetting(10))
    End If
    Set rsUser = Nothing
End Function
%>
