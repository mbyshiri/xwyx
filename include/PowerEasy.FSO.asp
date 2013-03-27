<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


Function CreateMultiFolder(ByVal strPath)
    On Error Resume Next
    Dim strCreate
    If strPath = "" Or IsNull(strPath) Then CreateMultiFolder = False: Exit Function
    strPath = Replace(strPath, "\", "/")
    If Right(strPath, 1) <> "/" Then strPath = strPath & "/"
    Do While InStr(2, strPath, "/") > 1
        strCreate = strCreate & Left(strPath, InStr(2, strPath, "/") - 1)
        strPath = Mid(strPath, InStr(2, strPath, "/"))
        If Not fso.FolderExists(Server.MapPath(strCreate)) Then
            fso.CreateFolder Server.MapPath(strCreate)
        End If
        If Err Then Err.Clear: CreateMultiFolder = False: Exit Function
    Loop
    CreateMultiFolder = True
End Function

Function ReadFileContent(sFileName)
    On Error Resume Next
    Dim hf
    If Not fso.FileExists(Server.MapPath(sFileName)) Then
        ReadFileContent = ""
        Exit Function
    End If
    Set hf = fso.OpenTextFile(Server.MapPath(sFileName), 1)
    If Not hf.AtEndOfStream Then
        ReadFileContent = hf.ReadAll
    End If
    hf.Close
    Set hf = Nothing
End Function

Sub WriteToFile(WriteToFileName, WriteToFileContent)
    Dim ErrMsg
    ErrMsg = WriteToFile_FSO(WriteToFileName, WriteToFileContent)
    If ErrMsg <> "" Then
        ErrMsg = WriteToFile_ADO(WriteToFileName, WriteToFileContent)
        If ErrMsg <> "" Then
            Response.Write "<li>生成 " & WriteToFileName & " 时出错。出错原因：" & ErrMsg & "</li>"
        End If
    End If
End Sub

'=================================================
'函数名：WriteToFile
'作  用：写入相应的内容到指定的文件
'参  数：WriteToFileName ---- 写入文件的文件名
'        WriteToFileContent ---- 写入文件的内容
'=================================================
Function WriteToFile_FSO(WriteToFileName, WriteToFileContent)
    On Error Resume Next
    Err.Clear
    Dim hf
    Set hf = fso.OpenTextFile(Server.MapPath(WriteToFileName), 2, True)
    hf.Write WriteToFileContent
    hf.Close
    Set hf = Nothing
    If Err Then
        WriteToFile_FSO = Err.Description
        Err.Clear
    Else
        WriteToFile_FSO = ""
    End If
End Function

Function WriteToFile_ADO(WriteToFileName, WriteToFileContent)
    On Error Resume Next
    Err.Clear
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
	stream.Type = 2
    stream.Mode = 3
    stream.Open
    stream.Position = 0
    stream.WriteText WriteToFileContent
    stream.SaveToFile Server.MapPath(WriteToFileName), 2
    stream.Close
    Set stream = Nothing
    If Err Then
        WriteToFile_ADO = Err.Description
        Err.Clear
    Else
        WriteToFile_ADO = ""
    End If
End Function

Sub DelSerialFiles(ByVal strFiles)
    On Error Resume Next
    fso.DeleteFile strFiles
End Sub

Sub DelFiles(strUploadFiles)
    On Error Resume Next
    If Trim(strUploadFiles) = "" Or ObjInstalled_FSO <> True Then Exit Sub
    
    Dim arrUploadFiles, strFileName, i
    If InStr(strUploadFiles, "|") > 0 Then
        arrUploadFiles = Split(strUploadFiles, "|")
        For i = 0 To UBound(arrUploadFiles)
            If Trim(arrUploadFiles(i)) <> "" Then
                strFileName = InstallDir & ChannelDir & "/" & arrUploadFiles(i)
                Response.Write strFileName & "<br>"
                If fso.FileExists(Server.MapPath(strFileName)) Then
                    fso.DeleteFile (Server.MapPath(strFileName))
                End If
            End If
        Next
    Else
        strFileName = InstallDir & ChannelDir & "/" & strUploadFiles
        If fso.FileExists(Server.MapPath(strFileName)) Then
            fso.DeleteFile (Server.MapPath(strFileName))
        End If
    End If
End Sub

Sub ClearAspFile(strFilePath)
    Dim TrueDir
    Dim fs, f
    TrueDir = Server.MapPath(strFilePath)
    If fso.FolderExists(TrueDir) Then
        Set fs = fso.GetFolder(TrueDir)
        For Each f In fs.Files
            If CheckFileExt(NoAllowExt, GetFileExt(f.Name)) = True Then
                f.Delete
            End If
        Next
        Set fs = Nothing
    End If
End Sub
'取得文件路径
Function GetFilePath(FullPath)
    If FullPath <> "" Then
        GetFilePath = Trim(Left(FullPath, InStrRev(FullPath, "\")))
    Else
        GetFilePath = ""
    End If
End Function

'取得文件名
Function GetFileName(FullPath)
    If FullPath <> "" Then
        GetFileName = Trim(Mid(FullPath, InStrRev(FullPath, "\") + 1))
    Else
        GetFileName = ""
    End If
End Function

'取得文件的后缀名
Function GetFileExt(FullPath)
    Dim strFileExt
    If FullPath <> "" Then
        strFileExt = ReplaceBadChar(Trim(LCase(Mid(FullPath, InStrRev(FullPath, ".") + 1))))
        If Len(strFileExt) > 10 Then
            GetFileExt = Left(strFileExt, 3)
        Else
            GetFileExt = strFileExt
        End If
    Else
        GetFileExt = ""
    End If
End Function

Function CheckFileExt(strArr, str1)
    CheckFileExt = False
    If strArr = "" Or IsNull(strArr) Then Exit Function
    Dim arrFileExt, i
    arrFileExt = Split(strArr, "|")
    For i = 0 To UBound(arrFileExt)
        If Trim(str1) = Trim(arrFileExt(i)) Then
            CheckFileExt = True
            Exit For
        End If
    Next
End Function


%>
