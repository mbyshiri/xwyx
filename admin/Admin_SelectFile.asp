<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = False   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"


Dim FileCount, TotalSize, TotalSize_Page
Dim TotalUnit, strTotalUnit, PageUnit, strPageUnit

Dim ShowFileStyle
Dim StrFileType
Dim TruePath, theFolder, theFile, thisfile
Dim FolderCount, theSubFolder
Dim RootDir, ParentDir, CurrentDir
Dim strPath, strPath2, strPath3
Dim DialogType

Dim req, sortBy, priorSort, curFiles, currentSlot, fileItem, reverse
Dim fname, fext, fsize, ftype, fcreate, fmod, faccess
Dim kind, minmax, minmaxSlot, temp, i, mark, j
Dim theFiles, SearchKeyword
Dim rsChannel, UpFileType

ShowFileStyle = GetUploadFileStyle()
SearchKeyword = Trim(Request("SearchKeyword"))

ParentDir = Replace(Replace(Replace(Trim(Request("ParentDir")), "../", ""), "..\", ""), "\", "/")
If Left(ParentDir, 1) = "/" Then ParentDir = Right(ParentDir, Len(ParentDir) - 1)

CurrentDir = Replace(Replace(Replace(Trim(Request("CurrentDir")), "/", ""), "\", ""), "..", "")
DialogType = LCase(Trim(Request("DialogType")))
ChannelID = PE_CLng(Trim(Request("ChannelID")))
If ChannelID = 0 Then
    Response.Write "请指定频道ID！"
    Response.End
Else
    Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & ChannelID & " order by OrderID")
    If rsChannel.BOF And rsChannel.EOF Then
        Response.Write "找不到指定的频道！"
        Response.End
    Else
        If rsChannel("Disabled") = True Then
            Response.Write "此频道已经被禁用！"
            Response.End
        End If
        ChannelDir = rsChannel("ChannelDir")
        ModuleType = rsChannel("ModuleType")
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
        Select Case DialogType
            Case "pic", "batchpic", "softpic", "adpic", "productthumb"
                UpFileType = Trim(arrFileType(0))
            Case "photo", "photos"
                UpFileType = Trim(arrFileType(0)) & "|" & Trim(arrFileType(1))
            Case "flash"
                UpFileType = Trim(arrFileType(1))
            Case "media"
                UpFileType = Trim(arrFileType(2))
            Case "rm"
                UpFileType = Trim(arrFileType(3))
            Case "fujian"
                UpFileType = Trim(arrFileType(4))
            Case "soft"
                UpFileType = Trim(arrFileType(1)) & "|" & Trim(arrFileType(2)) & "|" & Trim(arrFileType(3)) & "|" & Trim(arrFileType(4))
            Case "all"
                UpFileType = Trim(arrFileType(0)) & "|" & Trim(arrFileType(1)) & "|" & Trim(arrFileType(2)) & "|" & Trim(arrFileType(3)) & "|" & Trim(arrFileType(4))
            Case Else
                UpFileType = ""
        End Select
        If DialogType = "softpic" Then
            UploadDir = "UploadSoftPic"
        Else
            UploadDir = rsChannel("UploadDir")
        End If
    End If
    rsChannel.Close
    Set rsChannel = Nothing
End If
If ChannelDir = "" Then
    Response.Write "未指定相应的目录！"
    Response.End
End If
strFileName = "Admin_SelectFile.asp?ChannelID=" & ChannelID & "&DialogType=" & DialogType

RootDir = InstallDir & ChannelDir & "/" & UploadDir
strPath = RootDir
strPath2 = UploadDir
strPath3 = ""
If ParentDir <> "" Then
    strPath = strPath & "/" & ParentDir
    strPath2 = strPath2 & "/" & ParentDir
    strPath3 = ParentDir
End If
If CurrentDir <> "" Then
    strPath = strPath & "/" & CurrentDir
    strPath2 = strPath2 & "/" & CurrentDir
    If ParentDir <> "" Then
        strPath3 = strPath3 & "/" & CurrentDir & "/"
    Else
        strPath3 = CurrentDir & "/"
    End If
End If
strPath = Replace(strPath, "//", "/")
strPath2 = Replace(strPath2, "//", "/")
TruePath = Server.MapPath(strPath)

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>从已上传文件选择</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
If ObjInstalled_FSO = False Then
    Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    Response.Write "</body></html>"
    Response.End
End If

If SearchKeyword <> "" Then
    Response.Write "    <form name='myform' method='Post' action='" & strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & "'>" & vbCrLf
Else
    Response.Write "    <form name='myform' method='Post' action='" & strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "'>" & vbCrLf
End If
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>" & vbCrLf

Response.Write "  <tr class='tdbg'> "
Response.Write "    <td width='0' height='30'>"
Response.Write "    &nbsp;&nbsp;<img src='Images/admin_open.gif'  width='13' height='13' border='0'>&nbsp;&nbsp;&nbsp;"
If ShowFileStyle = 1 Then
    Response.Write "<a href='Admin_UploadFile_Style.asp?ShowFileStyle=2'>切换到缩略图方式 </a>" & vbCrLf
Else
    Response.Write "<a href='Admin_UploadFile_Style.asp?ShowFileStyle=1'>切换到详细信息方式</a>" & vbCrLf
End If

Response.Write "    </td>"

Response.Write "<td height='30'>"
Response.Write "    <table width='100%' border='0' align='center' cellpadding='1' cellspacing='1'><tr><td height='22' align='right'>"
Response.Write "&nbsp; 搜索当前目录文件：</td><td height='22'><input type='text' name='SearchKeyword' id='SearchKeyword' size='18' value=''>&nbsp;</td><td height='22'><input type='submit' name='submit1' value=' 搜索 '>"
Response.Write "    </td></tr></table></td>"
Response.Write "  </tr>"
Response.Write "</table>" & vbCrLf

If fso.FolderExists(TruePath) = False Then
    Response.Write "找不到文件夹！可能是配置有误！"
    Response.End
End If

Dim Add2Array
If fso.FolderExists(TruePath) = False Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>找不到文件夹！请上传文件后再进行管理！</li>"
    Response.End
End If
Response.Write "<Script Language=""JavaScript"">" & vbCrLf
Response.Write "function reSort(which)" & vbCrLf
Response.Write "{" & vbCrLf
If SearchKeyword <> "" Then
    Response.Write "document.myform.SearchKeyword.value = '" & SearchKeyword & "';" & vbCrLf
End If
Response.Write "document.myform.sortby.value = which;" & vbCrLf
Response.Write "document.myform.submit();" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</Script>" & vbCrLf
'Dim FolderCount

req = Trim(Request("sortBy"))
If Len(req) < 1 Or req = "-1" Then
    sortBy = 0
Else
    sortBy = CInt(req)
End If
req = Request("priorSort")
If Len(req) < 1 Or req = "-1" Then
    priorSort = -1
Else
    priorSort = CInt(req)
End If
'设置倒序
If sortBy = priorSort Then
    reverse = True
    priorSort = -1
Else
    reverse = False
    priorSort = sortBy
End If

Set theFolder = fso.GetFolder(TruePath)
Set curFiles = theFolder.Files

ReDim theFiles(500)
currentSlot = -1

For Each fileItem In curFiles
    Add2Array = False
    fname = fileItem.name
    If SearchKeyword <> "" Then
        If InStr(LCase(fname), LCase(SearchKeyword)) > 0 Then
            Add2Array = True
        End If
    Else
        Add2Array = True
    End If
    If Add2Array = True Then
        fext = InStrRev(fname, ".")
        If fext < 1 Then fext = "" Else fext = Mid(fname, fext + 1)
        ftype = fileItem.Type
        fsize = fileItem.size
        fcreate = fileItem.DateCreated
        fmod = fileItem.DateLastModified
        faccess = fileItem.DateLastAccessed
        currentSlot = currentSlot + 1
        If currentSlot > UBound(theFiles) Then
            ReDim Preserve theFiles(currentSlot + 99)
        End If

        theFiles(currentSlot) = Array(fname, fext, fsize, ftype, fcreate, fmod, faccess)
    End If
Next

If currentSlot > -1 Then
    FileCount = currentSlot ' 文件数量
    ReDim Preserve theFiles(currentSlot)


    If VarType(theFiles(0)(sortBy)) = 8 Then
        If reverse Then kind = 1 Else kind = 2
    Else
        If reverse Then kind = 3 Else kind = 4
    End If
    For i = FileCount To 0 Step -1
        minmax = theFiles(0)(sortBy)
        minmaxSlot = 0
        For j = 1 To i
            Select Case kind
                Case 1
                mark = (StrComp(theFiles(j)(sortBy), minmax, vbTextCompare) < 0)
                Case 2
                mark = (StrComp(theFiles(j)(sortBy), minmax, vbTextCompare) > 0)
                Case 3
                mark = (theFiles(j)(sortBy) < minmax)
                Case 4
                mark = (theFiles(j)(sortBy) > minmax)
            End Select
            If mark Then
                minmax = theFiles(j)(sortBy)
                minmaxSlot = j
            End If
        Next
        If minmaxSlot <> i Then
            temp = theFiles(minmaxSlot)
            theFiles(minmaxSlot) = theFiles(i)
            theFiles(i) = temp
        End If
    Next
Else
    FileCount = 0
End If


If ShowFileStyle = 1 Then
    Call ShowFileDetail
Else
    Call ShowFileThumb
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
If currentSlot > -1 Then
    Response.Write "<Script Language=""JavaScript"">" & vbCrLf
    Response.Write "var Sort=document.getElementById(""Sort" & sortBy & """);" & vbCrLf
    If reverse Then
        Response.Write "    Sort.src=""Images/Calendar_Down.gif"";" & vbCrLf
    Else
        Response.Write "    Sort.src=""Images/Calendar_Up.gif"";" & vbCrLf
    End If
    Response.Write "    Sort.style.display="""";    " & vbCrLf
    Response.Write "</Script>" & vbCrLf
End If


Sub ShowFileThumb()
    Response.Write "<br><table width='100%' cellpadding='2' cellspacing='1' class='border'><tr class='title' height='22'><td colspan='20'><b>子目录导航：</b></td></tr><tr class='tdbg'>"

    For Each theSubFolder In theFolder.SubFolders
        If ParentDir <> "" Then
            Response.Write "<td><a href='" & strFileName & "&ParentDir=" & ParentDir & "/" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        Else
            Response.Write "<td><a href='" & strFileName & "&ParentDir=" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        End If
        FolderCount = FolderCount + 1
        If FolderCount Mod 10 = 0 Then Response.Write "</td><tr class='tdbg'>"
    Next
    Response.Write "</tr></table>" & vbCrLf

    Response.Write "<br><table width='100%'><tr><td>当前目录：" & RootDir
    If ParentDir <> "" Then
        Response.Write "/" & ParentDir
    End If
    If CurrentDir <> "" Then
        Response.Write "/" & CurrentDir & "</td><td align='right'>"
        If ParentDir <> "" Then
            If InStrRev(ParentDir, "/") > 0 Then
                Response.Write "<a href='" & strFileName & "&ParentDir=" & Left(ParentDir, InStrRev(ParentDir, "/") - 1)
                Response.Write "&CurrentDir=" & Mid(ParentDir, InStrRev(ParentDir, "/") + 1)
            Else
                Response.Write "<a href='" & strFileName & "&ParentDir=&CurrentDir=" & ParentDir
            End If
        Else
            Response.Write "<a href='" & strFileName
        End If
        Response.Write "'>↑返回上级目录</a>"
    End If
    Response.Write "</td></tr></table>" & vbCrLf

    If SearchKeyword <> "" Then
        Response.Write "<br>&gt;&gt;&nbsp;当前目录文件名中含有的 <font color='red'>" & SearchKeyword & "</font> 文件"
    End If
    Response.Write "    <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    If currentSlot > -1 Then
        Response.Write "    <td height='18'> 排序方式：&nbsp;&nbsp;<a href=""javascript:reSort(0);"">文件名&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort0'></a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(2);"">大小&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort2'></a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(3);"">类型&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort3'></a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(5);"">上次修改时间&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort5'></a></td>" & vbCrLf
    Else
        Response.Write "    <td height='18'></td>" & vbCrLf
    End If
    Response.Write "    <td align='right'>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    
    If currentSlot = -1 Then
        Response.Write "    <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
        Response.Write "<tr class='tdbg'><td align='center' colspan='2'><br><br>当前目录下没有任何文件！<br><br></td>"
        Response.Write " </tr>"
        Response.Write "</table>" & vbCrLf
    Else
        strFileName = strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir


        TotalSize = 0
        'TruePath = Server.MapPath(strPath)
        'Set theFolder = fso.GetFolder(TruePath)
        TotalUnit = 1
        For Each theFile In theFolder.Files
            StrFileType = LCase(Mid(theFile.name, InStrRev(theFile.name, ".") + 1))
            If FoundInArr(UpFileType, StrFileType, "|") = True Then

                If TotalUnit = 1 Then
                    TotalSize = TotalSize + theFile.size / 1024
                    strTotalUnit = "KB"
                ElseIf TotalUnit = 2 Then
                    TotalSize = TotalSize + theFile.size / 1024 / 1024
                    strTotalUnit = "MB"
                ElseIf TotalUnit = 3 Then
                    TotalSize = TotalSize + theFile.size / 1024 / 1024 / 1024
                    strTotalUnit = "GB"
                End If
                If TotalSize > 1024 Then
                    TotalSize = TotalSize / 1024
                    TotalUnit = TotalUnit + 1
                End If
            End If
        Next
        TotalSize = Round(TotalSize, 2)

        totalPut = FileCount
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
        Dim c
        Dim FileNum
        FileNum = 0
        TotalSize_Page = 0
        PageUnit = 1
        
        'If totalPut > 0 Then
            'Response.Write "<br>"
            'Call showpage2(strFileName, totalPut, MaxPerPage, True)
            'Response.Write "<br>"
        'End If

        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "     <td>" & vbCrLf
        Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='3' class='border'>" & vbCrLf
        Response.Write "  <tr class='tdbg'>" & vbCrLf

        For i = 0 To FileCount
            StrFileType = LCase(theFiles(i)(1))
            If FoundInArr(UpFileType, StrFileType, "|") = True Then
                c = c + 1
            End If
            If FileNum >= MaxPerPage Then
                Exit For
            ElseIf c > MaxPerPage * (CurrentPage - 1) Then
                If FoundInArr(UpFileType, StrFileType, "|") = True Then
                    Response.Write "   <td>" & vbCrLf
                    Response.Write "      <table width='100%' height='100%' border='0' cellpadding='0' cellspacing='2'>" & vbCrLf
                    Response.Write "        <tr>" & vbCrLf
                    Response.Write "          <td colspan='2' align='center'>" & vbCrLf
                    If DialogType = "soft" Or DialogType = "photo" Or DialogType = "productthumb" Then
                        Response.Write "<a href='#' onClick=""window.returnValue='" & strPath3 & theFiles(i)(0) & "|" & Round(theFiles(i)(2) / 1024) & "';window.close();"">"
                    Else
                        If ModuleType = 1 Or ModuleType = 5 Or ModuleType = 6 Or ModuleType = 7 Then
                            Response.Write "<a href='#' onClick=""window.returnValue='" & strPath & "/" & theFiles(i)(0) & "|" & Round(theFiles(i)(2) / 1024) & "';window.close();"">"
                        Else
                            Response.Write "<a href='#' onClick=""window.returnValue='" & strPath2 & "/" & theFiles(i)(0) & "|" & Round(theFiles(i)(2) / 1024) & "';window.close();"">"
                        End If
                    End If
                    Select Case StrFileType
                    Case "gif", "jpg", "jpeg", "jpe", "bmp", "png"
                        Response.Write "<img src='" & strPath & "/" & theFiles(i)(0) & "' width='140' height='100' border='0' title='点此图片将返回，点下面的文件名将查看原始文件！'></a>"
                    Case "swf"
                        Response.Write "<img src='images/filetype_flash.gif' width='140' height='100' border='0'></a>"
                    Case "wmv", "avi", "asf", "mpg"
                        Response.Write "<img src='images/filetype_media.gif' width='140' height='100' border='0'></a>"
                    Case "rm", "ra", "ram"
                        Response.Write "<img src='images/filetype_rm.gif' width='140' height='100' border='0'></a>"
                    Case "rar"
                        Response.Write "<img src='images/filetype_rar.gif' width='140' height='100' border='0'></a>"
                    Case "zip"
                        Response.Write "<img src='images/filetype_zip.gif' width='140' height='100' border='0'></a>"
                    Case "exe"
                        Response.Write "<img src='images/filetype_exe.gif' width='140' height='100' border='0'></a>"
                    Case Else
                        Response.Write "<img src='images/filetype_other.gif' width='140' height='100' border='0'></a>"
                    End Select

                    Response.Write "          </td></tr>" & vbCrLf
                    Response.Write "        <tr>" & vbCrLf
                    Response.Write "          <td align='right'>文 件 名：</td>" & vbCrLf
                    Response.Write "          <td><a href='" & strPath & "/" & theFiles(i)(0) & "' target='_blank'>" & theFiles(i)(0) & "</a></td>" & vbCrLf
                    Response.Write "        </tr>" & vbCrLf
                    Response.Write "        <tr>" & vbCrLf
                    Response.Write "          <td align='right'>文件大小：</td>" & vbCrLf
                    Response.Write "          <td>" & Round(theFiles(i)(2) / 1024) & " KB</td>" & vbCrLf
                    Response.Write "        </tr>" & vbCrLf
                    Response.Write "        <tr>" & vbCrLf
                    Response.Write "          <td align='right'>文件类型：</td>" & vbCrLf
                    Response.Write "          <td>" & theFiles(i)(3) & "</td>" & vbCrLf
                    Response.Write "        </tr>" & vbCrLf
                    Response.Write "        <tr>" & vbCrLf
                    Response.Write "          <td align='right'>修改时间：</td>" & vbCrLf
                    Response.Write "          <td>" & theFiles(i)(5) & "</td>" & vbCrLf
                    Response.Write "        </tr>" & vbCrLf
                    Response.Write "      </table></td>" & vbCrLf
                
                    FileNum = FileNum + 1
                    If FileNum Mod 4 = 0 Then Response.Write "</td><tr class='tdbg'>"
                    If PageUnit = 1 Then
                        TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024
                        strPageUnit = "KB"
                    ElseIf PageUnit = 2 Then
                        TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024 / 1024
                        strPageUnit = "MB"
                    ElseIf PageUnit = 3 Then
                        TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024 / 1024 / 1024
                        strPageUnit = "GB"
                    End If
                    If TotalSize_Page > 1024 Then
                        TotalSize_Page = TotalSize_Page / 1024
                        PageUnit = PageUnit + 1
                    End If
                End If
            End If
        Next
        TotalSize_Page = Round(TotalSize_Page, 2)

        Response.Write "            </tr>" & vbCrLf
        Response.Write "        </table> " & vbCrLf
        Response.Write "     </td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
                
        Response.Write "<input type='hidden' name='priorsort' value='" & priorSort & "'>" & vbCrLf
        Response.Write "<input type='hidden' name='sortby' value='-1'>" & vbCrLf
        Response.Write "    </form>" & vbCrLf
        showpage2 strFileName, totalPut, MaxPerPage, True
        Response.Write "<br><div align='center'>本页共显示 <b>" & FileNum & "</b> 个文件，占用 <b>" & TotalSize_Page & "</b> " & strPageUnit & "</div>"
        Response.Write "</body></html>"
    End If
End Sub


Sub ShowFileDetail()
    If SearchKeyword <> "" Then
        Response.Write "<br>&gt;&gt;&nbsp;当前目录文件名中含有的 <font color='red'>" & SearchKeyword & "</font> 文件"
    Else
        Response.Write "<br>"
    End If
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0' ><tr><td>当前目录：" & RootDir
    If ParentDir <> "" Then
        Response.Write "/" & ParentDir
    End If
    If CurrentDir <> "" Then
        Response.Write "/" & CurrentDir & "</td><td align='right'>"
        If ParentDir <> "" Then
            If InStrRev(ParentDir, "/") > 0 Then
                Response.Write "<a href='" & strFileName & "&ParentDir=" & Left(ParentDir, InStrRev(ParentDir, "/") - 1)
                Response.Write "&CurrentDir=" & Mid(ParentDir, InStrRev(ParentDir, "/") + 1)
            Else
                Response.Write "<a href='" & strFileName & "&ParentDir=&CurrentDir=" & ParentDir
            End If
        Else
            Response.Write "<a href='" & strFileName
        End If
        Response.Write "'>↑返回上级目录</a>"
    End If
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "    <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf

    Response.Write "    <tr height='18'>" & vbCrLf
    Response.Write "    <td class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'"">&nbsp;&nbsp;<a href=""javascript:reSort(0);"">文件名&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort0'></a></td>" & vbCrLf
    Response.Write "    <td width='80' align=""right"" class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'""><a href=""javascript:reSort(2);"">大小&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort2'></a>&nbsp;</td>" & vbCrLf
    Response.Write "    <td width='180' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'"">&nbsp;<a href=""javascript:reSort(3);"">类型&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort3'></a></td>" & vbCrLf
    Response.Write "    <td width='140' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'""><a href=""javascript:reSort(5);"">上次修改时间&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort5'></a></td>" & vbCrLf

    Response.Write "    </tr>" & vbCrLf
    If sortBy <> 0 Then
        Call ShowFileDetail_Fil
        If SearchKeyword = "" Then
            Response.Write "<tr height=1><td></td></tr>" & vbCrLf
            Call ShowFileDetail_fol
        End If
    Else
        If SearchKeyword = "" Then
            Call ShowFileDetail_fol
            Response.Write "<tr height=1><td></td></tr>" & vbCrLf
        End If
        Call ShowFileDetail_Fil
    End If
    Response.Write "    </table>" & vbCrLf
    Response.Write "<input type='hidden' name='priorsort' value='" & priorSort & "'>" & vbCrLf
    Response.Write "<input type='hidden' name='sortby' value='-1'>" & vbCrLf
    If currentSlot > -1 Then
        Call ShowJS_Tooltip
    End If
    Response.Write "    </form>" & vbCrLf

End Sub


Sub ShowFileDetail_Fil()
    If currentSlot > -1 Then

        For i = 0 To FileCount
            Response.Write "<tr onmouseout=""this.className='tdbgmouseout1'"" onmouseover=""this.className='tdbg1'"">" & vbCrLf

            Response.Write "          <td align='left'>" & vbCrLf

            Select Case LCase(theFiles(i)(1))
            Case "jpeg", "jpe", "bmp", "png"
                Response.Write "<img src='images/Folder/img.gif'>"
            Case "swf"
                Response.Write "<img src='images/Folder/Ftype_flash.gif'>"
            Case "dll", "vbp"
                Response.Write "<img src='images/Folder/sys.gif'>"
            Case "wmv", "avi", "asf", "mpg"
                Response.Write "<img src='images/Folder/Ftype_media.gif'>"
            Case "rm", "ra", "ram"
                Response.Write "<img src='images/Folder/Ftype_rm.gif'>"
            Case "rar", "zip"
                Response.Write "<img src='images/Folder/zip.gif'>"
            Case "xml", "txt", "exe", "doc", "html", "htm", "jpg", "gif", "xls", "asp"
                Response.Write "<img src='images/Folder/" & theFiles(i)(1) & ".gif'>"
            Case Else
                Response.Write "<img src='images/Folder/other.gif'>"
            End Select
            
            If DialogType = "soft" Or DialogType = "photo" Or DialogType = "productthumb" Then
                Response.Write "<a href='#' onClick=""window.returnValue='" & strPath3 & theFiles(i)(0) & "|" & Round(theFiles(i)(2) / 1024) & "';window.close();"">"
            Else
                If ModuleType = 1 Or ModuleType = 5 Or ModuleType = 6 Or ModuleType = 7 Then
                    Response.Write "<a href='#' onClick=""window.returnValue='" & strPath & "/" & theFiles(i)(0) & "|" & Round(theFiles(i)(2) / 1024) & "';window.close();"">"
                Else
                    Response.Write "<a href='#' onClick=""window.returnValue='" & strPath2 & "/" & theFiles(i)(0) & "|" & Round(theFiles(i)(2) / 1024) & "';window.close();"">"
                End If
            End If
            Response.Write "<span onmouseover=""ShowADPreview('" & FixJs(GetFileContent(strPath & "/" & theFiles(i)(0), theFiles(i)(1))) & "')"" onmouseout=""hideTooltip('dHTMLADPreview')"">" & vbCrLf
            Response.Write theFiles(i)(0) & "</span></a></td>" & vbCrLf
            Response.Write " <td width='80' align='right'>" & FormatNumber(theFiles(i)(2) / 1024, 0, vbTrue, vbFalse, vbTrue) & " KB</td>" & vbCrLf
            Response.Write " <td width='180'>&nbsp;" & theFiles(i)(3) & "</td>" & vbCrLf
            Response.Write " <td width='140'>" & theFiles(i)(5) & "</td>" & vbCrLf
            Response.Write "</tr>" & vbCrLf
        Next
    End If
End Sub

Function GetFileContent(sPath, sType)
    If IsNull(sPath) Or sPath = "" Then
        GetFileContent = "&nbsp;此文件非图片或动画，无预览&nbsp;"
        Exit Function
    End If
    If IsNull(sType) Or sType = "" Then
        GetFileContent = "&nbsp;此文件非图片或动画，无预览&nbsp;"
        Exit Function
    End If

    If Not fso.FileExists(Server.MapPath(sPath)) Then
        GetFileContent = "&nbsp;此文件不存在&nbsp;"
        Exit Function
    End If
    Dim strFile

    Select Case LCase(sType)
    Case "jpeg", "jpe", "bmp", "png", "jpg", "gif"
        strFile = "<img src='" & sPath & "'"
        strFile = strFile & " width='200'"
        strFile = strFile & " height='120'"
        strFile = strFile & " border='0'>"
    Case "wmv", "avi", "asf", "mpg", "rm", "ra", "ram", "swf"
        strFile = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0'"
        strFile = strFile & " width='200'"
        strFile = strFile & " height='120'"
        strFile = strFile & "><param name='movie' value='" & sPath & "'>"
        strFile = strFile & "<param name='wmode' value='transparent'>"
        strFile = strFile & "<param name='quality' value='autohigh'>"
        strFile = strFile & "<embed src='" & sPath & "' quality='autohigh'"
        strFile = strFile & " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
        strFile = strFile & " wmode='transparent'"
        strFile = strFile & " width='200'"
        strFile = strFile & " height='120'"
        strFile = strFile & "></embed></object>"
    Case Else
        strFile = "&nbsp;此文件非图片或动画，无预览&nbsp;"
    End Select

    GetFileContent = strFile
End Function

Sub ShowFileDetail_fol()
    Dim strHtml
    Response.Write "<tr>"
    strHtml = ""
    For Each theSubFolder In theFolder.SubFolders
        If ParentDir <> "" Then
            strHtml = strHtml & "<td height='18'>&nbsp;&nbsp;<img src='Images/Folder/folderclosed.gif'><a href='" & strFileName & "&ParentDir=" & ParentDir & "/" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        Else
            strHtml = strHtml & "<td height='18'>&nbsp;&nbsp;<img src='Images/Folder/folderclosed.gif'><a href='" & strFileName & "&ParentDir=" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        End If
        strHtml = strHtml & "<td width='50' align=""right"">&nbsp;</td>"
        strHtml = strHtml & "<td width='180'>&nbsp;文件夹</td>"
        strHtml = strHtml & "<td width='140'>" & theSubFolder.DateLastModified & "</td>"
        strHtml = strHtml & "</tr><tr>"
    Next
    Response.Write strHtml
End Sub

Function FixJs(Str)
    If Str <> "" Then
        Str = Replace(Str, "&#39;", "'")
        Str = Replace(Str, "\", "\\")
        Str = Replace(Str, Chr(34), "\""")
        Str = Replace(Str, Chr(39), "\'")
        Str = Replace(Str, Chr(13), "\n")
        Str = Replace(Str, Chr(10), "\r")
        Str = Replace(Str, "'", "&#39;")
        Str = Replace(Str, """", "&quot;")
    End If
    FixJs = Str
End Function



Function ShowJS_Tooltip()
    Response.Write "<div id=dHTMLADPreview style='Z-INDEX: 1000; LEFT: 0px; VISIBILITY: hidden; WIDTH: 10px; POSITION: absolute; TOP: 0px; HEIGHT: 10px'></DIV>"
    Response.Write "<SCRIPT language = 'JavaScript'>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "var tipTimer;" & vbCrLf
    Response.Write "function locateObject(n, d)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "   var p,i,x;" & vbCrLf
    Response.Write "   if (!d) d=document;" & vbCrLf
    Response.Write "   if ((p=n.indexOf('?')) > 0 && parent.frames.length)" & vbCrLf
    Response.Write "   {d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}" & vbCrLf
    Response.Write "   if (!(x=d[n])&&d.all) x=d.all[n]; " & vbCrLf
    Response.Write "   for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];" & vbCrLf
    Response.Write "   for (i=0;!x&&d.layers&&i<d.layers.length;i++) x=locateObject(n,d.layers[i].document); return x;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ShowADPreview(ADContent)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  showTooltip('dHTMLADPreview',event, ADContent, '#ffffff','#000000','#000000','6000')" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function showTooltip(object, e, tipContent, backcolor, bordercolor, textcolor, displaytime)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "   window.clearTimeout(tipTimer)" & vbCrLf
    Response.Write "   if (document.all) {" & vbCrLf
    Response.Write "       locateObject(object).style.top=document.body.scrollTop+event.clientY+20" & vbCrLf
    Response.Write "       locateObject(object).innerHTML='<table style=""font-family:宋体; font-size: 9pt; border: '+bordercolor+'; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; background-color: '+backcolor+'"" width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td nowrap><font style=""font-family:宋体; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table> '" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clientWidth) > (document.body.clientWidth + document.body.scrollLeft)) {" & vbCrLf
    Response.Write "           locateObject(object).style.left = (document.body.clientWidth + document.body.scrollLeft) - locateObject(object).clientWidth-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).style.left=document.body.scrollLeft+event.clientX" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).style.visibility='visible';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else if (document.layers) {" & vbCrLf
    Response.Write "       locateObject(object).document.write('<table width=""10"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr bgcolor=""'+bordercolor+'""><td><table width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr bgcolor=""'+backcolor+'""><td nowrap><font style=""font-family:宋体; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table></td></tr></table>')" & vbCrLf
    Response.Write "       locateObject(object).document.close()" & vbCrLf
    Response.Write "       locateObject(object).top=e.y+20" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clip.width) > (window.pageXOffset + window.innerWidth)) {" & vbCrLf
    Response.Write "           locateObject(object).left = window.innerWidth - locateObject(object).clip.width-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).left=e.x;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).visibility='show';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else {" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function hideTooltip(object) {" & vbCrLf
    Response.Write "    if (document.all) {" & vbCrLf
    Response.Write "        locateObject(object).style.visibility = 'hidden';" & vbCrLf
    Response.Write "        locateObject(object).style.left = 1;" & vbCrLf
    Response.Write "        locateObject(object).style.top = 1;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    } else {" & vbCrLf
    Response.Write "        if (document.layers) {" & vbCrLf
    Response.Write "            locateObject(object).visibility = 'hide';" & vbCrLf
    Response.Write "            locateObject(object).left = 1;" & vbCrLf
    Response.Write "            locateObject(object).top = 1;" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        } else {" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Function

Function GetUploadFileStyle()
    ShowFileStyle = Request.Cookies("ShowFileStyle")
    If ShowFileStyle = "" Or Not IsNumeric(ShowFileStyle) Then
        ShowFileStyle = 1
    Else
        ShowFileStyle = Int(ShowFileStyle)
    End If
    GetUploadFileStyle = ShowFileStyle
End Function

Sub showpage2(sfilename, totalnumber, MaxPerPage, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        Exit Sub
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
        
    strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
    strTemp = strTemp & "共 <b>" & totalnumber & "</b> 个文件，占用 <b>" & TotalSize & "</b> " & strTotalUnit & "&nbsp;&nbsp;&nbsp;"
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage < 2 Then
        strTemp = strTemp & "首页 上一页&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a>&nbsp;"
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "下一页 尾页"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>尾页</a>"
    End If
    strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>页 "
    If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;<input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) submit();"">" & "个文件/页"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & "个文件/页"
    End If
    strTemp = strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"
    For i = 1 To TotalPage
        strTemp = strTemp & "<option value='" & i & "'"
        If PE_CLng(CurrentPage) = PE_CLng(i) Then strTemp = strTemp & " selected "
        strTemp = strTemp & ">第" & i & "页</option>"
    Next
    strTemp = strTemp & "</select>"
    strTemp = strTemp & "</td></tr></form></table>"
    Response.Write strTemp
End Sub
%>
