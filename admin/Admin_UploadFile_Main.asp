<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.CreateThumb.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 1   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim TruePath, theFolder, theSubFolder, theFile, thisfile, FileCount, TotalSize, TotalSize_Page
Dim TotalUnit, strTotalUnit, PageUnit, strPageUnit
Dim StrFileType, strFiles
Dim strDirName, tUploadDir, ShowFileStyle
Dim RootDir, ParentDir, CurrentDir
Dim strPath, strPath2, strPath3


'��ȡƵ���������
tUploadDir = Trim(Request("UploadDir"))
If ChannelID > 0 Then

Else
    If tUploadDir = "UploadAdPic" Then
        ChannelName = "��վ���"
        UploadDir = "UploadAdPic"
        ChannelDir = ADDir
    End If
End If

'������Ա����Ȩ��
If AdminPurview > 1 Then
    If ChannelID > 0 Then
        If AdminPurview_Channel = "" Then
            AdminPurview_Channel = 5
        Else
            AdminPurview_Channel = PE_CLng(AdminPurview_Channel)
        End If
        If AdminPurview_Channel > 1 Then
            PurviewPassed = False
        Else
            PurviewPassed = True
        End If
    Else
        If tUploadDir = "UploadAdPic" Then
            PurviewPassed = CheckPurview_Other(AdminPurview_Others, "AD")
        Else
            PurviewPassed = False
        End If
    End If
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
        Call WriteEntry(6, AdminName, "ԽȨ����")
        Response.End
    End If
End If

Dim req, sortBy, priorSort, curFiles, currentSlot, fileItem, reverse
Dim fname, fext, fsize, ftype, fcreate, fmod, faccess
Dim kind, minmax, minmaxSlot, temp, i, mark, j
Dim theFiles, SearchKeyword

'��ȡ�鿴��ʽ
ShowFileStyle = GetUploadFileStyle()

SearchKeyword = Trim(Request("SearchKeyword"))

ParentDir = Replace(Replace(Replace(Replace(Trim(Request("ParentDir")), "../", ""), "..\", ""), "\", "/"),"|","/")
If Left(ParentDir, 1) = "/" Then ParentDir = Right(ParentDir, Len(ParentDir) - 1)
CurrentDir = Replace(Replace(Replace(Trim(Request("CurrentDir")), "/", ""), "\", ""), "..", "")

Dim rs, sql
Select Case ModuleType
Case 1
    strDirName = ChannelName & "���ϴ��ļ�"
    sql = "select UploadFiles from PE_Article where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case 2
    If tUploadDir = "UploadSoftPic" Then
        UploadDir = "UploadSoftPic"
        strDirName = ChannelName & "�����ͼƬ"
        sql = "select SoftPicUrl from PE_Soft where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
    Else
        strDirName = ChannelName & "���ϴ����"
        sql = "select DownloadUrl from PE_Soft where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "$$$" & rs(0)
            End If
            rs.MoveNext
        Loop
    End If
Case 3
    strDirName = ChannelName & "���ϴ�ͼƬ"
    sql = "select PhotoThumb,PhotoUrl from PE_Photo"
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "$$$" & rs(0)
        End If
        If rs(1) <> "" Then
            strFiles = strFiles & "$$$" & rs(1)
        End If
        rs.MoveNext
    Loop
Case 5
    strDirName = ChannelName & "���ϴ�ͼƬ"
    sql = "select UploadFiles from PE_Product where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
'�������ģ���ͼƬ
'������
'2006-1-14
Case 6
    strDirName = ChannelName & "���ϴ�ͼƬ"
    sql = "select SupplyPicUrl from PE_Supply where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case 7 '�������ģ���ͼƬ
    Dim HouseTable
    strDirName = ChannelName & "���ϴ�ͼƬ"
    For i = 1 To 5
        Select Case i
        Case 1
            HouseTable = "PE_HouseCS"
        Case 2
            HouseTable = "PE_HouseCZ"
        Case 3
            HouseTable = "PE_HouseQG"
        Case 4
            HouseTable = "PE_HouseQZ"
        Case 5
            HouseTable = "PE_HouseHZ"
        End Select
        sql = "select UploadPhotos from " & HouseTable & " where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
    Next
Case 8 '����˲���Ƹģ���ͼƬ
    strDirName = ChannelName & "���ϴ�ͼƬ"
    sql = "select Photo from PE_Resume"
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0) <> "" Then
            strFiles = strFiles & "|" & rs(0)
        End If
        rs.MoveNext
    Loop
Case Else
    If tUploadDir = "UploadAdPic" Then
        strDirName = "�ϴ��Ĺ��ͼƬ"
        sql = "select ImgUrl from PE_Advertisement"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
    End If
End Select
rs.Close
Set rs = Nothing
strFiles = LCase(strFiles)

RootDir = InstallDir & ChannelDir & "/" & UploadDir

strPath = RootDir
strPath2 = UploadDir
strPath3 = ""
If ParentDir <> "" Then
    If InStr(ParentDir, ChannelDir & "/" & UploadDir) > 0 Then
        ParentDir = Replace(ParentDir, ChannelDir & "/" & UploadDir & "/", "")
        ParentDir = Replace(ParentDir, ChannelDir & "/" & UploadDir, "")
    End If
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

strFileName = "Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir

Response.Write "<html><head><title>�ϴ��ļ�����</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>"
Response.Write "<SCRIPT language='javascript'>" & vbCrLf
Response.Write "function unselectall(){" & vbCrLf
Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function CheckAll(form){" & vbCrLf
Response.Write "  for (var i=0;i<form.elements.length;i++)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "    var e = form.elements[i];" & vbCrLf
Response.Write "    if (e.Name != 'chkAll')" & vbCrLf
Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "function preloadImg(src) {" & vbCrLf
Response.Write "  var img=new Image();" & vbCrLf
Response.Write "  img.src=src" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "preloadImg('Images/admin_upload_open.gif');" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "var displayBar=false;" & vbCrLf
Response.Write "function switchBar(obj) {" & vbCrLf
Response.Write "  if (displayBar) {" & vbCrLf
Response.Write "    parent.frame.cols='0,*';" & vbCrLf
Response.Write "    displayBar=false;" & vbCrLf
Response.Write "    obj.src='Images/admin_upload_open.gif';" & vbCrLf
Response.Write "    obj.title='������ļ�Ŀ¼���͵���';" & vbCrLf
Response.Write "  } else {" & vbCrLf
Response.Write "    parent.frame.cols='160,*';" & vbCrLf
Response.Write "    displayBar=true;" & vbCrLf
Response.Write "    obj.src='Images/admin_upload_close.gif';" & vbCrLf
Response.Write "    obj.title='�ر�����ļ�Ŀ¼���͵���';" & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
If SearchKeyword <> "" Then
    Response.Write "    <form name='myform' method='Post' action='" & strFileName & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & "'>" & vbCrLf
Else
    Response.Write "    <form name='myform' method='Post' action='" & strFileName & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir & "'>" & vbCrLf
End If
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>" & vbCrLf

Response.Write "  <tr class='tdbg'> "
Response.Write "    <td width='0' height='30'>"
Response.Write "    <table width='100%' border='0' align='center' cellpadding='1' cellspacing='1'><tr><td height='22'>"
Response.Write "      &nbsp;<img onclick='switchBar(this)' src='Images/admin_upload_open.gif' title='������ļ�Ŀ¼���͵���' style='cursor:hand'></td><td height='22'>" & vbCrLf
Response.Write "</td><td>"
If ShowFileStyle = 1 Then
    Response.Write "<a href='Admin_UploadFile_Style.asp?ShowFileStyle=2'>�л�������ͼ��ʽ </a>" & vbCrLf
Else
    Response.Write "<a href='Admin_UploadFile_Style.asp?ShowFileStyle=1'>�л�����ϸ��Ϣ��ʽ</a>" & vbCrLf
End If
Response.Write "    </td></tr></table>"
Response.Write "    </td>"

Response.Write "<td height='30'>"
Response.Write "    <table width='100%' border='0' align='center' cellpadding='1' cellspacing='1'><tr><td height='22' align='right'>"
Response.Write "&nbsp; ������ǰĿ¼�ļ���</td><td height='22'><input type='text' name='SearchKeyword' id='SearchKeyword' size='18' value=''>&nbsp;</td><td height='22'><input type='submit' name='submit1' value=' ���� '>"
Response.Write "    </td></tr></table></td>"
Response.Write "  </tr>"
Response.Write "</table>" & vbCrLf
If ObjInstalled_FSO = False Then
    Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
    Response.End
End If

Select Case Action
Case "Del"
    Call DelFiles
Case "DelThisFolder"
    Call DelThisFolder
Case "DelCurrentDir"
    Call DelCurrentDir
Case "DelAll"
    Call DelAll
Case "DoAddWatermark"
    Call DoAddWatermark
Case "DoAddWatermark_CurrentDir"
    Call DoAddWatermark_CurrentDir
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
If currentSlot > -1 And FoundErr = False Then
    Response.Write "<Script Language=""JavaScript"">" & vbCrLf
    Response.Write "setTimeout('Change()',1000);"
    Response.Write "function Change(){"
    Response.Write "var Sort=document.getElementById(""Sort" & sortBy & """);" & vbCrLf
    If reverse Then
        Response.Write "    Sort.src=""Images/Calendar_Down.gif"";" & vbCrLf
    Else
        Response.Write "    Sort.src=""Images/Calendar_Up.gif"";" & vbCrLf
    End If
    Response.Write "    Sort.style.display="""";    " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</Script>" & vbCrLf
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim Add2Array
    If fso.FolderExists(TruePath) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ����ļ��У����ϴ��ļ����ٽ��й���</li>"
        Exit Sub
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
    '���õ���
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
        FileCount = currentSlot ' �ļ�����
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
End Sub

Sub ShowFileThumb()
    If SearchKeyword = "" Then
        Response.Write "<br><table width='100%' cellpadding='2' cellspacing='1'><tr height='22'><td align='left'>��ǰĿ¼��" & RootDir
        If ParentDir <> "" Then
            Response.Write "/" & ParentDir
        End If
        If CurrentDir <> "" Then
            Response.Write "/" & CurrentDir
        End If
        Response.Write "</td>" & vbCrLf
        Response.Write "    <td align='right'>" & vbCrLf
        If CurrentDir <> "" Then
            If ParentDir <> "" Then
                If InStrRev(ParentDir, "/") > 0 Then
                    Response.Write "<a href='" & strFileName & "&ParentDir=" & Replace(Left(ParentDir, InStrRev(ParentDir, "/") - 1),"/","|")
                    Response.Write "&CurrentDir=" & Mid(ParentDir, InStrRev(ParentDir, "/") + 1)
                Else
                    Response.Write "<a href='" & strFileName & "&ParentDir=&CurrentDir=" & Replace(ParentDir,"/","|")
                End If
            Else
                Response.Write "<a href='" & strFileName
            End If
            Response.Write "'>�������ϼ�Ŀ¼</a>"
        End If
        Response.Write "</td></tr></table>" & vbCrLf
        Response.Write "<table width='100%' cellpadding='2' cellspacing='1' class='border'><tr class='title' height='22'><td colspan='20'><b>��Ŀ¼����</b>" & vbCrLf
        Response.Write "</td></tr><tr class='tdbg'>"
        Dim FolderCount
        Set theFolder = fso.GetFolder(TruePath)
        For Each theSubFolder In theFolder.SubFolders
            If ParentDir <> "" Then
                Response.Write "<td><a href='" & strFileName & "&ParentDir=" & Replace(ParentDir,"/","|") & "|" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
            Else
                Response.Write "<td><a href='" & strFileName & "&ParentDir=" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
            End If
            FolderCount = FolderCount + 1
            If FolderCount Mod 10 = 0 Then Response.Write "</td><tr class='tdbg'>"
        Next
        Response.Write "</tr></table><br>" & vbCrLf
    Else
        Response.Write "<br>&gt;&gt;&nbsp;��ǰĿ¼�ļ����к��е� <font color='red'>" & SearchKeyword & "</font> �ļ�"
    End If
    
    Response.Write "    <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf

    Response.Write "    <tr>" & vbCrLf
    If currentSlot > -1 Then
        Response.Write "    <td height='18'>����ʽ��&nbsp;&nbsp;<a href=""javascript:reSort(0);"">�ļ���&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort0'></a>" & vbCrLf
        'Response.Write "   <a href=""javascript:reSort(1);"">��չ��</a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(2);"">��С&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort2'></a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(3);"">����&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort3'></a>" & vbCrLf
        'Response.Write "   <a href=""javascript:reSort(4);"">����ʱ��</a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(5);"">�ϴ��޸�ʱ��&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort5'></a></td>" & vbCrLf
    Else
        Response.Write "    <td height='18'></td>" & vbCrLf
    End If
    Response.Write "    <td align='right'>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf

    If currentSlot = -1 Then
        Response.Write "<tr class='tdbg'><td align='center' colspan='2'><br><br>��ǰĿ¼��û���κ��ļ���<br><br></td>"
        Response.Write " </tr>"
        Response.Write "</table>" & vbCrLf
    Else
        strFileName = strFileName & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir

        TotalSize = 0
        TotalUnit = 1
        For Each theFile In theFolder.Files
            
            If TotalUnit = 1 Then
                TotalSize = TotalSize + theFile.size / 1024
            ElseIf TotalUnit = 2 Then
                TotalSize = TotalSize + theFile.size / 1024 / 1024
            ElseIf TotalUnit = 3 Then
                TotalSize = TotalSize + theFile.size / 1024 / 1024 / 1024
            End If
            If TotalSize > 1024 Then
                TotalSize = TotalSize / 1024
                TotalUnit = TotalUnit + 1
            End If
            If TotalUnit = 1 Then
                strTotalUnit = "KB"
            ElseIf TotalUnit = 2 Then
                strTotalUnit = "MB"
            ElseIf TotalUnit = 3 Then
                strTotalUnit = "GB"
            End If
        Next
        TotalSize = Round(TotalSize, 2)
        totalPut = FileCount + 1
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
            If (CurrentPage - 1) * MaxPerPage >= totalPut Then
                CurrentPage = 1
            End If
        End If

        Dim c
        Dim theFileName, tUsed, FileNum
        FileNum = 0
        TotalSize_Page = 0
        PageUnit = 1

        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr>"

        Response.Write "     <td colspan='2'><table width='100%' border='0' align='center' cellpadding='0' cellspacing='3' class='border'>"
        Response.Write "  <tr class='tdbg'>" & vbCrLf

        For i = 0 To FileCount
            c = c + 1
            If FileNum >= MaxPerPage Then
                Exit For
            ElseIf c > MaxPerPage * (CurrentPage - 1) Then
                Response.Write "    <td>"
                Response.Write "      <table width='100%' height='100%' border='0' cellpadding='0' cellspacing='2'>"
                Response.Write "        <tr>"
                Response.Write "          <td colspan='2' align='center'>"
                theFileName = strPath & "/" & theFiles(i)(0)

                Response.Write "<a href='" & theFileName & "'>"
                'StrFileType = LCase(Mid(theFile.Name, InStrRev(theFile.Name, ".") + 1))
                'Select Case StrFileType
                Select Case LCase(theFiles(i)(1))
                Case "gif", "jpg", "jpeg", "jpe", "bmp", "png"
                    Response.Write "<img src='" & theFileName & "'"
                Case "swf"
                    Response.Write "<img src='images/filetype_flash.gif'"
                Case "wmv", "avi", "asf", "mpg"
                    Response.Write "<img src='images/filetype_media.gif'"
                Case "rm", "ra", "ram"
                    Response.Write "<img src='images/filetype_rm.gif'"
                Case "rar", "zip", "exe"
                    Response.Write "<img src='images/filetype_" & theFiles(i)(1) & ".gif'"
                Case Else
                    Response.Write "<img src='images/filetype_other.gif'"
                End Select
                Response.Write " width='130' height='90'"
                If InStr(strFiles, LCase(theFiles(i)(0))) > 0 Then
                    tUsed = True
                Else
                    tUsed = False
                End If
                If tUsed = True Then
                    Response.Write " border='0' Title='�� �� ����" & theFiles(i)(0) & vbCrLf & "�ļ���С��" & Round(theFiles(i)(2) / 1024) & " KB" & vbCrLf & "�ļ����ͣ�" & theFiles(i)(3) & vbCrLf & "�޸�ʱ�䣺" & theFiles(i)(5) & "'>"
                Else
                    Response.Write " border='2' Title='���õ��ϴ��ļ�" & vbCrLf & "�� �� ����" & theFiles(i)(0) & vbCrLf & "�ļ���С��" & Round(theFiles(i)(2) / 1024) & " KB" & vbCrLf & "�ļ����ͣ�" & theFiles(i)(3) & vbCrLf & "�޸�ʱ�䣺" & theFiles(i)(5) & "'>"
                End If

                Response.Write "</a>"
                Response.Write "          </td>"
                Response.Write "        </tr>" & vbCrLf
                Response.Write "        <tr>"
                'Response.Write "          <td align='right'>�ļ�����</td>"
                Response.Write "          <td align='center'>"
                If tUsed = True Then
                    Response.Write "<a href='" & theFileName & "' target='_blank'>" & CutStr(theFiles(i)(0)) & "</a>"
                Else
                    Response.Write "<a href='" & theFileName & "' target='_blank' title='���õ��ϴ��ļ�'><font color=red>" & CutStr(theFiles(i)(0)) & "</font></a>"
                End If

                Response.Write "       </td>"
                Response.Write "        </tr>" & vbCrLf

                Response.Write "        <tr>"
                'Response.Write "          <td align='right'>������</td>"
                Response.Write "          <td align='center'><input name='FileName' type='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
                If tUsed = False Then Response.Write " checked"
                Response.Write "> ѡ��&nbsp;&nbsp;<a href='" & strFileName & "&Action=Del&FileName=" & theFiles(i)(0) & "' onclick=""return confirm('�����Ҫɾ�����ļ���!');"">ɾ��</a></td>"
                Response.Write "        </tr>" & vbCrLf
                Response.Write "      </table>"
                Response.Write "    </td>" & vbCrLf
                FileNum = FileNum + 1
                If FileNum Mod 4 = 0 Then Response.Write "</td><tr class='tdbg'>"
                If PageUnit = 1 Then
                    TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024
                ElseIf PageUnit = 2 Then
                    TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024 / 1024
                ElseIf PageUnit = 3 Then
                    TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024 / 1024 / 1024
                End If
                If TotalSize_Page > 1024 Then
                    TotalSize_Page = TotalSize_Page / 1024
                    PageUnit = PageUnit + 1
                End If
                If PageUnit = 1 Then
                    strPageUnit = "KB"
                ElseIf PageUnit = 2 Then
                    strPageUnit = "MB"
                ElseIf PageUnit = 3 Then
                    strPageUnit = "GB"
                End If
            End If
        Next
        TotalSize_Page = Round(TotalSize_Page, 2)

        Response.Write "  </tr>"
        Response.Write "</table>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form);' value='checkbox'> ѡ�б�ҳ�����ļ�</td><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "    <td><input name='Action' type='hidden' id='Action' value=''><input name='UploadDir' type='hidden' value='" & UploadDir & "'><input name='CurrentDir' type='hidden' value='" & CurrentDir & "'>"
        Response.Write "        <input type='submit' name='Submit' value='ɾ��ѡ�е��ļ�' onclick=""document.myform.Action.value='Del';return confirm('ȷ��Ҫɾ��ѡ�е��ļ���');"">&nbsp;&nbsp;<input type='submit' name='Submit2' value='ɾ����ǰĿ¼�������ļ�' onClick=""document.myform.Action.value='DelCurrentDir';return confirm('ȷ��Ҫɾ����ǰĿ¼�µ������ļ���')"">"
        If ParentDir = "" And CurrentDir = "" Then
            Response.Write "&nbsp;&nbsp;<input type='submit' name='Submit2' value='ɾ�������ļ�����Ŀ¼' onClick=""document.myform.Action.value='DelAll';return confirm('ȷ��Ҫɾ�������ļ�����Ŀ¼��')"">"
        End If
        
        If IsObjInstalled("Persits.Jpeg") = True Then
            Response.Write "&nbsp;&nbsp;<br><input type='submit' name='Submit3' onClick=""document.myform.Action.value='DoAddWatermark'"" value='��ѡ�е�ͼƬ���ˮӡ' >&nbsp;&nbsp;<input type='submit' name='Submit4' onClick=""document.myform.Action.value='DoAddWatermark_CurrentDir'""  value='����ǰĿ¼���ͼƬˮӡ' >"
        End If
       
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf

        Response.Write "<input type='hidden' name='priorsort' value='" & priorSort & "'>" & vbCrLf
        Response.Write "<input type='hidden' name='sortby' value='-1'>" & vbCrLf


        Response.Write "</td></form></tr></table>" & vbCrLf
        Response.Write showpage2(strFileName, totalPut, MaxPerPage)
        Response.Write "<br><div align='center'>��ҳ����ʾ <b>" & FileNum & "</b> ���ļ���ռ�� <b>" & TotalSize_Page & "</b> " & strPageUnit & "</div>"

    End If

End Sub

Sub ShowFileDetail()
    If SearchKeyword <> "" Then
        Response.Write "<br>&gt;&gt;&nbsp;��ǰĿ¼�ļ����к��е� <font color='red'>" & SearchKeyword & "</font> �ļ�"
    Else
        Response.Write "<br>"
    End If
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0' ><tr><td>��ǰĿ¼��" & RootDir
    If ParentDir <> "" Then
        Response.Write "/" & ParentDir
    End If
    If CurrentDir <> "" Then
        Response.Write "/" & CurrentDir & "</td><td align='right'>"
        If ParentDir <> "" Then
            If InStrRev(ParentDir, "/") > 0 Then
                Response.Write "<a href='" & strFileName & "&ParentDir=" & replace(Left(ParentDir, InStrRev(ParentDir, "/") - 1),"/","|")
                Response.Write "&CurrentDir=" & Mid(ParentDir, InStrRev(ParentDir, "/") + 1)
            Else
                Response.Write "<a href='" & strFileName & "&ParentDir=&CurrentDir=" & ParentDir
            End If
        Else
            Response.Write "<a href='" & strFileName
        End If
        Response.Write "'>�������ϼ�Ŀ¼</a>"
    End If
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "    <table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf

    Response.Write "    <tr>" & vbCrLf
    Response.Write "    <td height='18' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'"">&nbsp;&nbsp;<a href=""javascript:reSort(0);"">�ļ���&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort0'></a></td>" & vbCrLf
    'Response.Write "   <a href=""javascript:reSort(1);"">��չ��</a>" & vbCrLf
    Response.Write "    <td width='80' align=""right"" class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'""><a href=""javascript:reSort(2);"">��С&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort2'></a>&nbsp;</td>" & vbCrLf
    Response.Write "    <td width='180' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'"">&nbsp;<a href=""javascript:reSort(3);"">����&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort3'></a></td>" & vbCrLf
    'Response.Write "   <a href=""javascript:reSort(4);"">����ʱ��</a>" & vbCrLf
    Response.Write "    <td width='140' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'""><a href=""javascript:reSort(5);"">�ϴ��޸�ʱ��&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort5'></a></td>" & vbCrLf
    Response.Write "    <td width='30' align='center' class='title0'>����&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    If sortBy <> 0 Then
        Call ShowFileDetail_Fil
        If SearchKeyword = "" Then
            Response.Write "<tr><td height=1></td></tr>" & vbCrLf
            Call ShowFileDetail_fol
        End If
    Else
        If SearchKeyword = "" Then
            Call ShowFileDetail_fol
            Response.Write "<tr><td height=1></td></tr>" & vbCrLf
        End If
        Call ShowFileDetail_Fil
    End If
    Response.Write "    </table>" & vbCrLf
    Response.Write "<input type='hidden' name='priorsort' value='" & priorSort & "'>" & vbCrLf
    Response.Write "<input type='hidden' name='sortby' value='-1'>" & vbCrLf
    If currentSlot > -1 Then
        Call ShowJS_Tooltip
        Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form);' value='checkbox'> ѡ�б�ҳ�����ļ�</td><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "    <td><input name='Action' type='hidden' id='Action' value=''><input name='UploadDir' type='hidden' value='" & UploadDir & "'><input name='CurrentDir' type='hidden' value='" & CurrentDir & "'>"
        Response.Write "        <input type='submit' name='Submit' value='ɾ��ѡ�е��ļ�' onclick=""document.myform.Action.value='Del';return confirm('ȷ��Ҫɾ��ѡ�е��ļ���');"">&nbsp;&nbsp;<input type='submit' name='Submit2' value='ɾ����ǰĿ¼�������ļ�' onClick=""document.myform.Action.value='DelCurrentDir';return confirm('ȷ��Ҫɾ����ǰĿ¼�µ������ļ���')"">"
        If ParentDir = "" And CurrentDir = "" Then
            Response.Write "&nbsp;&nbsp;<input type='submit' name='Submit2' value='ɾ�������ļ�����Ŀ¼' onClick=""document.myform.Action.value='DelAll';return confirm('ȷ��Ҫɾ�������ļ�����Ŀ¼��')"">"
        End If
        
        If IsObjInstalled("Persits.Jpeg") = True Then
            Response.Write "&nbsp;&nbsp;<br><input type='submit' name='Submit3' onClick=""document.myform.Action.value='DoAddWatermark'"" value='��ѡ�е�ͼƬ���ˮӡ' >&nbsp;&nbsp;<input type='submit' name='Submit4' onClick=""document.myform.Action.value='DoAddWatermark_CurrentDir'""  value='����ǰĿ¼���ͼƬˮӡ' >"
        End If
       
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf
    End If

    
    Response.Write "    </form>" & vbCrLf

End Sub

Function GetUploadFileStyle()
    ShowFileStyle = Request.Cookies("ShowFileStyle")
    If ShowFileStyle = "" Or Not IsNumeric(ShowFileStyle) Then
        ShowFileStyle = 1
    Else
        ShowFileStyle = Int(ShowFileStyle)
    End If
    GetUploadFileStyle = ShowFileStyle
End Function



Sub ShowFileDetail_Fil()
    If currentSlot > -1 Then

        For i = 0 To FileCount
            Response.Write "<tr onmouseout=""this.className='tdbgmouseout1'"" onmouseover=""this.className='tdbg1'"">" & vbCrLf
            If InStr(strFiles, LCase(theFiles(i)(0))) > 0 Then
                'Response.Write " Title='�� �� ����" & theFiles(i)(0) & vbCrLf & "�ļ���С��" & Round(theFiles(i)(2) / 1024) & " K" & vbCrLf & "�ļ����ͣ�" & theFiles(i)(3) & vbCrLf & "�޸�ʱ�䣺" & theFiles(i)(5) & "'>"
                Response.Write "          <td align='left'><input name='FileName' type='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
                Response.Write ">"
            Else
                'Response.Write " Title='���õ��ϴ��ļ�" & vbCrLf & "�� �� ����" & theFiles(i)(0) & vbCrLf & "�ļ���С��" & Round(theFiles(i)(2) / 1024) & " K" & vbCrLf & "�ļ����ͣ�" & theFiles(i)(3) & vbCrLf & "�޸�ʱ�䣺" & theFiles(i)(5) & "'>"
                Response.Write "          <td align='left'><input name='FileName' type='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
                Response.Write " checked"
                Response.Write ">"
            End If

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
            Response.Write "<a href='" & strPath & "/" & theFiles(i)(0) & "'><span onmouseover=""ShowADPreview('" & FixJs(GetFileContent(strPath & "/" & theFiles(i)(0), theFiles(i)(1))) & "')"" onmouseout=""hideTooltip('dHTMLADPreview')"">" & vbCrLf
            Response.Write theFiles(i)(0) & "</span></a></td>" & vbCrLf
            Response.Write " <td width='80' align='right'>" & FormatNumber(theFiles(i)(2) / 1024, 0, vbTrue, vbFalse, vbTrue) & " KB</td>" & vbCrLf
            Response.Write " <td width='180'>&nbsp;" & CutStr(theFiles(i)(3)) & "</td>" & vbCrLf
            Response.Write " <td width='140'>" & theFiles(i)(5) & "</td>" & vbCrLf
            Response.Write "<td width='30' align='center'><a href='" & strFileName & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir & "&Action=Del&FileName=" & theFiles(i)(0) & "' onclick=""return confirm('�����Ҫɾ�����ļ���!');"">ɾ��</a>&nbsp;"
            Response.Write "</td></tr>" & vbCrLf
        Next
    End If
End Sub

Function GetFileContent(ByVal sPath, sType)
    If IsNull(sPath) Or sPath = "" Then
        GetFileContent = "&nbsp;���ļ���ͼƬ�򶯻�����Ԥ��&nbsp;"
        Exit Function
    End If
    If IsNull(sType) Or sType = "" Then
        GetFileContent = "&nbsp;���ļ���ͼƬ�򶯻�����Ԥ��&nbsp;"
        Exit Function
    End If

    If Not fso.FileExists(Server.MapPath(sPath)) Then
        GetFileContent = "&nbsp;���ļ�������&nbsp;"
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
        strFile = "&nbsp;���ļ���ͼƬ�򶯻�����Ԥ��&nbsp;"
    End Select

    GetFileContent = strFile
End Function

Sub ShowFileDetail_fol()
    Response.Write "<tr>"
    For Each theSubFolder In theFolder.SubFolders
        If ParentDir <> "" Then
            Response.Write "<td height='18'>&nbsp;&nbsp;<img src='Images/Folder/folderclosed.gif'><a href='" & strFileName & "&ParentDir=" & replace(ParentDir,"/","|") & "|" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        Else
            Response.Write "<td height='18'>&nbsp;&nbsp;<img src='Images/Folder/folderclosed.gif'><a href='" & strFileName & "&ParentDir=" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        End If
        Response.Write "<td width='50' align=""right"">&nbsp;</td>"
        Response.Write "<td width='180'>&nbsp;�ļ���</td>"
        Response.Write "<td width='140'>" & theSubFolder.DateLastModified & "</td>"

        Response.Write "<td width='30' align='center'><a href='" & strFileName & "&ParentDir=" & Replace(ParentDir,"/","|") & "|" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "&Action=DelThisFolder' onclick=""return confirm('�����Ҫɾ�����ļ��м�������ļ���!');"">ɾ��</a>&nbsp;"
        Response.Write "</td></tr><tr>"
        
    Next
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

Sub DelFiles()
    Dim whichfile, arrFileName, i
    whichfile = Trim(Request("FileName"))
    If whichfile = "" Then Exit Sub
    If InStr(whichfile, ",") > 0 Then
        arrFileName = Split(whichfile, ",")
        For i = 0 To UBound(arrFileName)
            whichfile = Server.MapPath(strPath & "/" & Trim(arrFileName(i)))
            If fso.FileExists(whichfile) Then fso.DeleteFile whichfile
        Next
    Else
        whichfile = Server.MapPath(strPath & "/" & whichfile)
        If fso.FileExists(whichfile) Then fso.DeleteFile whichfile
    End If
    Call main
End Sub

Sub DelCurrentDir()
    Set theFolder = fso.GetFolder(Server.MapPath(strPath))
    For Each theFile In theFolder.Files
        theFile.Delete True
    Next
    Call main
End Sub

Sub DelAll()
    Set theFolder = fso.GetFolder(Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir))
    For Each theSubFolder In theFolder.SubFolders
        theSubFolder.Delete True
    Next
    For Each theFile In theFolder.Files
        theFile.Delete True
    Next
    Call main
End Sub

Sub DelThisFolder()
    On Error Resume Next
    fso.DeleteFolder Server.MapPath(strPath)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>ɾ���ļ���ReplaceDBContent.asp��ʧ�ܣ�����ԭ��" & Err.Description & "<br>���ֶ�ɾ�����ļ���"
        Err.Clear
        Exit Sub
    Else
        If SearchKeyword <> "" Then
            Call Refresh("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & Replace(ParentDir,"/","|") & "&SearchKeyword=" & SearchKeyword,0)		
            'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&SearchKeyword=" & SearchKeyword & """>"
        Else
            Call Refresh("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & Replace(ParentDir,"/","|"),0)		
            'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & """>"
        End If
    End If
End Sub

Sub DoAddWatermark()
    Dim whichfile, arrFileName, i, bTemp
    whichfile = Trim(Request("FileName"))
    If whichfile = "" Then Exit Sub

    Dim PE_Thumb
    Set PE_Thumb = New CreateThumb
    If InStr(whichfile, ",") > 0 Then
        arrFileName = Split(whichfile, ",")
        For i = 0 To UBound(arrFileName)
            whichfile = strPath & "/" & Trim(arrFileName(i))
            bTemp = PE_Thumb.AddWatermark(whichfile)
        Next
    Else
        whichfile = strPath & "/" & whichfile
        bTemp = PE_Thumb.AddWatermark(whichfile)
    End If

    Set PE_Thumb = Nothing
    If SearchKeyword <> "" Then
        Call Refresh("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword,0)		
        'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & """>"
    Else
        Call Refresh("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir,0)	
        'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & """>"
    End If
End Sub

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
    Response.Write "       locateObject(object).innerHTML='<table style=""font-family:����; font-size: 9pt; border: '+bordercolor+'; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; background-color: '+backcolor+'"" width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td nowrap><font style=""font-family:����; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table> '" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clientWidth) > (document.body.clientWidth + document.body.scrollLeft)) {" & vbCrLf
    Response.Write "           locateObject(object).style.left = (document.body.clientWidth + document.body.scrollLeft) - locateObject(object).clientWidth-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).style.left=document.body.scrollLeft+event.clientX" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).style.visibility='visible';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else if (document.layers) {" & vbCrLf
    Response.Write "       locateObject(object).document.write('<table width=""10"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr bgcolor=""'+bordercolor+'""><td><table width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr bgcolor=""'+backcolor+'""><td nowrap><font style=""font-family:����; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table></td></tr></table>')" & vbCrLf
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
Sub DoAddWatermark_CurrentDir()
    Dim whichfile, bTemp
    Dim PE_Thumb
    Set PE_Thumb = New CreateThumb
    Set theFolder = fso.GetFolder(Server.MapPath(strPath))
    For Each theFile In theFolder.Files
        whichfile = strPath & "/" & theFile.name
        bTemp = PE_Thumb.AddWatermark(whichfile)
    Next
    'Call main
    Set PE_Thumb = Nothing
    If SearchKeyword <> "" Then
        'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & """>"
        Call Refresh("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword,0)		
    Else
       ' Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & """>"
        Call Refresh("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & Replace(ParentDir,"/","|") & "&CurrentDir=" & CurrentDir,0)			
    End If
End Sub

Function showpage2(sfilename, totalnumber, MaxPerPage)
    Dim n, i, strTemp
    If totalnumber Mod MaxPerPage = 0 Then
        n = totalnumber \ MaxPerPage
    Else
        n = totalnumber \ MaxPerPage + 1
    End If
    If SearchKeyword <> "" Then
        strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "&SearchKeyword=" & SearchKeyword & "'><tr><td>"
    Else
         strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
    End If
    strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
    strTemp = strTemp & "�� <b>" & totalnumber & "</b> ���ļ���ռ�� <b>" & TotalSize & "</b> " & strTotalUnit & "&nbsp;&nbsp;&nbsp;"
    sfilename = JoinChar(sfilename)
    If CurrentPage < 2 Then
            strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
    Else
            strTemp = strTemp & "<a href='" & sfilename & "page=1'>��ҳ</a>&nbsp;"
            strTemp = strTemp & "<a href='" & sfilename & "page=" & (CurrentPage - 1) & "'>��һҳ</a>&nbsp;"
    End If

    If n - CurrentPage < 1 Then
            strTemp = strTemp & "��һҳ βҳ"
    Else
            strTemp = strTemp & "<a href='" & sfilename & "page=" & (CurrentPage + 1) & "'>��һҳ</a>&nbsp;"
            strTemp = strTemp & "<a href='" & sfilename & "page=" & n & "'>βҳ</a>"
    End If
    strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & "���ļ�/ҳ"
    strTemp = strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:submit()'>"
    For i = 1 To n
        strTemp = strTemp & "<option value='" & i & "'"
        If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
        strTemp = strTemp & ">��" & i & "ҳ</option>"
    Next
    strTemp = strTemp & "</select>"
    strTemp = strTemp & "</td></tr></form></table>"
    showpage2 = strTemp
End Function


Function CutStr(Str)
    If Len(Str) > 18 Then
        CutStr = "..." & Right(Str, 18)
    Else
        CutStr = Str
    End If
End Function
%>
