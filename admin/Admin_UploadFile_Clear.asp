<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim TruePath, theFolder, theSubFolder, theFile, thisfile, FileCount, TotalSize, TotalSize_Page
Dim TotalUnit, strTotalUnit, PageUnit, strPageUnit
Dim StrFileType, strFiles
Dim strDirName, tUploadDir
Dim RootDir, ParentDir, CurrentDir
Dim strPath, strPath2, strPath3
Dim ItemIntro, UpFileType

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

strFileName = "Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir

Response.Write "<html><head><title>�ϴ��ļ�����</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
Call ShowPageTitle(ChannelName & "����----�ϴ��ļ�����", 10012)
Response.Write "  <tr class='tdbg'> "
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>"
Response.Write "    <td height='30'><a href='" & strFileName & "'>�ϴ��ļ�������ҳ</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&Action=Clear'>��������ļ�</a></td>"
Response.Write "  </tr>"
Response.Write "</table>" & vbCrLf
If ObjInstalled_FSO = False Then
    Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
    Response.End
End If

Select Case Action
Case "Clear"
    Call Clear
Case "DoClear"
    Call DoClear
Case Else
    Call Clear
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn



Sub Clear()
    Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>�������õ��ϴ��ļ�</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
    Response.Write "<form name='form1' method='post' action='Admin_UploadFile_Clear.asp'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;���������ʱ������������ϴ���ͼƬ��ȴ����û��ʹ�õ������ʱ��һ�ã��ͻ�����������������ļ���������Ҫ����ʹ�ñ����ܽ�������"
    Response.Write "<p>&nbsp;&nbsp;&nbsp;&nbsp;����ϴ��ļ��ܶ࣬������Ϣ�����϶ִ࣬�б�������Ҫ�ķ��൱����ʱ�䣬���ڷ�������ʱִ�б�������</p>"
    Response.Write "<p align='center'><input name='Action' type='hidden' id='Action' value='DoClear'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='UploadDir' type='hidden' value='" & tUploadDir & "'><input name='CurrentDir' type='hidden' value='" & CurrentDir & "'><input type='submit' name='Submit3' value=' ��ʼ���� '></p>"
    Response.Write "</form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub


Sub DoClear()
    ParentDir = Replace(Replace(Replace(Trim(Request("ParentDir")), "../", ""), "..\", ""), "\", "/")
    If Left(ParentDir, 1) = "/" Then ParentDir = Right(ParentDir, Len(ParentDir) - 1)
    CurrentDir = Replace(Replace(Replace(Trim(Request("CurrentDir")), "/", ""), "\", ""), "..", "")
    
    Dim rs, sql
    Select Case ModuleType
    Case 1
        strDirName = ChannelName & "���ϴ��ļ�"
        sql = "select UploadFiles,Intro from PE_Article where ChannelID=" & ChannelID
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            If rs(1) <> "" Then
                ItemIntro = ItemIntro & "|" & rs(1)
            End If
            rs.MoveNext
        Loop
    Case 2
        If tUploadDir = "UploadSoftPic" Then
            UploadDir = "UploadSoftPic"
            strDirName = ChannelName & "�����ͼƬ"
            sql = "select SoftPicUrl,SoftIntro from PE_Soft where ChannelID=" & ChannelID
            Set rs = Conn.Execute(sql)
            Do While Not rs.EOF
                If rs(0) <> "" Then
                    strFiles = strFiles & "|" & rs(0)
                End If
                If rs(1) <> "" Then
                    ItemIntro = ItemIntro & "|" & rs(1)
                End If
                rs.MoveNext
            Loop
        Else
            strDirName = ChannelName & "���ϴ����"
            sql = "select DownloadUrl,SoftIntro from PE_Soft where ChannelID=" & ChannelID
            Set rs = Conn.Execute(sql)
            Do While Not rs.EOF
                If rs(0) <> "" Then
                    strFiles = strFiles & "$$$" & rs(0)
                End If
                If rs(1) <> "" Then
                    ItemIntro = ItemIntro & "|" & rs(1)
                End If				
                rs.MoveNext
            Loop
        End If
    Case 3
        strDirName = ChannelName & "���ϴ�ͼƬ"
        sql = "select PhotoThumb,PhotoUrl,PhotoIntro from PE_Photo"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "$$$" & rs(0)
            End If
            If rs(1) <> "" Then
                strFiles = strFiles & "$$$" & rs(1)
            End If
            If rs(2) <> "" Then
                ItemIntro = ItemIntro & "|" & rs(2)
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
        Dim i, HouseTable
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

    If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Then
        Dim tempStr, tempi, TempArray
		UpFileType = "gif|jpg|jpeg|jpe|bmp|png"
        regEx.Pattern = "<img.+?[^\>]>" '��ѯ���������� <img..>
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '�ۼ�����
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '�ָ�����
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "src\s*=\s*.+?\.(" & UpFileType & ")" '��ѯsrc =�ڵ�����
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '�ۼӵõ� ���Ӽ�$Array$ �ַ�
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "src\s*=\s*" '���� src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		
        strFiles = strFiles & tempStr

        UpFileType = "swf|rm|ra|ram"
        regEx.Pattern = "<param\s*name\s*=\s*""*src""*.+?[^\>]>" 
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '�ۼ�����
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '�ָ�����
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "value\s*=\s*.+?\.(" & UpFileType & ")" '��ѯvalue =�ڵ�����
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '�ۼӵõ� ���Ӽ�$Array$ �ַ�
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "value\s*=\s*" '���� src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		
        strFiles = strFiles & tempStr
		

        UpFileType = "mp3|avi|wmv|mpg|asf"
        regEx.Pattern = "<param\s*name\s*=\s*""*url""*.+?[^\>]>"
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '�ۼ�����
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '�ָ�����
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "value\s*=\s*.+?\.(" & UpFileType & ")" '��ѯvalue =�ڵ�����
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '�ۼӵõ� ���Ӽ�$Array$ �ַ�
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "value\s*=\s*" '���� src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		strFiles = strFiles & tempStr	
		

        UpFileType = "zip|rar|doc"
        regEx.Pattern = "<a.+?[^\>](rar\""*\s*|zip\""*\s*|doc\""*\s*)>" '��ѯ���������к�zip��rar��doc���ֵĸ���
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '�ۼ�����
            Else
                tempStr = Match.value
            End If
        Next
		
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '�ָ�����
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "href\s*=\s*.+?\.(" & UpFileType & ")" '��ѯhref =�ڵ�����
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '�ۼӵõ� ���Ӽ�$Array$ �ַ�
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "href\s*=\s*" '���� href =
            tempStr = regEx.Replace(tempStr, "")
        End If
        strFiles = strFiles & tempStr					
		
    End If

    strFiles = LCase(strFiles)
    
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

    i = 0
    If fso.FolderExists(Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir)) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ����ļ��У����ϴ��ļ����ٽ��й���</li>"
        Exit Sub
    End If

    Set theFolder = fso.GetFolder(Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir))
    For Each theFile In theFolder.Files
        If InStr(strFiles, LCase(theFile.name)) <= 0 Then
            theFile.Delete True
            i = i + 1
        End If
    Next
    For Each theSubFolder In theFolder.SubFolders
        For Each theFile In theSubFolder.Files
            If InStr(strFiles, LCase(theSubFolder.name & "/" & theFile.name)) <= 0 Then
                theFile.Delete True
                i = i + 1
            End If
        Next
    Next

    Call WriteSuccessMsg("���������ļ��ɹ�����ɾ���� <font color=red><b>" & i & "</b></font> �����õ��ļ���", ComeUrl)
End Sub
%>
