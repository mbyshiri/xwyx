<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
<!--#include file="../Include/PowerEasy.Edition.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
Const NeedCheckComeUrl = True

Const AdminType = True
Const EnableGuestCheck = "Yes"

Dim FilesNum, i, theFiles, ObjInstalled_XML
Dim FileInfoURL

FileInfoURL = "http://www.powereasy.net/FileList/SiteWeaver/" & SystemEdition & ".txt"


ObjInstalled_XML = IsObjInstalled("MSXML2.XMLHTTP")

strFileName = "Admin_CompareFilesOnline.asp"
                                                    
Response.Write "<html><head><title>�ϴ��ļ�����</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
Call ShowPageTitle("���߱Ƚ���վ�ļ�", 10031)

If Action <> "" Then
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='70' height='30'><strong>��������</strong></td>"
    Response.Write "    <td height='30'><a href='" & strFileName & "?Action=ShowAllResult'>ȫ����ʾ</a>&nbsp;|&nbsp;<a href='" & strFileName & "?Action=ShowOnlyDif'>ֻ��ʾ���첿��</a></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='80' height='30'><strong>����ĺ��壺</strong></td>"
    Response.Write "    <td height='30'> " & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "<tr>" & vbCrLf
    Response.Write "    <td><b>'= '</b>----���ߴ�Сʱ����ȫ��ͬ</td>" & vbCrLf
    Response.Write "    <td><font color='red'><b>'��'</b></font>----���ߴ�С����ͬ</td>" & vbCrLf
    Response.Write "    <td><font color='gray'><b>'��'</b></font>----���߽���ʱ�䲻ͬ</td>" & vbCrLf
    Response.Write "</tr><tr>" & vbCrLf
    Response.Write "    <td><font color='red'>��ɫ</font>----����ͬ���޸Ļ���¹����ļ�</td>" & vbCrLf
    Response.Write "    <td><font color='blue'>��ɫ</font>----���ز����ڵ��ļ�</td>" & vbCrLf
    Response.Write "    <td><font color='gray'>��ɫ</font>----�ٷ������ļ���������δ���µ��ļ�</td>" & vbCrLf
    Response.Write "</tr><tr>" & vbCrLf
    Response.Write "    <td><font color='black'>��ɫ</font>----��ͬ�ļ���ٷ��ļ�</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</td>"
    Response.Write "  </tr>"
Else
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='80' height='30'><strong>��������</strong></td>"
    Response.Write "    <td height='30'><a href='" & strFileName & "'>���߱Ƚ���վ�ļ���Ϣ</a> </td>"
    Response.Write "  </tr>"
End If

Response.Write "</table>" & vbCrLf
If ObjInstalled_FSO = False Then
    Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
    Response.End
End If
If ObjInstalled_XML = False Then
    Response.Write "<b><font color=red>��ķ�������֧�� XMLHTTP ���! ����ʹ�ñ�����</font></b>"
    Response.End
End If


Select Case Action
Case "ShowOnlyDif"
    Call ShowOnlyDif
Case "ShowAllResult"
    Call ShowAllResult
Case Else
    Call Main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub Main()
    Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>���߱Ƚ���վ�ļ���Ϣ</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
    Response.Write "<form name='form1' method='post' action='" & strFileName & "'>"
    Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;����Ա�������ñ����ܣ����߱Ƚ�Web�ռ��е���վASP�ļ��Ͷ��׹ٷ���������Ӧ�汾��ԭʼASP�ļ�������Web�ռ��ļ�����<br>������������ֽԿ���ʹ�ñ����ܽ��бȽϣ�<font color='green'><br>&nbsp;&nbsp;&nbsp;&nbsp;1�����ٷ������ļ�ʱ��<br>&nbsp;&nbsp;&nbsp;&nbsp;2��������վ��ASP�ļ�����ɾ��������޸�ʱ��<br>&nbsp;&nbsp;&nbsp;&nbsp;3�����ٷ�����©������ʱ��</font>"
    Response.Write "<p>&nbsp;&nbsp;&nbsp;&nbsp;�����վƵ���ܶ࣬���������ٶȱȽ�����ִ�б�������Ҫ�ķ��൱����ʱ�䣬���ڷ�������ʱִ�б�������</p>"
    Response.Write "<p align='center'><input name='Action' type='hidden' id='Action' value='ShowAllResult'>"
    Response.Write "<input type='submit' name='Submit3' value=' ��ʼ�Ƚ� '></p>"
    Response.Write "</form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub ShowAllResult()
    Dim Html, GetFiles, FileInfo
    Dim f, fPath, FileSize, FileDate, theFilePath, FileName, interHtml

    Html = GetHttpPage(FileInfoURL, 0)
'response.write Html
'Exit sub
    If Html = "" Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>��ȡ�ٷ�����ʧ�ܣ����������ķ�������֧�� XMLHTTP ���������ͨ������������������硣</font></p>"
        Exit Sub
    End If
    If AdminDir <> "Admin" Then
        Html = Replace(Html, "Admin/", AdminDir & "/")
    End If
    If ADDir <> "AD" Then
        Html = Replace(Html, "AD/", ADDir & "/")
    End If
    GetFiles = Split(Html, vbCrLf)
    FilesNum = UBound(GetFiles)
    ReDim theFiles(FilesNum - 1)
    For i = 0 To FilesNum - 1
        FileInfo = Split(GetFiles(i), "|")
        theFiles(i) = FileInfo
    Next
    '����strChannel���͵�Ƶ����ʼ
    Call ChangeArr("Article", 1)
    Call ChangeArr("Soft", 2)
    Call ChangeArr("Photo", 3)
    '����strChannel���͵�Ƶ�����
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class='border'>" & vbCrLf
    Response.Write "<tr class='title0'>" & vbCrLf
    Response.Write "    <td>&nbsp;����(�ٷ�)</td>" & vbCrLf
    Response.Write "    <td>&nbsp;��С</td>" & vbCrLf
    Response.Write "    <td>&nbsp;&nbsp;�޸�ʱ��</td>" & vbCrLf
    Response.Write "    <td class='tdtop'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "    <td>&nbsp;����(��վ)</td>" & vbCrLf
    Response.Write "    <td>&nbsp;��С</td>" & vbCrLf
    Response.Write "    <td>&nbsp;&nbsp;�޸�ʱ��</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
    Dim j, dyNum, bdyNum, ydyNUm, bczNum
    j = 1
    dyNum = 0
    bdyNum = 0
    ydyNUm = 0
    bczNum = 0

    For i = 0 To FilesNum - 1
        theFilePath = Replace(InstallDir & theFiles(i)(0), "//", "/")
        fPath = Server.MapPath(theFilePath)
        If j Mod 2 = 0 Then
            Response.Write "<tr class='tdbg1' onmouseout=""this.className='tdbg1'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
        Else
            Response.Write "<tr onmouseout=""this.className='tdbgmouseout1'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
        End If
        If fso.FileExists(fPath) Then
            Set f = fso.GetFile(fPath)
            FileName = theFiles(i)(0)
            FileSize = f.size
            FileDate = f.DateLastModified
            If theFiles(i)(1) <> CStr(FileSize) Then
                interHtml = "red'>��"
                bdyNum = bdyNum + 1
            Else
                interHtml = "gray'>��"
                If CDate(theFiles(i)(2)) <> FileDate Then
                    ydyNUm = ydyNUm + 1
                End If
            End If
 
            If theFiles(i)(1) = CStr(FileSize) And CDate(theFiles(i)(2)) = FileDate Then
                Response.Write "    <td><b>��</b>" & theFiles(i)(0) & "</td>" & vbCrLf
                Response.Write "    <td align='right'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</td>" & vbCrLf
                Response.Write "    <td>" & theFiles(i)(2) & "</td>" & vbCrLf
                Response.Write "    <td class='tdinter'><b>=</b></td>" & vbCrLf
                Response.Write "    <td><b>��</b>" & FileName & "</td>" & vbCrLf
                Response.Write "    <td align='right'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</td>" & vbCrLf
                Response.Write "    <td>" & FileDate & "</td>" & vbCrLf
                dyNum = dyNum + 1
            Else
                If CDate(theFiles(i)(2)) > FileDate Then
                    Response.Write "    <td><font color='red'><b>��</b>" & theFiles(i)(0) & "</font></td>" & vbCrLf
                    Response.Write "    <td align='right'><font color='red'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "</font>&nbsp;&nbsp;</td>" & vbCrLf
                    Response.Write "    <td><font color='red'>" & theFiles(i)(2) & "</font></td>" & vbCrLf
                    Response.Write "   <td class='tdinter'><b><font color='" & interHtml & "</font></b></td>" & vbCrLf
                    Response.Write "    <td><font color='gray'><b>��</b>" & FileName & "</font></td>" & vbCrLf
                    Response.Write "    <td align='right'><font color='gray'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "</font>&nbsp;&nbsp;</td>" & vbCrLf
                    Response.Write "    <td><font color='gray'>" & FileDate & "</font></td>" & vbCrLf
                Else
                    Response.Write "    <td><b>��</b>" & theFiles(i)(0) & "</td>" & vbCrLf
                    Response.Write "    <td align='right'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</td>" & vbCrLf
                    Response.Write "    <td>" & theFiles(i)(2) & "</td>" & vbCrLf
                    Response.Write "   <td class='tdinter'><b><font color='" & interHtml & "</font></b></td>" & vbCrLf
                    If interHtml = "gray'>��" Then
                        Response.Write "    <td><b>��</b>" & FileName & "</td>" & vbCrLf
                        Response.Write "    <td align='right'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</td>" & vbCrLf
                    Else
                        Response.Write "    <td><font color='red'><b>��</b>" & FileName & "</font></td>" & vbCrLf
                        Response.Write "    <td align='right'><font color='red'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "</font>&nbsp;&nbsp;</td>" & vbCrLf
                    End If
                    Response.Write "    <td><font color='red'>" & FileDate & "</font></td>" & vbCrLf
                End If
            End If
 
        Else
            Response.Write "    <td><font color='blue'><b>��</b>" & theFiles(i)(0) & "</font></td>" & vbCrLf
            Response.Write "    <td align='right'><font color='blue'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</font></td>" & vbCrLf
            Response.Write "    <td><font color='blue'>" & theFiles(i)(2) & "</font></td>" & vbCrLf
            Response.Write "    <td class='tdinter'>&nbsp;</td>" & vbCrLf
            Response.Write "    <td></td>" & vbCrLf
            Response.Write "    <td></td>" & vbCrLf
            Response.Write "    <td></td>" & vbCrLf
            bczNum = bczNum + 1
        End If

        Response.Write "</tr>" & vbCrLf
        j = j + 1
    Next
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%'>" & vbCrLf
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='5'><b>�ٷ��ͱ�վ�ȽϽ��ͳ�ƣ�</b></td>" & vbCrLf
    Response.Write "</tr><tr>" & vbCrLf
    Response.Write "    <td>���ߴ�Сʱ����ȫ��ͬ��<font color='red'>" & dyNum & "</font> ��</td>" & vbCrLf
    Response.Write "    <td>���ߴ�С����ͬ��<font color='green'>" & bdyNum & "</font> ��</td>" & vbCrLf
    Response.Write "    <td>���߽���ʱ�䲻ͬ��<font color='gray'>" & ydyNUm & "</font> ��</td>" & vbCrLf
    Response.Write "</tr><tr>" & vbCrLf
    Response.Write "    <td>���ز����ڵ��ļ���<font color='blue'>" & bczNum & "</font> ��</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf


End Sub

Sub ChangeArr(strChannel, ModuleType)
    '����strChannel���͵�Ƶ����ʼ
    Dim arrstrChannel, arrstrChannelNum, arrstrChannelAll, arrstrChannelAllNum
    Dim sqlstrChannel, rsstrChannel, arrResult
    arrstrChannelNum = -1
    arrstrChannelAllNum = -1
    ReDim arrstrChannel(20)
    For i = 0 To FilesNum - 1
        If InStr(theFiles(i)(0), strChannel & "/") > 0 Then
            arrstrChannelNum = arrstrChannelNum + 1
            If arrstrChannelNum > UBound(arrstrChannel) Then
                ReDim Preserve arrstrChannel(arrstrChannelNum + 20)
            End If
            arrstrChannel(arrstrChannelNum) = theFiles(i)
        End If
    Next
    ReDim Preserve arrstrChannel(arrstrChannelNum)

    sqlstrChannel = "select ChannelDir from PE_Channel where ModuleType=" & ModuleType & " and ChannelType=1 and ChannelDir<>'" & strChannel & "' order by ChannelID asc"

    Set rsstrChannel = Server.CreateObject("ADODB.Recordset")
    rsstrChannel.Open sqlstrChannel, Conn, 1, 3
    ReDim arrstrChannelAll((arrstrChannelNum + 1) * rsstrChannel.RecordCount + 20)
    Do While Not rsstrChannel.EOF
        For i = 0 To arrstrChannelNum
            arrstrChannelAllNum = arrstrChannelAllNum + 1
            arrResult = Replace(arrstrChannel(i)(0), strChannel & "/", rsstrChannel("ChannelDir") & "/")
            arrstrChannelAll(arrstrChannelAllNum) = Array(arrResult, arrstrChannel(i)(1), arrstrChannel(i)(2))
        Next
        rsstrChannel.MoveNext
    Loop
    rsstrChannel.Close
    Set rsstrChannel = Nothing
    If arrstrChannelAllNum > -1 Then
        ReDim Preserve arrstrChannelAll(arrstrChannelAllNum)
        ReDim Preserve theFiles(FilesNum + arrstrChannelAllNum)
        For i = FilesNum To FilesNum + arrstrChannelAllNum
            theFiles(i) = arrstrChannelAll(i - FilesNum)
        Next
        FilesNum = FilesNum + arrstrChannelAllNum + 1
    End If
    '����strChannel���͵�Ƶ�����
End Sub

Sub ShowOnlyDif()
    Dim Html, GetFiles, FileInfo
    Dim f, fPath, FileSize, FileDate, theFilePath, FileName, interHtml, trHtml

    Html = GetHttpPage(FileInfoURL, 0)
    If Html = "" Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>��ȡ�ٷ�����ʧ�ܣ����������ķ�������֧�� XMLHTTP ���������ͨ������������������硣</font></p>"
        Exit Sub
    End If

    GetFiles = Split(Html, vbCrLf)
    FilesNum = UBound(GetFiles)
    ReDim theFiles(FilesNum - 1)
    For i = 0 To FilesNum - 1
        FileInfo = Split(GetFiles(i), "|")
        theFiles(i) = FileInfo

    Next
    '����strChannel���͵�Ƶ����ʼ
    Call ChangeArr("Article", 1)
    Call ChangeArr("Soft", 2)
    Call ChangeArr("Photo", 3)
    '����strChannel���͵�Ƶ�����
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class='border'>" & vbCrLf
    Response.Write "<tr class='title0'>" & vbCrLf
    Response.Write "    <td>&nbsp;����(�ٷ�)</td>" & vbCrLf
    Response.Write "    <td>&nbsp;��С</td>" & vbCrLf
    Response.Write "    <td>&nbsp;&nbsp;�޸�ʱ��</td>" & vbCrLf
    Response.Write "    <td class='tdtop'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "    <td>&nbsp;����(��վ)</td>" & vbCrLf
    Response.Write "    <td>&nbsp;��С</td>" & vbCrLf
    Response.Write "    <td>&nbsp;&nbsp;�޸�ʱ��</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
    Dim j
    j = 1
    For i = 0 To FilesNum - 1
        theFilePath = Replace(InstallDir & theFiles(i)(0), "//", "/")
        fPath = Server.MapPath(theFilePath)
        If j Mod 2 = 0 Then
            trHtml = "<tr class='tdbg1' onmouseout=""this.className='tdbg1'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
        Else
            trHtml = "<tr onmouseout=""this.className='tdbgmouseout1'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
        End If
        If fso.FileExists(fPath) Then
            Set f = fso.GetFile(fPath)
            FileName = theFiles(i)(0)
            FileSize = f.size
            FileDate = f.DateLastModified
            If theFiles(i)(1) <> CStr(FileSize) Then
                interHtml = "red'>��"
            Else
                interHtml = "gray'>��"
            End If
 
            If theFiles(i)(1) = CStr(FileSize) And CDate(theFiles(i)(2)) = FileDate Then
                j = j - 1
            Else
                If CDate(theFiles(i)(2)) > FileDate Then
                    Response.Write trHtml & vbCrLf
                    Response.Write "    <td><font color='red'><b>��</b>" & theFiles(i)(0) & "</font></td>" & vbCrLf
                    Response.Write "    <td align='right'><font color='red'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "</font>&nbsp;&nbsp;</td>" & vbCrLf
                    Response.Write "    <td><font color='red'>" & theFiles(i)(2) & "</font></td>" & vbCrLf
                    Response.Write "   <td class='tdinter'><b><font color='" & interHtml & "</font></b></td>" & vbCrLf
                    Response.Write "    <td><font color='gray'><b>��</b>" & FileName & "</font></td>" & vbCrLf
                    Response.Write "    <td align='right'><font color='gray'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "</font>&nbsp;&nbsp;</td>" & vbCrLf
                    Response.Write "    <td><font color='gray'>" & FileDate & "</font></td>" & vbCrLf
                    Response.Write "</tr>" & vbCrLf
                Else
                    Response.Write trHtml & vbCrLf
                    Response.Write "    <td><b>��</b>" & theFiles(i)(0) & "</td>" & vbCrLf
                    Response.Write "    <td align='right'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</td>" & vbCrLf
                    Response.Write "    <td>" & theFiles(i)(2) & "</td>" & vbCrLf
                    Response.Write "   <td class='tdinter'><b><font color='" & interHtml & "</font></b></td>" & vbCrLf
                    If interHtml = "gray'>��" Then
                        Response.Write "    <td><b>��</b>" & FileName & "</td>" & vbCrLf
                        Response.Write "    <td align='right'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</td>" & vbCrLf
                    Else
                        Response.Write "    <td><font color='red'><b>��</b>" & FileName & "</font></td>" & vbCrLf
                        Response.Write "    <td align='right'><font color='red'>" & FormatNumber(FileSize, 0, vbTrue, vbFalse, vbTrue) & "</font>&nbsp;&nbsp;</td>" & vbCrLf
                    End If
                    Response.Write "    <td><font color='red'>" & FileDate & "</font></td>" & vbCrLf
                    Response.Write "</tr>" & vbCrLf
                End If
            End If
 
        Else
            Response.Write trHtml & vbCrLf
            Response.Write "    <td><font color='blue'><b>��</b>" & theFiles(i)(0) & "</font></td>" & vbCrLf
            Response.Write "    <td align='right'><font color='blue'>" & FormatNumber(theFiles(i)(1), 0, vbTrue, vbFalse, vbTrue) & "&nbsp;&nbsp;</font></td>" & vbCrLf
            Response.Write "    <td><font color='blue'>" & theFiles(i)(2) & "</font></td>" & vbCrLf
            Response.Write "    <td class='tdinter'>&nbsp;</td>" & vbCrLf
            Response.Write "    <td></td>" & vbCrLf
            Response.Write "    <td></td>" & vbCrLf
            Response.Write "    <td></td>" & vbCrLf
            Response.Write "</tr>" & vbCrLf
        End If
        j = j + 1
    Next
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf


End Sub
%>
