<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 1      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

FileName = "Admin_City.asp"
strFileName = FileName & "?Field=" & strField & "&keyword=" & Keyword

Response.Write "<html><head><title>��վ�����������</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� �� �� �� �� ��", 10032)
Response.Write "<tr class='tdbg'><td width='70'>��������</td><td>"
Response.Write " <a href='Admin_City.asp'>ȫ����������</a> | "
Response.Write " <a href='Admin_City.asp?Action=PostCodeAdd'>�����������</a>"
Response.Write "</table>" & vbCrLf

Select Case Action
Case "PostCodeAdd"
    Call PostCodeAdd
Case "SavePostCodeAdd"
    Call SavePostCodeAdd
Case "PostCodeDel"
    Call PostCodeDel
Case "PostCodeEdit"
    Call PostCodeEdit
Case "SavePostCodeEdit"
    Call SavePostCodeEdit
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsPostCode, sqlPostCode
    Dim strAction, strSql
    If Keyword <> "" Then
        strAction = "���������"
        Select Case strField
        Case "Province"
            strSql = " and Province like '%" & Keyword & "%' "
            strAction = strAction & "ʡ���к��йؼ���<font color='red'>" & Keyword & "</font>�ļ�¼"
        Case "City"
            strSql = " and City like '%" & Keyword & "%' "
            strAction = strAction & "�����к��йؼ���<font color='red'>" & Keyword & "</font>�ļ�¼"
        Case "Area"
            strSql = " and Area like '%" & Keyword & "%' "
            strAction = strAction & "�����к��йؼ���<font color='red'>" & Keyword & "</font>�ļ�¼"
        Case "PostCode"
            strSql = " and PostCode like '%" & Keyword & "%' "
            strAction = strAction & "���������к��йؼ���<font color='red'>" & Keyword & "</font>�ļ�¼"
        Case "AreaCode"
            strSql = " and AreaCode like '%" & Keyword & "%' "
            strAction = strAction & "�����к��йؼ���<font color='red'>" & Keyword & "</font>�ļ�¼"
        Case Else
            strSql = " and Area like '%" & Keyword & "%' "
            strAction = strAction & "�����к��йؼ���<font color='red'>" & Keyword & "</font>�ļ�¼"
        End Select
    Else
        strAction = "ȫ����������"
    End If
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã������������&nbsp;&gt;&gt;&nbsp;"
    Response.Write strAction
    Response.Write "</td></tr></table>"
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "      <td>"
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' align='center' height='22'>"
    Response.Write "          <td><strong>����ʡ</strong></td>"
    Response.Write "          <td><strong>��������</strong></td>"
    Response.Write "          <td width='150'><strong>����</strong></td>"
    Response.Write "          <td width='140'><strong>��������</strong></td>"
    Response.Write "          <td width='110'><strong>����</strong></td>"
    Response.Write "          <td width='150'><strong>����</strong></td>"
    Response.Write "        </tr>"

    sqlPostCode = "select * from PE_City where 1=1 "

    If Keyword <> "" Then
        sqlPostCode = sqlPostCode & strSql
    End If
    sqlPostCode = sqlPostCode & " order by AreaID asc"
    Set rsPostCode = Server.CreateObject("adodb.recordset")
    rsPostCode.Open sqlPostCode, Conn, 1, 1
    If rsPostCode.EOF Then
        Response.Write "<tr class='tdbg' align='center' height='50'><td colspan='10'>�޴���Ϣ��</td></tr>"
    Else
        totalPut = rsPostCode.RecordCount
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
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsPostCode.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Dim PostCodeNum
        PostCodeNum = 0
        Do While Not rsPostCode.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td>" & rsPostCode("Province") & "</td>"
            Response.Write "        <td>" & rsPostCode("City") & "</td>"
            Response.Write "        <td width='150'>&nbsp;&nbsp;&nbsp;" & rsPostCode("Area") & "</td>"
            Response.Write "        <td width='140' align='center'>" & rsPostCode("PostCode") & "</td>"
            Response.Write "        <td width='110' align='center'>" & rsPostCode("AreaCode") & "</td>"
            Response.Write "        <td width='150' align='center'><a href='Admin_City.asp?Action=PostCodeEdit&AreaID=" & rsPostCode("AreaID") & "'>�༭</a> | <a href='Admin_City.asp?Action=PostCodeDel&AreaID=" & rsPostCode("AreaID") & "' onclick=""return confirm('ȷ��Ҫɾ��������¼��');"">ɾ��</a></td>"
            Response.Write "      </tr>"
            PostCodeNum = PostCodeNum + 1
            If PostCodeNum >= MaxPerPage Then Exit Do
            rsPostCode.MoveNext
        Loop
    End If
    rsPostCode.Close
    Set rsPostCode = Nothing
    
    Response.Write "</table>"

    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����¼", True)

    Response.Write PostCodeSearch
End Sub

Sub PostCodeAdd()
    Response.Write "    <form method='post' action='" & FileName & "' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
    Response.Write "         <tr class='title'>" & vbCrLf
    Response.Write "            <td height='22' colspan='2'> " & vbCrLf
    Response.Write "               <div align='center'><strong>�����������</strong></div>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "          </tr>    " & vbCrLf
    Response.Write "         <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>����ʡ�ݣ�</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Province' type='text' id='Province' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>�������У�</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='City' type='text' id='City' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>����������</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Area' type='text' id='Area' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>�������룺</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='PostCode' type='text' id='PostCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>�������ţ�</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='AreaCode' type='text' id='AreaCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>     " & vbCrLf
    Response.Write "                     <td colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "                     " & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SavePostCodeAdd'>        <input  type='submit' name='Submit' value=' �� �� '  style='cursor:hand;'>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_City.asp'"" style='cursor:hand;'>" & vbCrLf
    Response.Write "                     </td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </form>" & vbCrLf
    Response.Write PostCodeSearch
End Sub


Sub SavePostCodeAdd()
    Dim Province, City, Area, PostCode, AreaCode, sql, rs
    Province = ReplaceBadChar(Trim(Request.Form("Province")))
    City = ReplaceBadChar(Trim(Request.Form("City")))
    Area = ReplaceBadChar(Trim(Request.Form("Area")))
    PostCode = ReplaceBadChar(Trim(Request.Form("PostCode")))
    AreaCode = ReplaceBadChar(Trim(Request.Form("AreaCode")))
    If Province = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����дʡ�ݣ�"
        Exit Sub
    End If
    If City = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д���У�"
        Exit Sub
    End If
    If Area = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д������"
        Exit Sub
    End If
    If PostCode = "" Or Not IsTrueCode(PostCode, "PostCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ���������룡"
        Exit Sub
    End If
    If AreaCode = "" Or Not IsTrueCode(AreaCode, "AreaCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ�����ţ�"
        Exit Sub
    End If
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_City where Province='" & Province & "' and City='" & City & "' and Area='" & Area & "'"
    rs.Open sql, Conn, 1, 3
    If rs.EOF And rs.BOF Then
        rs.AddNew
        rs("Country") = "�л����񹲺͹�"
        rs("Province") = Province
        rs("City") = City
        rs("Area") = Area
        rs("PostCode") = PostCode
        rs("AreaCode") = AreaCode
        rs.Update
        Call WriteSuccessMsg("����������ӳɹ���", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ĵ����Ѿ����ڡ�</li>"
    End If
    rs.Close
    Set rs = Nothing
End Sub


Sub PostCodeEdit()
    Dim AreaID, PostCode, sql, rs

    AreaID = PE_CLng(Trim(Request("AreaID")))

    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_City where AreaID=" & AreaID & ""
    rs.Open sql, Conn, 1, 3
    If rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "�����ڸü�¼��"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    Response.Write "    <form method='post' action='" & FileName & "' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
    Response.Write "         <tr class='title'>" & vbCrLf
    Response.Write "            <td height='22' colspan='2'> " & vbCrLf
    Response.Write "               <div align='center'><strong>�޸���������</strong></div>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>����ʡ�ݣ�</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Province' value='" & rs("Province") & "' type='text' id='Province' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>�������У�</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='City' value='" & rs("City") & "' type='text' id='City' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "               <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>����������</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Area' value='" & rs("Area") & "' type='text' id='Area' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>�������룺</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='PostCode' value='" & rs("PostCode") & "' type='text' id='PostCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>�������ţ�</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='AreaCode' value='" & rs("AreaCode") & "' type='text' id='AreaCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>     " & vbCrLf
    Response.Write "                     <td colspan='2' align='center' class='tdbg'><input name='AreaID' type='hidden' id='AreaID' value='" & AreaID & "'>" & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SavePostCodeEdit'>        <input  type='submit' name='Submit' value='�����޸Ľ��'  style='cursor:hand;'>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_City.asp'"" style='cursor:hand;'>" & vbCrLf
    Response.Write "                     </td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </form>" & vbCrLf
    rs.Close
    Set rs = Nothing
    Response.Write PostCodeSearch
End Sub

Sub SavePostCodeEdit()
    Dim AreaID, Province, City, Area, PostCode, AreaCode, sql, rs
    AreaID = PE_CLng(Trim(Request.Form("AreaID")))
    Province = ReplaceBadChar(Trim(Request.Form("Province")))
    City = ReplaceBadChar(Trim(Request.Form("City")))
    Area = ReplaceBadChar(Trim(Request.Form("Area")))
    PostCode = ReplaceBadChar(Trim(Request.Form("PostCode")))
    AreaCode = ReplaceBadChar(Trim(Request.Form("AreaCode")))
    If Province = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����дʡ�ݣ�"
    End If
    If City = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д���У�"
    End If
    If PostCode = "" Or Not IsTrueCode(PostCode, "PostCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ���������룡"
    End If
    If AreaCode = "" Or Not IsTrueCode(AreaCode, "AreaCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ�����ţ�"
    End If
    If FoundErr = True Then Exit Sub

    Dim trs
    Set trs = Conn.Execute("select top 1 AreaID from PE_City where AreaID<>" & AreaID & " and Province='" & Province & "' and City='" & City & "' and Area='" & Area & "'")
    If Not (trs.BOF And trs.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������������Ѿ����ڣ�</li>"
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub

    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_City where AreaID=" & AreaID & ""
    rs.Open sql, Conn, 1, 3
    If Not (rs.EOF And rs.BOF) Then
        rs("Province") = Province
        rs("City") = City
        rs("Area") = Area
        rs("PostCode") = PostCode
        rs("AreaCode") = AreaCode
        rs.Update
        Call WriteSuccessMsg("���������޸ĳɹ���", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "�޸�ʧ�ܣ�ԭʼ���ݶ�ʧ��"
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub PostCodeDel()
    Dim AreaID, RowCount

    AreaID = PE_CLng(Trim(Request("AreaID")))

    Conn.Execute ("delete from PE_City where AreaID=" & AreaID & ""), RowCount
    If RowCount = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "��¼ɾ��ʧ�ܡ�"
    Else
        Call WriteSuccessMsg("��¼ɾ���ɹ���", ComeUrl)
    End If
End Sub

Function IsTrueCode(thisCode, CodeType)
    Dim temp
    IsTrueCode = False
    If CodeType = "PostCode" Then
        regEx.Pattern = "^\d{6}$"
    Else
        regEx.Pattern = "^\d{3,7}$"
    End If

    IsTrueCode = regEx.Test(thisCode)
End Function

Function PostCodeSearch()
    Dim strHtml
    strHtml = "<br>"
    strHtml = strHtml & "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    strHtml = strHtml & "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    strHtml = strHtml & "  <tr class='tdbg'>"
    strHtml = strHtml & "   <td width='130' align='right'><strong>��������������</strong></td>"
    strHtml = strHtml & "   <td>"
    strHtml = strHtml & "<select name='Field' size='1'>"
    strHtml = strHtml & "<option value='Province'>����ʡ��</option>"
    strHtml = strHtml & "<option value='City'>��������</option>"
    strHtml = strHtml & "<option value='Area' selected>����</option>"
    strHtml = strHtml & "<option value='PostCode'>��������</option>"
    strHtml = strHtml & "<option value='AreaCode'>��������</option>"
    strHtml = strHtml & "</select>"
    strHtml = strHtml & "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    strHtml = strHtml & "<input type='submit' name='Submit'  value='����'>"
    strHtml = strHtml & "</td></tr></table></form>"
    PostCodeSearch = strHtml
End Function
%>
