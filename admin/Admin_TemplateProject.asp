<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Template"   '����Ȩ��

strFileName = "Admin_TemplateProject.asp?Action=" & Action

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>��վģ�巽������</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

If Action <> "TemplateProject" Then
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
    Call ShowPageTitle("��վģ�巽������", 10005)
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""70"" height=""30""><strong>��������</strong></td>" & vbCrLf
    Response.Write "    <td height=""30""><a href=Admin_TemplateProject.asp?Action=Main>������ҳ</a> | <a href=""Admin_TemplateProject.asp?Action=AddProject"">�����ģ�巽����Ŀ</a> | <a href=""Admin_TemplateProject.asp?Action=Import"">����ģ�巽��</a> | <a href=""Admin_TemplateProject.asp?Action=Export"">����ģ�巽��</a> | <a href=""Admin_TemplateProject.asp?Action=TemplateBatchMove"">������ģ��Ǩ�� </a> | <a href=""Admin_TemplateProject.asp?Action=SkinBatchMove"">��������Ǩ��</a> | </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
End If

Select Case Action
Case "AddProject", "ModifyProject"
    Call AddProject
Case "SaveAdd", "SaveModify"
    Call SaveProject
Case "Del"
    Call DelTemplateProject
Case "Del2"
    Call DelTemplateProject2
Case "Set"
    Call SetDefault
Case "Import"                   '��Ŀ�����һ��
    Call Import
Case "Import2"                  '��Ŀ����ڶ���
    Call Import2
Case "DoImport"                 '������Ŀ����
    Call DoImport
Case "Export"                   '��������
    Call Export
Case "DoExport"                 '������������
    Call DoExport
Case "TemplateBatchMove"                'ģ������Ǩ��
    Call TemplateBatchMove
Case "DoTemplateBatchMove"              'ģ������Ǩ�ƴ���
    Call DoTemplateBatchMove
Case "SkinBatchMove"                    '�������Ǩ��
    Call SkinBatchMove
Case "DoSkinBatchMove"                  '�������Ǩ�ƴ���
    Call DoSkinBatchMove
Case "TemplateProject"
    Call TemplateProject
Case Else
        Call main
End Select
Response.Write "</body></html>"
Call CloseConn


'=================================================
'��������main
'��  �ã�������Ŀ
'=================================================
Sub main()
    Dim rs, sql, sysIsDefault

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_TemplateProject.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "  <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "      <td width='50' align='center'><strong>ѡ��</strong></td>"
    Response.Write "      <td align='center' width='80'><strong>��������</strong></td>"
    Response.Write "      <td align='center' width='200'><strong>�������</strong></td>"
    Response.Write "      <td width='60' align='center'><strong>�Ƿ�Ĭ��</strong></td>"
    Response.Write "      <td width='240' height='22' align='center'><strong> ��������</strong></td>"
    Response.Write "      <td width='200' height='22' align='center'><strong> ��������</strong></td>"
    Response.Write "  </tr>"

    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' align='center' height='50'><br>��û��ģ�巽����<br><br></td></tr>"
    Else

        Do While Not rs.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "      <td width='50' align='center' height=""30"">" & rs("TemplateProjectID") & "</td>"
            Response.Write "      <td align='center' width='80'>" & rs("TemplateProjectName") & "</td>"
            Response.Write "      <td align='center' width='200'>" & rs("Intro") & "</td>"
            Response.Write "      <td width='60' align='center'>"

            If rs("IsDefault") = True Then
                Response.Write "<b>��</b>"
            End If

            Response.Write "</td>"
            Response.Write " <td align='center' width='240'>"
            Response.Write " <a href='Admin_Template.asp?Action=Main&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "' >����÷����µ�ģ��</a>" & vbCrLf
            Response.Write " <a href='Admin_Skin.asp?Action=main&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "' >����÷����µķ��</a>" & vbCrLf
            Response.Write "</td>"

            Response.Write "      <td width='200' align='center'><a href='Admin_TemplateProject.asp?Action=ModifyProject&TemplateProjectID=" & rs("TemplateProjectID") & "'>�޸ķ���</a>&nbsp;&nbsp;"

            If rs("IsDefault") = False Then
                Response.Write "<a href='Admin_TemplateProject.asp?Action=Del&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "' onClick=""return confirm('ȷ��Ҫɾ���˷�����ɾ���˷����󷽰�������ģ��,��� �����ᱻɾ��,���ϸ�ע��!');"">ɾ������</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_TemplateProject.asp?Action=Set&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "'  onClick=""return confirm('��ȷ���÷�����ģ��ͷ����Ĭ��������ô,���û��������ӻ򷽰�Ǩ��!');"">��ΪĬ��</a>"
            Else
                Response.Write "<font color='gray'>ɾ������&nbsp;&nbsp;��ΪĬ��</font>"
            End If

            Response.Write "      </td>"
            Response.Write "    </tr>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing

    Response.Write "</table>"
    Response.Write "</form></tr></table>"
End Sub

'=================================================
'��������AddProject
'��  �ã������Ŀ
'=================================================
Sub AddProject()
        
    '������������ ����д
    Dim rsItem, sql, TemplateProjectID
    Dim SaveType, SaveName
    Dim TemplateProjectName, Intro, IsDefault
    Dim iTemplateType, i, Num
    Dim SkinID

    '������ȡ�� ����д
    TemplateProjectID = PE_CLng(Request("TemplateProjectID"))
    FoundErr = False
    SaveType = "SaveAdd"
    SaveName = " �� �� "

    '�Ƿ����޸�
    If TemplateProjectID > 0 Then
        SaveType = "SaveModify"
        SaveName = " �� �� "
        'ȡ������
        sql = "select TemplateProjectID,TemplateProjectName,Intro,IsDefault from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
        Set rsItem = Server.CreateObject("adodb.recordset")
        rsItem.Open sql, Conn, 1, 1

        If rsItem.EOF Then   'û���ҵ�����Ŀ
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������û���ҵ��÷�����</li>"
        Else
            TemplateProjectID = rsItem("TemplateProjectID")
            TemplateProjectName = rsItem("TemplateProjectName")
            Intro = rsItem("Intro")
            IsDefault = rsItem("IsDefault")
        End If

        rsItem.Close
        Set rsItem = Nothing
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<script language = ""JavaScript"">" & vbCrLf
    Response.Write "    function CheckForm(){" & vbCrLf
    Response.Write "        if (document.myform.TemplateProjectName.value==""""){" & vbCrLf
    Response.Write "            alert(""�������Ʋ���Ϊ�գ�"");" & vbCrLf
    Response.Write "            document.myform.TemplateProjectName.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (document.myform.Intro.value==""""){" & vbCrLf
    Response.Write "            alert(""������鲻��Ϊ�գ�"");" & vbCrLf
    Response.Write "            document.myform.Intro.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<FORM name=myform action='Admin_TemplateProject.asp' method=post>" & vbCrLf
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    Response.Write "      <td height='22' colspan='2'><strong> " & SaveName & " �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "     <td class='tdbg'  valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr class='tdbg'> " & vbCrLf '�ı�
    Response.Write "          <td width='150' class='tdbg5' align='right' ><strong> �������ƣ�&nbsp;</strong></td>" & vbCrLf
    Response.Write "          <td class='tdbg'>" & vbCrLf
    Response.Write "            <input name='TemplateProjectName' type='text' id='TemplateProjectName' size='30' maxlength='30' value='" & TemplateProjectName & "'>" & vbCrLf
    Response.Write "            <font color=red> * </font>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'> " & vbCrLf '�ı���
    Response.Write "          <td width='150' class='tdbg5' align='right'><strong> ������飺&nbsp;</strong></td>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <textarea name='Intro' style='width:450px;height:100px' id='Intro'>" & PE_ConvertBR(Intro) & "</textarea>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "     </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<center>" & vbCrLf
    Response.Write "  <Input id='TemplateProjectID' type='hidden' value=" & TemplateProjectID & " name='TemplateProjectID'>" & vbCrLf
    Response.Write "  <Input id='Action' type='hidden' value='" & SaveType & "' name='Action'>" & vbCrLf
    Response.Write "  <Input type='submit' value=' ȷ �� ' name='Submit' onClick=""return CheckForm();"">&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "  <Input type='Reset' name='Reset' value=' �� �� '>" & vbCrLf
    Response.Write "</center>" & vbCrLf
    Response.Write "</FORM>" & vbCrLf

End Sub

'=================================================
'��������Save
'��  �ã�������Ŀ
'=================================================
Sub SaveProject()
    '����������
    Dim TemplateProjectName, Intro, SaveName, TemplateProjectID
    Dim rsItem, rsModify, mrs, sql

    '������ȡ��
    TemplateProjectID = PE_CLng(Request("TemplateProjectID"))
    TemplateProjectName = Replace(ReplaceBadChar(ReplaceText(Trim(Request("TemplateProjectName")), 2)), "nbsp", "")
    Intro = Trim(Request("Intro"))

    '���������
    If TemplateProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ⲻ��Ϊ�գ�</li>"
    End If

    If Len(TemplateProjectName) > 250 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������������ӦС��250����</li>"
    End If

    If Intro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������鲻��Ϊ�գ�</li>"
    End If
        
    sql = "Select TemplateProjectName From PE_TemplateProject Where TemplateProjectName='" & TemplateProjectName & "'"
    Set rsItem = Server.CreateObject("Adodb.Recordset")
    rsItem.Open sql, Conn, 1, 1

    If rsItem.EOF And rsItem.BOF Then
    Else

        If Action = "SaveModify" Then
            sql = "select * from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
            Set rsModify = Server.CreateObject("Adodb.Recordset")
            rsModify.Open sql, Conn, 1, 3

            If rsModify.BOF And rsModify.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ���ָ���ķ�����</li>"
            Else

                If TemplateProjectName <> rsModify("TemplateProjectName") Then
                    FoundErr = True
                End If
            End If

            rsModify.Close
            Set rsModify = Nothing
        Else
            FoundErr = True
        End If

        ErrMsg = ErrMsg & "<li>�����������Ѿ�����Ӧ�ķ�������,�뷵�������������ƣ�</li>"
    End If

    rsItem.Close
    Set rsItem = Nothing

    '���������Ҫ��д�߼�����
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    TemplateProjectName = PE_HTMLEncode(TemplateProjectName)
    Intro = PE_HTMLEncode(Intro)
        
    If FoundErr <> True Then
        '���ݴ洢��
        Set rsItem = Server.CreateObject("adodb.recordset")

        If Action = "SaveAdd" Then
            SaveName = "���"
            Set mrs = Conn.Execute("select max(TemplateProjectID) from PE_TemplateProject")

            If IsNull(mrs(0)) Then
                TemplateProjectID = 1
            Else
                TemplateProjectID = mrs(0) + 1
            End If

            Set mrs = Nothing
            sql = "select top 1 * from PE_TemplateProject"
            rsItem.Open sql, Conn, 1, 3
            rsItem.addnew
            rsItem("TemplateProjectID") = TemplateProjectID
        ElseIf Action = "SaveModify" Then
            SaveName = "�޸�"

            If TemplateProjectID = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����ȷ��������ID!</li>"
                Exit Sub
            Else
                sql = "select * from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
                rsItem.Open sql, Conn, 1, 3

                If rsItem.BOF And rsItem.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Ҳ���ָ���ķ�����</li>"
                    rsItem.Close
                    Set rsItem = Nothing
                    Exit Sub
                End If
            End If
        End If

        '����ģ��,���
        Conn.Execute ("update PE_Skin set ProjectName='" & TemplateProjectName & "' where ProjectName='" & rsItem("TemplateProjectName") & "'")
        Conn.Execute ("update PE_Template set ProjectName='" & TemplateProjectName & "' where ProjectName='" & rsItem("TemplateProjectName") & "'")

        rsItem("TemplateProjectName") = TemplateProjectName
        rsItem("Intro") = Intro
        rsItem.Update
        rsItem.Close
        Set rsItem = Nothing
    Else
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Call WriteSuccessMsg("<Li>" & SaveName & "�����ɹ���", "Admin_TemplateProject.asp?Action=Main")
    Call CloseConn

End Sub

'=================================================
'��������Import
'��  �ã�������Ŀ��һ��
'=================================================
Sub Import()

    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' action='Admin_TemplateProject.asp' method='post' >"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>��վ�������루��һ����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ����ķ������ݿ���ļ�����"
    Response.Write "        <input name='ItemMdb' type='text' id='ItemMdb' value='../temp/PE_TemplateProject.mdb' size='50' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'> </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������Import2
'��  �ã�����ģ�巽���ڶ���
'=================================================
Sub Import2()
    On Error Resume Next
    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    mdbname = Replace(Trim(Request.Form("ItemMdb")), "'", "")

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�������ݿ���"
    End If

    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='post' action='Admin_TemplateProject.asp?action=DoImport'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>��վ�������루�ڶ�����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>��������ķ�����Ŀ</strong><br>"
    Response.Write "<select name='TemplateProjectID' size='2' multiple style='height:300px;width:250px;'>"
    sql = "select * from PE_TemplateProject"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectID") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</select></td>"
    Response.Write "            <td width='80'><input type='submit' name='Submit' value='����&gt;&gt;' "

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "></td>"
    Response.Write "            <td><strong>ϵͳ���Ѿ����ڵķ�����Ŀ</strong><br>"
    Response.Write "             <select name='tItemID' size='2' multiple style='height:300px;width:250px;' disabled>"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectID") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "              </select></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "            <br><b>��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b><br>"
    Response.Write "        <input name='mdbname' type='hidden' id='mdbname' value='" & mdbname & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "        <br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������DoImport
'��  �ã�����ģ�巽����Ŀ����
'=================================================
Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, rs, trs, mrs
    Dim rsTemplate, trsTemplate, rsSkin, trsSkin, rsLabel, trsLabel
    Dim TemplateProjectID
    
    TemplateProjectID = Trim(Request("TemplateProjectID"))
    If IsValidID(TemplateProjectID) = False Then
        TemplateProjectID = ""
    End If

    '��õ���ģ�����ݿ�·��
    mdbname = Replace(Trim(Request.Form("mdbname")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("mdbname")), "'", "")
    End If

    mdbname = Replace(mdbname, "��", "/") '��ֹ�ⲿ���Ӱ�ȫ����

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
        Exit Sub
    End If

    If TemplateProjectID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ��������վ����ID!</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If

    '��������
    Set rs = tconn.Execute("select * from PE_TemplateProject where TemplateProjectID in (" & TemplateProjectID & ")  order by TemplateProjectID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_TemplateProject", Conn, 1, 3

    Do While Not rs.EOF

        If PE_CLng(Conn.Execute("select count(*) from PE_TemplateProject where TemplateProjectName='" & rs("TemplateProjectName") & "'")(0)) > 0 Then
            ErrMsg = ErrMsg & "<li><font color=red >" & rs("TemplateProjectName") & "</font>ϵͳ���Ѿ�����ͬ�ķ���û�е���!</li>"
        Else
            Set mrs = Conn.Execute("select max(TemplateProjectID) from PE_TemplateProject")

            If IsNull(mrs(0)) Then
                TemplateProjectID = 1
            Else
                TemplateProjectID = mrs(0) + 1
            End If

            Set mrs = Nothing

            trs.addnew
            trs("TemplateProjectID") = TemplateProjectID
            trs("TemplateProjectName") = rs("TemplateProjectName")
            trs("Intro") = rs("Intro")
            trs("IsDefault") = False
            'ģ��������������
            Set rsTemplate = tconn.Execute("select * from PE_Template where ProjectName='" & rs("TemplateProjectName") & "' order by TemplateID")
            Set trsTemplate = Server.CreateObject("adodb.recordset")
            trsTemplate.Open "select * from PE_Template", Conn, 1, 3

            If rsTemplate.BOF Or rsTemplate.EOF Then
            Else

                Do While Not rsTemplate.EOF
                    trsTemplate.addnew
                    trsTemplate("ChannelID") = rsTemplate("ChannelID")
                    trsTemplate("TemplateName") = rsTemplate("TemplateName")
                    trsTemplate("TemplateType") = rsTemplate("TemplateType")
                    trsTemplate("TemplateContent") = rsTemplate("TemplateContent")
                    trsTemplate("IsDefault") = False
                    trsTemplate("ProjectName") = rsTemplate("ProjectName")
                    trsTemplate("IsDefaultInProject") = rsTemplate("IsDefaultInProject")
                    trsTemplate("Deleted") = rsTemplate("Deleted")
                    trsTemplate.Update
                    rsTemplate.MoveNext
                Loop

            End If

            trsTemplate.Close
            Set trsTemplate = Nothing
            rsTemplate.Close
            Set rsTemplate = Nothing
            '���������������
            Set rsSkin = tconn.Execute("select * from PE_Skin where ProjectName='" & rs("TemplateProjectName") & "' order by SkinID")
            Set trsSkin = Server.CreateObject("adodb.recordset")
            trsSkin.Open "select * from PE_Skin", Conn, 1, 3

            If rsSkin.BOF Or rsSkin.EOF Then
            Else

                Do While Not rsSkin.EOF
                    trsSkin.addnew
                    trsSkin("SkinName") = rsSkin("SkinName")
                    trsSkin("IsDefault") = False
                    trsSkin("Skin_CSS") = rsSkin("Skin_CSS")
                    trsSkin("IsDefaultInProject") = rsSkin("IsDefaultInProject")
                    trsSkin("ProjectName") = rsSkin("ProjectName")
                    trsSkin.Update
                    rsSkin.MoveNext
                Loop

            End If

            trsSkin.Close
            Set trsSkin = Nothing
            rsSkin.Close
            Set rsSkin = Nothing
            ErrMsg = ErrMsg & "<li><font color=blue >" & rs("TemplateProjectName") & "</font>��������ɹ�!</li>"
            trs.Update
        End If

        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    '�Զ����ǩ����
    Set trsLabel = tconn.Execute("select * from PE_Label")
    Set rsLabel = Server.CreateObject("adodb.recordset")
    rsLabel.Open "select * from PE_Label", Conn, 1, 3

    If Not trsLabel.EOF Then

        Do While Not trsLabel.EOF

            If PE_CLng(Conn.Execute("select count(*) from PE_Label where LabelName='" & trsLabel("LabelName") & "'")(0)) > 0 Then
            Else
                rsLabel.addnew
                rsLabel("LabelName") = trsLabel("LabelName")
                rsLabel("LabelClass") = trsLabel("LabelClass")
                rsLabel("LabelType") = trsLabel("LabelType")
                rsLabel("PageNum") = trsLabel("PageNum")
                rsLabel("reFlashTime") = trsLabel("reFlashTime")
                rsLabel("fieldlist") = trsLabel("fieldlist")
                rsLabel("LabelIntro") = trsLabel("LabelIntro")
                rsLabel("Priority") = trsLabel("Priority")
                rsLabel("LabelContent") = trsLabel("LabelContent")
                rsLabel("AreaCollectionID") = trsLabel("AreaCollectionID")
                rsLabel.Update
            End If

            trsLabel.MoveNext
        Loop

    End If

    Set trsLabel = Nothing
    rsLabel.Close
    Set rsLabel = Nothing
   
    tconn.Close
    Set tconn = Nothing
    Response.Write "<br>"
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>����������ʾ��Ϣ</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>" & ErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td><a href='" & ComeUrl & "'>&lt;&lt; ������һҳ</a></td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Call CreatSkinFile
End Sub

'=================================================
'��������Export
'��  �ã�����ģ�巽����Ŀ
'=================================================
Sub Export()
    Dim rs, sql, iCount
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)
    Response.Write "<br>" & vbCrLf
    Response.Write "<FORM name=myform action='Admin_TemplateProject.asp' method=post>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>��վ��������</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='TemplateProjectID' size='2' multiple style='height:300px;width:450px;'>"

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>��û�з�����Ŀ��</option>"
        '�ر��ύ��ť
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("TemplateProjectID") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "         </select>"
    Response.Write "       </td>"
    Response.Write "       <td align='left'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'>"
    Response.Write "       <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='30'>"
    Response.Write "        <td colspan='2'>Ŀ�����ݿ⣺<input name='Itemmdb' type='text' id='ItemMdb' value='../Temp/PE_TemplateProject.mdb' size='30' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> �����Ŀ�����ݿ�</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='ִ�е�������'>"
    Response.Write "          <input name='Action' type='hidden' id='Action' value='DoExport'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateProjectID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateProjectID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateProjectID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateProjectID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'��������DoExport
'��  �ã�����ģ�巽����Ŀ
'=================================================
Sub DoExport()
    On Error Resume Next
    
    Dim rs, trs, sql, rsLabel, trsLabel, rsTemplate, trsTemplate, rsSkin, trsSkin
    Dim mdbname, tconn
    Dim TemplateProjectID, TemplateProjectName, FormatConn

    FormatConn = Request.Form("FormatConn")
    TemplateProjectID = Trim(Request("TemplateProjectID"))
    mdbname = Replace(Trim(Request.Form("Itemmdb")), "'", "")
    If IsValidID(TemplateProjectID) = False Then
        TemplateProjectID = ""
    End If
    
    If TemplateProjectID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ��������վ����ID!</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�������ݿ���"
    End If

    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        Set tconn = Nothing
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    tconn.Execute ("select TemplateProjectID from PE_TemplateProject")

    If Err Then
        Set trs = Nothing
        ErrMsg = ErrMsg & "<li>��Ҫ���������ݿ�,����ϵͳ�������ݿ�,��ʹ��ϵͳ�������ݿ⡣"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If FormatConn <> "" Then 'Ҫɾ��������
        tconn.Execute ("delete from PE_Label")
        tconn.Execute ("delete from PE_Skin")
        tconn.Execute ("delete from PE_Template")
        tconn.Execute ("delete from PE_TemplateProject")
    End If

    '��������
    Set rs = Conn.Execute("select * from PE_TemplateProject where TemplateProjectID in (" & TemplateProjectID & ")  order by TemplateProjectID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_TemplateProject", tconn, 1, 3

    Do While Not rs.EOF
        trs.addnew
        trs("TemplateProjectID") = rs("TemplateProjectID")
        trs("TemplateProjectName") = rs("TemplateProjectName")
        trs("Intro") = rs("Intro")
        trs("IsDefault") = rs("IsDefault")
        'ģ��������������
        Set rsTemplate = Conn.Execute("select * from PE_Template where ProjectName='" & rs("TemplateProjectName") & "' order by TemplateID")
        Set trsTemplate = Server.CreateObject("adodb.recordset")
        trsTemplate.Open "select * from PE_Template", tconn, 1, 3

        If rsTemplate.BOF Or rsTemplate.EOF Then
        Else

            Do While Not rsTemplate.EOF
                trsTemplate.addnew
                trsTemplate("TemplateID") = rsTemplate("TemplateID")
                trsTemplate("ChannelID") = rsTemplate("ChannelID")
                trsTemplate("TemplateName") = rsTemplate("TemplateName")
                trsTemplate("TemplateType") = rsTemplate("TemplateType")
                trsTemplate("TemplateContent") = rsTemplate("TemplateContent")
                trsTemplate("IsDefault") = rsTemplate("IsDefault")
                trsTemplate("ProjectName") = rsTemplate("ProjectName")
                trsTemplate("IsDefaultInProject") = rsTemplate("IsDefaultInProject")
                trsTemplate("Deleted") = rsTemplate("Deleted")
                trsTemplate.Update
                rsTemplate.MoveNext
            Loop

        End If

        trsTemplate.Close
        Set trsTemplate = Nothing
        rsTemplate.Close
        Set rsTemplate = Nothing
        '���������������
        Set rsSkin = Conn.Execute("select * from PE_Skin where ProjectName='" & rs("TemplateProjectName") & "' order by SkinID")
        Set trsSkin = Server.CreateObject("adodb.recordset")
        trsSkin.Open "select * from PE_Skin", tconn, 1, 3

        If rsSkin.BOF Or rsSkin.EOF Then
        Else

            Do While Not rsSkin.EOF
                trsSkin.addnew
                trsSkin("SkinID") = rsSkin("SkinID")
                trsSkin("SkinName") = rsSkin("SkinName")
                trsSkin("IsDefault") = rsSkin("IsDefault")
                trsSkin("Skin_CSS") = rsSkin("Skin_CSS")
                trsSkin("IsDefaultInProject") = rsSkin("IsDefaultInProject")
                trsSkin("ProjectName") = rsSkin("ProjectName")
                trsSkin.Update
                rsSkin.MoveNext
            Loop

        End If

        trsSkin.Close
        Set trsSkin = Nothing
        rsSkin.Close
        Set rsSkin = Nothing

        trs.Update
        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    '�Զ����ǩ����
    Set trsLabel = Conn.Execute("select * from PE_Label")
    Set rsLabel = Server.CreateObject("adodb.recordset")
    rsLabel.Open "select * from PE_Label", tconn, 1, 3

    If Not trsLabel.EOF Then

        Do While Not trsLabel.EOF
            rsLabel.addnew
            rsLabel("LabelName") = trsLabel("LabelName")
            rsLabel("LabelClass") = trsLabel("LabelClass")
            rsLabel("LabelType") = trsLabel("LabelType")
            rsLabel("PageNum") = trsLabel("PageNum")
            rsLabel("reFlashTime") = trsLabel("reFlashTime")
            rsLabel("fieldlist") = trsLabel("fieldlist")
            rsLabel("LabelIntro") = trsLabel("LabelIntro")
            rsLabel("Priority") = trsLabel("Priority")
            rsLabel("LabelContent") = trsLabel("LabelContent")
            rsLabel("AreaCollectionID") = trsLabel("AreaCollectionID")
            rsLabel.Update
            trsLabel.MoveNext
        Loop

    End If

    Set trsLabel = Nothing
    rsLabel.Close
    Set rsLabel = Nothing
   
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ�����ѡ�еķ���������ָ�������ݿ��У�", ComeUrl)
End Sub

'*************************  ��ģ�����������  *******************************
'*************************  ��ģ����չ��ʼ  *******************************
'=================================================
'��������SetDefault
'��  �ã����÷���Ĭ��
'=================================================
Sub SetDefault()
    Dim TemplateProjectID, ProjectName
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If TemplateProjectID = 0 Then
        FoundErr = True
        ErrMsg = "<li>����ID ����Ϊ��!</li>"
        Exit Sub
    End If

    '������ϵͳĬ��
    Conn.Execute ("update PE_Skin set IsDefault=" & PE_False & " where IsDefault=" & PE_True & "")
    Conn.Execute ("update PE_Skin set IsDefault=" & PE_True & " where IsDefaultInProject=" & PE_True & " and ProjectName='" & ProjectName & "'")
    '����ģ��ϵͳĬ��
    Conn.Execute ("update PE_Template set IsDefault=" & PE_False & " where IsDefault=" & PE_True & "")
    Conn.Execute ("update PE_Template set IsDefault=" & PE_True & " where IsDefaultInProject=" & PE_True & " and ProjectName='" & ProjectName & "'")
    '���巽��ϵͳĬ��
    Conn.Execute ("update PE_TemplateProject set IsDefault=" & PE_False & " where IsDefault=" & PE_True & "")
    Conn.Execute ("update PE_TemplateProject set IsDefault=" & PE_True & " where TemplateProjectName='" & ProjectName & "'")

    Call WriteSuccessMsg("<li>�ɹ���ѡ���ķ�������Ϊ����Ĭ�Ϸ���</li><li>�ɹ���ѡ���ķ������Ϊ����Ĭ�Ϸ��</li><li>�ɹ���ѡ����ģ������Ϊ����Ĭ��ģ��</li>", ComeUrl)
    Call CreatSkinFile
    Call ClearSiteCache(0)
End Sub

'=================================================
'��������DelTemplateProject
'��  �ã�ȷ��ɾ������
'=================================================
Sub DelTemplateProject()
    Dim TemplateProjectID, ProjectName, strTemp
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
    Response.Write "        <br>" & vbCrLf
    Response.Write "        <table border='0' align='center' cellpadding='0' cellspacing='1' width='350' height='150' class='border'>" & vbCrLf
    Response.Write "          <tr class='title' height='22'>" & vbCrLf
    Response.Write "           <td align='center' ><strong>��ȷ��ɾ������ô</strong></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td  align='center' class='tdbg' valign='top'>"
    Response.Write "           <br><br>&nbsp;&nbsp;ȷ��Ҫ<FONT color='red'>ɾ���˷�����</font>ɾ���˷����󷽰�������<FONT color='blue'>ģ��,���</font> �����ᱻɾ��,�����ע��!<br><br><br>"
    Response.Write "                <FONT color='red'> <a href='Admin_TemplateProject.asp?action=Del2&TemplateProjectID=" & TemplateProjectID & "&ProjectName=" & ProjectName & "'>ȷ��ɾ��</a></FONT>&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "                <FONT color='blue'> <a href='Admin_TemplateProject.asp?action=main'> �� �� </a></FONT> " & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
        
End Sub

'=================================================
'��������DelTemplateProject2
'��  �ã�����ɾ������
'=================================================
Sub DelTemplateProject2()
    Dim rs, sql
    Dim TemplateProjectID, ProjectName, strTemp
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If TemplateProjectID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��TemplateProjectID</li>"
        Exit Sub
    End If

    sql = "select * from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���ķ�����</li>"
    Else

        If rs("IsDefault") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ǰ����ΪĬ�Ϸ���������ɾ�������Ƚ�Ĭ�ϸ�Ϊ��������������ɾ���˷�����</li>"
        End If
    End If

    If FoundErr = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.Delete
    rs.Update
    rs.Close
    Set rs = Nothing

    Conn.Execute ("delete from PE_Skin where ProjectName='" & ProjectName & "'")
    Conn.Execute ("delete from PE_Template where ProjectName='" & ProjectName & "'")

    strTemp = strTemp & "<li>�ɹ�ɾ��ѡ���ķ�����</li>"
    strTemp = strTemp & "<li>�ɹ�ɾ��ѡ���ķ����е�����ģ�塣</li>"
    strTemp = strTemp & "<li>�ɹ�ɾ��ѡ���ķ����е����з��</li>"

    Call WriteSuccessMsg(strTemp, "Admin_TemplateProject.asp?Action=main")
End Sub

'=================================================
'��������TemplateProject
'��  �ã�ģ�巽��Ƶ��ѡ��
'=================================================
Sub TemplateProject()

    Dim sql, rs
    Dim iTemplateType, iChannelID, i, Num
    iChannelID = 0
    iTemplateType = 0
    i = 0
    Num = 1
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
        
    sql = "select * from PE_Template where Deleted=" & PE_False & " and ChannelID=" & ChannelID & " order by TemplateType,ChannelID"
        
    Set rs = Conn.Execute(sql)
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "     <tr class='title' height='22'>"
    Response.Write "      <td width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "      <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "      <td width='150' align='center'><b>ģ������</b></td>"
    Response.Write "      <td height='22' align='center'><strong>ģ������</strong></td>"
    Response.Write "      <td width='80' align='center'><strong>�Ƿ�Ĭ��</strong></td>"
    Response.Write "     </tr>"
    i = 0

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td width='100%' colspan='6' align='center'> û �� �� �� ģ ��</td></tr>"
    Else

        Do While Not rs.EOF

            If i > 0 And rs("TemplateType") <> iTemplateType Or i > 0 And rs("ChannelID") <> iChannelID Then
                Num = Num + 1
                Response.Write "<tr height='10'><td colspan='6'></td></tr>"
            End If

            iChannelID = rs("ChannelID")
            iTemplateType = rs("TemplateType")
            i = i + 1

            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "  <td width=""30"" align=""center"" height=""30"">" & vbCrLf
            Response.Write "    <input TYPE='radio' value='" & rs("TemplateID") & "' name=""TemplateID" & Num & """"

            If rs("IsDefault") = True Then Response.Write "checked"
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "      <td width='30' align='center'>" & rs("TemplateID") & "</td>"
            Response.Write "      <td width='150' align='center'>" & GetTemplateTypeName(rs("TemplateType"), rs("ChannelID")) & "</td>"
            Response.Write "      <td align='center'><a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&TemplateID=" & rs("TemplateID") & "'>" & rs("TemplateName") & "</a></td>"
            Response.Write "      <td width='80' align='center'><b>"

            If rs("IsDefault") = True Then
                Response.Write "��"
            Else
                Response.Write "��"
            End If

            Response.Write "</td>"
            Response.Write "</tr>"

            rs.MoveNext
        Loop

        Response.Write "<Input TYPE='hidden' Name='Num' value='" & Num & "'>"

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</table>  "
    Response.Write "</form>"
End Sub

'=================================================
'��������CreatSkinFile
'��  �ã���ʾ����������css�ļ�
'=================================================
Sub CreatSkinFile()

    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    If Not fso.FolderExists(Server.MapPath(InstallDir)) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���Ƚ�����վ���ú��ٽ��д��������</li>"
        Exit Sub
    End If

    If Not fso.FolderExists(Server.MapPath(InstallDir & "Skin")) Then
        fso.CreateFolder (Server.MapPath(InstallDir & "Skin"))
    End If

    Dim rsSkin, sqlSkin, hf, strSkin
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Conn.Execute(sqlSkin)

    Do While Not rsSkin.EOF
        strSkin = Replace_CaseInsensitive(rsSkin("Skin_CSS"), "Skin/", InstallDir & "Skin/")
        Call WriteToFile(InstallDir & "Skin/Skin" & rsSkin("SkinID") & ".css", strSkin)
        rsSkin.MoveNext
    Loop

    rsSkin.Close
    sqlSkin = "select * from PE_Skin where IsDefault=" & PE_True & ""
    Set rsSkin = Conn.Execute(sqlSkin)

    If rsSkin.BOF And rsSkin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�㻹û�н�����һ�������ΪĬ�Ϸ��Ŷ����ǵ�һ��Ҫ����һ��ѽ��</li>"
    Else
        strSkin = Replace_CaseInsensitive(rsSkin("Skin_CSS"), "Skin/", InstallDir & "Skin/")
        Call WriteToFile(InstallDir & "Skin/DefaultSkin.css", strSkin)
    End If

    rsSkin.Close
    Set rsSkin = Nothing
End Sub

'=================================================
'��������TemplateBatchMove
'��  �ã�����Ǩ��ģ��
'=================================================
Sub TemplateBatchMove()
    Dim rs, sql
    Dim TemplateID, TemplateProjectID, ProjectName, TemplateChannelID

    TemplateID = ReplaceBadChar(Trim(Request("TemplateID")))
    TemplateChannelID = PE_CLng(Trim(Request("TemplateChannelID")))
    TemplateProjectID = ReplaceBadChar(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If ProjectName = "" Then
        Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")

        If rs.BOF And rs.EOF Then
            Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
            Exit Sub
        Else
            ProjectName = rs("TemplateProjectName")
        End If

        Set rs = Nothing
    End If
    
    Response.Write "<form method=""post"" action=""Admin_TemplateProject.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td  align='center'><b>������ģ��Ǩ�� </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "          <td class='tdbg5' valign='top' width='50%'>" & vbCrLf
    Response.Write "            <table width='100%' border='0' cellpadding='2' cellspacing='1'>" & vbCrLf
    Response.Write "              <tr>" & vbCrLf
    Response.Write "                <td width='80'></td>" & vbCrLf
    Response.Write "                <td>" & vbCrLf
    Response.Write "                                &nbsp;&nbsp;&nbsp;&nbsp;<b>ѡ�񷽰���ҪǨ�Ƶĵ�ģ��</b><br>" & vbCrLf
    Response.Write "            <select name='ProjectName' style='width:150px;'  onChange='document.form1.submit();'>"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "            </select>"
    Response.Write "            <br>"
    sql = "SELECT DISTINCT t.ChannelID, c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID"
    Set rs = Conn.Execute(sql)
    Response.Write "<select name='TemplateChannelID' id='TemplateChannelID' onChange='document.form1.submit();'>"

    If rs.BOF And rs.EOF Then
        Response.Write "<option value="" selected>��û�����Ƶ����</option> "
    Else

        Do While Not rs.EOF
            Response.Write "<option value=" & rs("ChannelID") & " " & OptionValue(rs("ChannelID"), TemplateChannelID) & ">" & rs("ChannelName") & "</option>"
            rs.MoveNext
        Loop

        Response.Write "<option value='0' " & OptionValue(0, TemplateChannelID) & ">ϵͳͨ��ģ��</option> "
        Response.Write "<option value='999999' " & OptionValue(999999, TemplateChannelID) & ">��������ģ��</option> "
    End If

    Response.Write "</select>"
    rs.Close
    Set rs = Nothing
    Response.Write "              <br>"
    sql = "select ChannelID,TemplateID,TemplateName from PE_Template where "

    If TemplateChannelID <> 999999 Then
        If TemplateChannelID > 0 Then
            sql = sql & " ChannelID=" & TemplateChannelID & " and "
        ElseIf TemplateChannelID = 0 Then
            sql = sql & " ChannelID=0 and "
        End If
    End If

    sql = sql & " ProjectName='" & ProjectName & "' and Deleted=" & PE_False
    '��ʾģ��
    Response.Write "              <select name='BatchTemplateID' id='BatchTemplateID' size='2' multiple style='height:250px;width:250px;' >"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        'û��ģ��ʱָ���ر��ύ��ť
        Response.Write "                <option value='0'>�÷�������û���κ�ģ��</option>"
    Else

        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"

    Response.Write "  <br>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br>" & vbCrLf
    Response.Write "  <FONT style='font-size:12px' color=''><b>��ס��Ctrl����Shift�������Զ�ѡ</b></FONT>" & vbCrLf
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "    function SelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.BatchTemplateID.length;i++){" & vbCrLf
    Response.Write "        document.form1.BatchTemplateID.options[i].selected=true;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function UnSelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.BatchTemplateID.length;i++){" & vbCrLf
    Response.Write "        document.form1.BatchTemplateID.options[i].selected=false;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function CheckForm(){" & vbCrLf
    Response.Write "        if (document.form1.BatchTemplateID.value==""""){" & vbCrLf
    Response.Write "            alert(""Ǩ��ģ�岻��Ϊ�գ�"");" & vbCrLf
    Response.Write "            document.form1.BatchTemplateID.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (document.form1.MoveTemplateProjectName.value==""""){" & vbCrLf
    Response.Write "            alert(""Ǩ�Ƶķ�������Ϊ�գ�"");" & vbCrLf
    Response.Write "            document.form1.MoveTemplateProjectName.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (document.form1.ProjectName.value==document.form1.MoveTemplateProjectName.value){" & vbCrLf
    Response.Write "            alert(""����Ǩ�Ʋ����Լ����Լ��ƶ����ƣ�"");" & vbCrLf
    Response.Write "            document.form1.ProjectName.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        document.form1.Action.value='DoTemplateBatchMove';" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "                </td>" & vbCrLf
    Response.Write "              </tr>" & vbCrLf
    Response.Write "            </table>" & vbCrLf
    Response.Write "          </td>" & vbCrLf

    Response.Write "          <td width='80' class='tdbg' align='center'>" & vbCrLf
        
    Response.Write "<Input TYPE='radio' Name='BatchTypeName' value='�ƶ�' > �ƶ��� &gt;&gt;" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<Input TYPE='radio' Name='BatchTypeName' value='����' > ���Ƶ� &gt;&gt;" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "           <input type='submit' name='Submit' value=' ȷ �� ' onClick=""javascript:return CheckForm()"" >" & vbCrLf
    Response.Write "          </td>"
    Response.Write "          <td class='tdbg' align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>ϵͳ���Ѿ����ڵķ�����Ŀ</strong><br>"
    Response.Write "             &nbsp;&nbsp;<select name='MoveTemplateProjectName' size='2'  style='height:300px;width:200px;' >"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "           </select></td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "     </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<center><FONT color='red'> ע:</FONT>�ƶ���ʱ��,<FONT color='#3366FF'>ϵͳĬ��,����Ĭ��</FONT>�ǲ����ƶ��ġ�</center> " & vbCrLf
    Response.Write "<input name=""Action"" type=""hidden"" id=""Action"" value=""TemplateBatchMove"">" & vbCrLf
    Response.Write "</form>" & vbCrLf

End Sub

'=================================================
'��������DoTemplateBatchMove
'��  �ã�����Ǩ��ģ�崦��
'=================================================
Sub DoTemplateBatchMove()

    Dim rs, trs, jrs, sql
    Dim TemplateType, TemplateID, TemplateProjectName, TemplateChannelID, BatchTemplateID
    Dim ProjectName, MoveTemplateProjectName, BatchTypeName
    Dim tempIsDefault, tempIsDefaultInProject, SysDefault '��ʱ����
        
    FoundErr = False
    tempIsDefault = False
    tempIsDefaultInProject = False

    BatchTypeName = Trim(Request.Form("BatchTypeName"))
    TemplateProjectName = ReplaceBadChar(Trim(Request.Form("TemplateProjectName")))
    TemplateChannelID = PE_CLng(Trim(Request.Form("TemplateChannelID")))
    BatchTemplateID = Trim(Request.Form("BatchTemplateID"))
    ProjectName = Trim(Request.Form("ProjectName"))
    MoveTemplateProjectName = ReplaceBadChar(Trim(Request.Form("MoveTemplateProjectName")))
    If IsValidID(BatchTemplateID) = False Then
        BatchTemplateID = ""
    End If
    
    If BatchTypeName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ѡ��ҪǨ�Ƶ�����,���ƶ����Ǹ��ơ�</li>"
    End If

    If BatchTemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ģ��ID��,�뷵������Ҫ" & BatchTypeName & "��ģ��ID</li>"
    End If

    If FoundInArr(MoveTemplateProjectName, ProjectName, ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ͬ�ķ�������" & BatchTypeName & ",�뷵������" & BatchTypeName & "��ͬ�ķ���</li>"
    End If

    TemplateID = BatchTemplateID

    If MoveTemplateProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��Ҫ" & BatchTypeName & "�ķ���</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    '�õ�ϵͳ����Ĭ������
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ѡ���ģ�����Ͳ���</li>"
    Else
        SysDefault = rs("TemplateProjectName")
    End If

    Set rs = Nothing

    If FoundErr = True Then
        Exit Sub
    End If

    sql = "select * from PE_Template where "

    If InStr(TemplateID, ",") > 0 Then
        sql = sql & " TemplateID in (" & TemplateID & ")"
    Else
        sql = sql & " TemplateID=" & TemplateID
    End If

    If BatchTypeName = "�ƶ�" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, Conn, 1, 3

        '��Ҫ�Ӽ���
        Do While Not rs.EOF
            If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
                ErrMsg = ErrMsg & "<li>��" & rs("ProjectName") & "������Ĭ��ģ�岻��<FONT color='red'>�ƶ�</Font>!"
            Else
                rs("IsDefault") = False
                rs("IsDefaultInProject") = False
                rs("ProjectName") = MoveTemplateProjectName
                ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("TemplateName") & "</FONT>ģ��ɹ�" & BatchTypeName & "��" & MoveTemplateProjectName & "����!"
                rs.Update
            End If
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
    Else
        '��Ҫ�Ӽ���
        Set rs = Conn.Execute(sql)
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Template", Conn, 1, 3

        Do While Not rs.EOF
            trs.addnew
            trs("ChannelID") = rs("ChannelID")
            trs("TemplateName") = rs("TemplateName")
            trs("TemplateType") = rs("TemplateType")
            trs("TemplateContent") = rs("TemplateContent")
            
            '��������ظ�
            Set jrs = Conn.Execute("select * from PE_Template where ChannelID=" & trs("ChannelID") & " and ProjectName='" & MoveTemplateProjectName & "' and TemplateType=" & trs("TemplateType"))

            If jrs.BOF And jrs.EOF Then
                tempIsDefault = True
            Else

                Do While Not jrs.EOF

                    If tempIsDefault = False Then
                        If jrs("IsDefault") = True Or trs("ProjectName") = MoveTemplateProjectName Or jrs("ProjectName") <> SysDefault Then
                            tempIsDefault = True
                        End If
                    End If

                    If tempIsDefaultInProject = False Then
                        If jrs("IsDefaultInProject") = True Or trs("ProjectName") = MoveTemplateProjectName Then
                            tempIsDefaultInProject = True
                        End If
                    End If

                    If tempIsDefault = True And tempIsDefaultInProject = True Then
                        Exit Do
                    End If

                    jrs.MoveNext
                Loop

            End If

            Set jrs = Nothing

            If tempIsDefault = True Then
                trs("IsDefault") = False
            Else
                trs("IsDefault") = rs("IsDefault")
            End If

            If tempIsDefaultInProject = True Then
                trs("IsDefaultInProject") = False
            Else
                trs("IsDefaultInProject") = rs("IsDefaultInProject")
            End If

            trs("ProjectName") = MoveTemplateProjectName
            trs("Deleted") = rs("Deleted")
            ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("TemplateName") & "</FONT>ģ��ɹ�" & BatchTypeName & "��" & MoveTemplateProjectName & "����!"
            tempIsDefaultInProject = False
            trs.Update
            rs.MoveNext
        Loop

        trs.Close
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If

    Call WriteSuccessMsg(ErrMsg, "Admin_TemplateProject.asp?action=Main&ProjectName=" & SysDefault)
End Sub

'=================================================
'��������SkinBatchMove
'��  �ã�����Ǩ�Ʒ��
'=================================================
Sub SkinBatchMove()

    Dim rs, sql
    Dim SkinID, ProjectName, BatchTypeName

    SkinID = ReplaceBadChar(Trim(Request("SkinID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
    BatchTypeName = Trim(Request("BatchTypeName"))

    If ProjectName = "" Then
        ProjectName = "���з���"
    End If

    Response.Write "<form method=""post"" action=""Admin_TemplateProject.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td  align='center'><b>��������Ǩ�� </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "          <td class='tdbg5' valign='top' width='50%' >" & vbCrLf
    Response.Write "                <table border='0' cellpadding='0' cellspacing='1' width='100%' height='100%'>" & vbCrLf
    Response.Write "                  <tr class='tdbg'>" & vbCrLf
    Response.Write "                    <td width='100' class='tdbg5'></td>" & vbCrLf
    Response.Write "                    <td align='left' class='tdbg5'>" & vbCrLf
    Response.Write "                     <select name='ProjectName' style='width:150px;' onChange='document.form1.submit();'>"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "                <option value='0'>û���κη�����Ŀ</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

        Response.Write "<option value='���з���' " & OptionValue("���з���", ProjectName) & ">���з���</option>"
    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                       </select>"
    Response.Write "                </td>" & vbCrLf
    Response.Write "                   </tr>" & vbCrLf
    Response.Write "                   <tr class='tdbg'>" & vbCrLf
    Response.Write "                     <td width='100' class='tdbg5'></td>" & vbCrLf
    Response.Write "                     <td  align='left' class='tdbg5'>"
    Response.Write "                    <select name='SkinID'  size='2' multiple style='height:250px;width:250px;'>"
    sql = "select * from PE_Skin"

    If ProjectName <> "���з���" And ProjectName <> "" Then
        sql = sql & " where ProjectName='" & ProjectName & "'"
    End If

    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη��</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("SkinID") & "' "

            If SkinID <> "" Then
                If InStr(SkinID, ",") > 0 Then
                    If FoundInArr(SkinID, rs("SkinID"), ",") = True Then Response.Write "selected"
                Else
                    SkinID = PE_CLng(Trim(SkinID))

                    If rs("SkinID") = SkinID Then Response.Write "selected"
                End If
            End If

            Response.Write ">" & rs("SkinName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"
    Response.Write "  <br>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br>" & vbCrLf
    Response.Write "  <FONT style='font-size:12px' color=''><b>��ס��Ctrl����Shift�������Զ�ѡ</b></FONT>" & vbCrLf
    Response.Write "  <script language='javascript'>" & vbCrLf
    Response.Write "    function SelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.SkinID.length;i++){" & vbCrLf
    Response.Write "        document.form1.SkinID.options[i].selected=true;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function UnSelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.SkinID.length;i++){" & vbCrLf
    Response.Write "        document.form1.SkinID.options[i].selected=false;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  </script>" & vbCrLf
    Response.Write "                  </td>" & vbCrLf
    Response.Write "                     </tr>" & vbCrLf
    Response.Write "                    </table>" & vbCrLf
    Response.Write "               </td>" & vbCrLf
    Response.Write "               <td width='80' class='tdbg' align='center'>" & vbCrLf
    Response.Write "                 <Input TYPE='radio' Name='BatchTypeName' value='�ƶ�' " & IsRadioChecked(BatchTypeName, "�ƶ�") & "   > �ƶ��� &gt;&gt;<br>" & vbCrLf
    Response.Write "                 <Input TYPE='radio' Name='BatchTypeName' value='����' " & IsRadioChecked(BatchTypeName, "����") & " > ���Ƶ� &gt;&gt;<br>" & vbCrLf
    Response.Write "                 <Input type='submit' name='Submit' value=' ȷ �� ' onClick=""document.form1.Action.value='DoSkinBatchMove';"" >" & vbCrLf
    Response.Write "               </td>"
    Response.Write "               <td class='tdbg' align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>ϵͳ���Ѿ����ڵķ�����Ŀ</strong><br>"
    Response.Write "             &nbsp;&nbsp;<select name='MoveTemplateProjectName' size='2'  style='height:300px;width:200px;' >"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "           </select></td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "     </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<center><FONT color='red'> ע:</FONT>�ƶ���ʱ��,<FONT color='#3366FF'>ϵͳĬ��,����Ĭ��</FONT>�ǲ����ƶ��ġ�</center> " & vbCrLf
    Response.Write "<input name=""Action"" type=""hidden"" id=""Action"" value=""SkinBatchMove"">" & vbCrLf
    Response.Write "</form>" & vbCrLf

End Sub

'=================================================
'��������DoSkinBatchMove
'��  �ã���������Ǩ�Ʒ��
'=================================================
Sub DoSkinBatchMove()

    Dim rs, trs, jrs, sql
    Dim SkinID
    Dim MoveTemplateProjectName, BatchTypeName, SysDefault

    Dim tempIsDefault, tempIsDefaultInProject '��ʱ����
        
    FoundErr = False
    tempIsDefault = False
    tempIsDefaultInProject = False

    BatchTypeName = Trim(Request.Form("BatchTypeName"))
    SkinID = Trim(Request.Form("SkinID"))
    MoveTemplateProjectName = ReplaceBadChar(Trim(Request.Form("MoveTemplateProjectName")))
    If IsValidID(SkinID) = False Then
        SkinID = ""
    End If

    If BatchTypeName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ���ƶ���������</li>"
    End If

    If SkinID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��Ҫ" & BatchTypeName & "�ķ��</li>"
    End If

    If MoveTemplateProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��Ҫ" & BatchTypeName & "�ķ���</li>"
    End If

    '�õ�ϵͳ����Ĭ������
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ѡ���ģ�����Ͳ���</li>"
    Else
        SysDefault = rs("TemplateProjectName")
    End If

    Set rs = Nothing

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    sql = "select * from PE_Skin where "

    If InStr(SkinID, ",") > 0 Then
        sql = sql & " SkinID in (" & SkinID & ")"
    Else
        sql = sql & " SkinID=" & SkinID
    End If

    If BatchTypeName = "�ƶ�" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, Conn, 1, 3

        Do While Not rs.EOF
            If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
                ErrMsg = ErrMsg & "<li><FONT color='red'> " & rs("SkinName") & "</FONT>��" & rs("ProjectName") & "������Ĭ�Ϸ�����ƶ�!"
            Else
                rs("IsDefault") = False
                rs("IsDefaultInProject") = False
                rs("ProjectName") = MoveTemplateProjectName
                ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("SkinName") & "</FONT>���ɹ�" & BatchTypeName & "��" & MoveTemplateProjectName & "����!"
                rs.Update
            End If
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
    Else
        Set rs = Conn.Execute(sql)
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Skin", Conn, 1, 3

        Do While Not rs.EOF
            trs.addnew
            trs("SkinName") = rs("SkinName")
            trs("Skin_CSS") = rs("Skin_CSS")
            '��������ظ�
            Set jrs = Conn.Execute("select * from PE_Skin where ProjectName='" & MoveTemplateProjectName & "'")

            If jrs.BOF And jrs.EOF Then
                tempIsDefault = True
            Else

                Do While Not rs.EOF

                    If tempIsDefault = False Then
                        If jrs("IsDefault") = True Or trs("ProjectName") = MoveTemplateProjectName Or jrs("ProjectName") <> SysDefault Then
                            tempIsDefault = True
                        End If
                    End If

                    If tempIsDefaultInProject = False Then
                        If jrs("IsDefaultInProject") = True Or trs("ProjectName") = MoveTemplateProjectName Then
                            tempIsDefaultInProject = True
                        End If
                    End If

                    If tempIsDefault = True And tempIsDefaultInProject = True Then
                        Exit Do
                    End If

                    jrs.MoveNext
                Loop

            End If

            Set jrs = Nothing

            If tempIsDefault = True Then
                trs("IsDefault") = False
            Else
                trs("IsDefault") = rs("IsDefault")
            End If

            If tempIsDefaultInProject = True Then
                trs("IsDefaultInProject") = False
            Else
                trs("IsDefaultInProject") = rs("IsDefaultInProject")
            End If

            trs("ProjectName") = MoveTemplateProjectName
            ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("SkinName") & "</FONT>���ɹ�" & BatchTypeName & "��" & MoveTemplateProjectName & "����!"
            tempIsDefaultInProject = False
            trs.Update
            rs.MoveNext
        Loop

        trs.Close
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If

    Call WriteSuccessMsg(ErrMsg, "Admin_TemplateProject.asp?action=main")
End Sub

'*************************  ��ģ����չ�����  *******************************
'*************************  ��ģ�麯��ͨ�ÿ�ʼ  *****************************
'=================================================
'��������GetTemplateTypeName
'��  �ã���ʾ��ǰƵ����ģ������
'��  ����iTemplateType --- �����ģ��ֵ
'=================================================
Function GetTemplateTypeName(iTemplateType, _
                                     ChannelID)

    If ChannelID > 0 Then
        If ModuleType = 4 Then

            Select Case iTemplateType

                Case 1
                    GetTemplateTypeName = "������ҳģ��"

                Case 3
                    GetTemplateTypeName = "���Է���ģ��"

                Case 4
                    GetTemplateTypeName = "���Իظ�ģ��"

                Case 5
                    GetTemplateTypeName = "��������ҳģ��"
            End Select

        Else

            Select Case iTemplateType

                Case 1
                    GetTemplateTypeName = "Ƶ����ҳģ��"

                Case 2
                    GetTemplateTypeName = "Ƶ����Ŀģ��"

                Case 3
                    GetTemplateTypeName = "Ƶ������ҳģ��"

                Case 4
                    GetTemplateTypeName = "Ƶ��ר��ҳģ��"

                Case 5
                    GetTemplateTypeName = "Ƶ������ҳģ��"

                Case 6
                    GetTemplateTypeName = "����" & ChannelShortName & "ҳģ��"

                Case 7
                    GetTemplateTypeName = "�Ƽ�" & ChannelShortName & "ҳģ��"

                Case 8
                    GetTemplateTypeName = "�ȵ�" & ChannelShortName & "ҳģ��"

                Case 16
                    GetTemplateTypeName = "����" & ChannelShortName & "ҳģ��"

                Case 9
                    GetTemplateTypeName = "���ﳵģ��"

                Case 10
                    GetTemplateTypeName = "����̨ģ��"

                Case 11
                    GetTemplateTypeName = "Ԥ������ģ��"

                Case 12
                    GetTemplateTypeName = "�����ɹ�ҳģ��"

                Case 13
                    GetTemplateTypeName = "����֧����һ��ģ��"

                Case 14
                    GetTemplateTypeName = "����֧���ڶ���ģ��"

                Case 15
                    GetTemplateTypeName = "����֧��������ģ��"

                Case 17
                    GetTemplateTypeName = "��ӡģ��"

                Case 101
                    GetTemplateTypeName = "�Զ����б�ģ��"

                Case 19
                    GetTemplateTypeName = "�ؼ���Ʒҳģ��"

                Case 20
                    GetTemplateTypeName = "���ߺ���ҳģ��"

                Case 21
                    GetTemplateTypeName = "�̳ǰ���ҳģ��"

                Case 22
                    GetTemplateTypeName = "Ƶ��ר���б�ҳģ��"

                Case 23
                    GetTemplateTypeName = "�������" & ChannelShortName & "ҳģ��"
                
            End Select

        End If

    Else

        Select Case iTemplateType

            Case 1
                GetTemplateTypeName = "��վ��ҳģ��"

            Case 3
                GetTemplateTypeName = "��վ����ҳģ��"

            Case 4
                GetTemplateTypeName = "��վ����ҳģ��"

            Case 5
                GetTemplateTypeName = "��������ҳģ��"

            Case 6
                GetTemplateTypeName = "��վ����ҳģ��"

            Case 7
                GetTemplateTypeName = "��Ȩ����ҳģ��"

            Case 8
                GetTemplateTypeName = "��Ա��Ϣҳģ��"

            Case 102
                GetTemplateTypeName = "��Ա����ͨ��ģ��"
				
            Case 9
                GetTemplateTypeName = "��Ա�б�ҳģ��"

            Case 10
                GetTemplateTypeName = "������ʾҳģ��"

            Case 11
                GetTemplateTypeName = "�����б�ҳģ��"

            Case 12
                GetTemplateTypeName = "��Դ��ʾҳģ��"

            Case 13
                GetTemplateTypeName = "��Դ�б�ҳģ��"
				
            Case 103
                GetTemplateTypeName = "����Ͷ��ģ��"				

            Case 14
                GetTemplateTypeName = "������ʾҳģ��"

            Case 15
                GetTemplateTypeName = "�����б�ҳģ��"

            Case 16
                GetTemplateTypeName = "Ʒ����ʾҳģ��"

            Case 17
                GetTemplateTypeName = "Ʒ���б�ҳģ��"

            Case 101
                GetTemplateTypeName = "�Զ����б�ģ��"

            Case 18
                GetTemplateTypeName = "��Աע��ҳģ�壨���Э�飩"

            Case 19
                GetTemplateTypeName = "��Աע��ҳģ�壨������Ŀ��"

            Case 20
                GetTemplateTypeName = "��Աע��ҳģ�壨ѡ����Ŀ��"

            Case 21
                GetTemplateTypeName = "��Աע��ҳģ�壨ע������"

            Case 22
                GetTemplateTypeName = "�����б�ҳģ��"
                'Case 22
                '    GetTemplateTypeName = "��������ҳģ�� (��̨)"
                'Case 23
                '    GetTemplateTypeName = "��������ҳģ�� (��̨)"
                'Case 24
                '    GetTemplateTypeName = "�鿴����ҳģ�� (��̨)"
                'Case 999
                '    GetTemplateTypeName = "ͨ����ʾҳģ��"
        End Select

    End If

    If iTemplateType = 0 Then
        GetTemplateTypeName = "��ǰ��������ģ��"
    End If

End Function

'**************************************************
'��������ReplaceText
'��  �ã����˷Ƿ��ַ���
'��  ����iText-----�����ַ���
'����ֵ���滻���ַ���
'**************************************************
Function ReplaceText(iText, _
                             iType)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol

    If PE_Cache.GetValue("Site_ReplaceText") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText from PE_KeyLink where isUse=1 and LinkType=" & iType
        rsKey.Open sqlKey, Conn, 1, 1

        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_ReplaceText_" & iType, rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceText = iText
            Exit Function
        End If
    End If

    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_ReplaceText_" & iType), "@@@")

    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        rText = Replace(rText, Keycol(0), Keycol(1))
    Next

    ReplaceText = rText
End Function

'**************************************************
'��������IsOptionSelected
'��  �ã������˵�Ĭ�ϱȽ�
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Function IsOptionSelected(ByVal Compare1, _
                                  ByVal Compare2)

    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If

End Function

'**************************************************
'��������IsFontChecked
'��  �ã���ѡ,��ѡĬ��
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Function IsFontChecked(ByVal Compare1, _
                               ByVal Compare2)

    If Compare1 = Compare2 Then
        IsFontChecked = " color='red'"
    Else
        IsFontChecked = ""
    End If

End Function

'**************************************************
'��������IsRadioChecked
'��  �ã���ѡ,��ѡĬ��
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Function IsRadioChecked(ByVal Compare1, _
                                ByVal Compare2)

    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If

End Function
%>
