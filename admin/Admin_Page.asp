<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Class.asp"-->
<!--#include file="../Include/PowerEasy.Special.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<!--#include file="../Include/PowerEasy.Photo.asp"-->
<!--#include file="../Include/PowerEasy.Product.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "ShowPage"   '����Ȩ��

Response.Write "<html><head><title>�Զ���ҳ�����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle("�� �� �� ҳ �� �� ��", 10027)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>��������</strong></td><td>"
Response.Write "<a href='Admin_Page.asp'>�Զ���ҳ�������ҳ</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Page.asp?Action=AddClass'>����Զ������</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Page.asp?Action=AddPage'>����Զ���ҳ��</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Page.asp?Action=import'>�����Զ������</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Page.asp?Action=export'>�����Զ������</a>&nbsp;|&nbsp;"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

strFileName = "Admin_Page.asp"
Select Case Action
Case "AddClass"
    Call AddClass
Case "ModifyClass"
    Call ModifyClass
Case "SaveClass", "SaveModifyClass"
    Call SaveClass
Case "DelClass"
    Call DelClass
Case "ListPage"
    Call ListPage
Case "AddPage"
    Call AddPage
Case "ModifyPage"
    Call ModifyPage
Case "SavePage", "SaveModifyPage"
    Call SavePage
Case "DelPage"
    Call DelPage
Case "CreateFile"
    Call CreateFile("")
Case "CreateClassFile"
    Call CreateClassFile
Case "import"
    Call Import
Case "import2"
    Call import2
Case "Doimport"
    Call DoImport
Case "export"
    Call Export
Case "Doexport"
    Call DoExport
Case Else
    Call main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
If Action = "DelClass" Or Action = "DelPage" Or Action = "CreateFile" Or Action = "CreateClassFile" Then
    Response.Redirect ComeUrl
Else
    Response.Write "</body></html>"
End If
Call CloseConn

Sub main()
    Dim sqlClass, rsClass, ClassName, rsPage, UseAsp
    Dim iCount
    UseAsp = False

    Response.Write "<form name='myform' method='post' action=''>"
    Set rsClass = Server.CreateObject("Adodb.RecordSet")
    sqlClass = "select ID,ClassName,ClassIntro,ClassType from PE_PageClass Order by ID"
    rsClass.Open sqlClass, Conn, 1, 1
    If rsClass.BOF And rsClass.EOF Then
        rsClass.Close
        Set rsClass = Nothing
        Response.Write "<center>��δ��ӷ��࣡</center>"
        Exit Sub
    End If
    
    totalPut = rsClass.RecordCount
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
            rsClass.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If

    Do While Not rsClass.EOF
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr align='center' class='title'><td width='50' height='22'>���ࣺ" & rsClass("ID") & "</td><td width='200' height='22'>�������ƣ�<a href='Admin_Page.asp?Action=ListPage&ClassName=" & rsClass("ClassName") & "'>" & rsClass("ClassName") & "</a></td><td width='160' height='22'>"
        If rsClass("ClassType") = 0 Then
            Response.Write "ϵͳ����"
        Else
            Response.Write "�Զ������"
        End If
        Response.Write "</td><td><a href='Admin_Page.asp?Action=ListPage&ClassName=" & rsClass("ClassName") & "'>" & PE_HTMLEncode(rsClass("ClassIntro")) & "</a></td>"
        Response.Write "<td width='210' align='center'><a href='Admin_Page.asp?Action=CreateClassFile&ClassID=" & rsClass("ID") & "'>���ɱ���</a>&nbsp;&nbsp;<a href='Admin_Page.asp?Action=ModifyClass&ClassID=" & rsClass("ID") & "'>�޸�</a>&nbsp;&nbsp;<a href='Admin_Page.asp?Action=AddPage&ClassName=" & rsClass("ClassName") & "'>������ҳ��</a>&nbsp;&nbsp;<a href='Admin_Page.asp?Action=DelClass&ClassID=" & rsClass("ID") & "' onclick=""return confirm('���Ҫɾ���˷�����');"">ɾ��</a>&nbsp;&nbsp;</td></tr>"
        Set rsPage = Conn.Execute("select ID,PageName,PageUrl,PageFileName,PageIntro from PE_Page Where ClassName='" & rsClass("ClassName") & "' Order by ID")
        If Not (rsPage.BOF And rsPage.EOF) Then
            Response.Write "<tr bgColor='#dddddd' align='center'><td>ID</td><td>ҳ������</td><td>ҳ���ַ</td><td>���</td><td>����</td></tr>"
            Do While Not rsPage.EOF
                Response.Write "  <tr class='tdbg' align='center' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
                Response.Write "    <td>" & rsPage("ID") & "</td>"
                Response.Write "    <td>" & rsPage("PageName") & "</td>"
                If Trim(rsPage("PageFileName") & "") = "" Then
                    Response.Write "<td><a href='" & InstallDir & "showpage.asp?id=" & rsPage("ID") & "' target='_blank'>" & InstallDir & "showpage.asp?id=" & rsPage("ID") & "</a></td>"
                    Response.Write "<td>" & PE_HTMLEncode(rsPage("PageIntro")) & "&nbsp;&nbsp;</td><td>"
                Else
                    On Error Resume Next
                    If ObjInstalled_FSO = True And rsPage("PageFileName") <> "" Then
                        If fso.FileExists(Server.MapPath(Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName"))) Then
                            Response.Write "<td><a href='" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "' target='_blank'>" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "</a></td>"
                            Response.Write "<td>" & PE_HTMLEncode(rsPage("PageIntro")) & "&nbsp;&nbsp;"
                            If Err Then
                                Response.Write "<font color=red><b>δ����</b></font>"
                            Else
                                Response.Write "<b><a href='" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "' target='_blank'>������</a></b>"
                            End If
                        Else
                            Response.Write "<td>" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "</td>"
                            Response.Write "<td>" & PE_HTMLEncode(rsPage("PageIntro")) & "&nbsp;&nbsp;"
                            Response.Write "<font color=red><b>δ����</b></font>"
                        End If
                    Else
                        Response.Write "<td>" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "</td>"
                        Response.Write "<td>" & PE_HTMLEncode(rsPage("PageIntro")) & "&nbsp;&nbsp;"
                        Response.Write "<font color=red><b>δ����</b></font>"
                    End If
                    Response.Write "</td><td>"
                    If ObjInstalled_FSO = False Then
                        Response.Write "<font color=red>FSO��</font>&nbsp;&nbsp;"
                    ElseIf Err Then
                        Response.Write "<font color=red>·����</font>&nbsp;&nbsp;"
                    Else
                        Response.Write "<a href='Admin_Page.asp?Action=CreateFile&PageID=" & rsPage("ID") & "'>���ɱ�ҳ</a>&nbsp;&nbsp;"
                    End If
                    Err.Clear
                End If
                Response.Write "<a href='Admin_Page.asp?Action=ModifyPage&PageID=" & rsPage("ID") & "'>�޸�</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_Page.asp?Action=DelPage&PageID=" & rsPage("ID") & "' onclick=""return confirm('���Ҫɾ����ҳ����');"">ɾ��</a>&nbsp;&nbsp;"
                Response.Write "</td></tr>"
                rsPage.movenext
            Loop
        Else
            Response.Write "<tr bgColor='#dddddd' align='center'><td colspan='5'>��������δ���ҳ��</td></tr>"
        End If
        rsPage.Close
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        Response.Write "</table><br>"
        rsClass.movenext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    Set rsPage = Nothing
    Response.Write "</form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������", True)
End Sub

Sub ListPage()
    Dim sqlPage, rsPage, rsClass, ClassName
    Dim iCount
    
    ClassName = Trim(Request("ClassName"))
    If ClassName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
        Exit Sub
    Else
        ClassName = ReplaceBadChar(ClassName)
        Set rsClass = Conn.Execute("select ClassName from PE_PageClass Where ClassName='" & ClassName & "'")
        If rsClass.BOF And rsClass.EOF Then
            rsClass.Close
            Set rsClass = Nothing
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>��δ��ӷ���" & ClassName & "��</li>"
            Exit Sub
        End If
        rsClass.Close
        Set rsClass = Nothing
    End If
    
    Response.Write "<form name='myform' method='post' action=''>"
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='40' height='22'>ҳ��ID</td>"
    Response.Write "    <td width='120' height='22'>ҳ������</td>"
    Response.Write "    <td width='200' height='22'>ҳ���ַ</td>"
    Response.Write "    <td>���</td>"
    Response.Write "    <td width='50' height='22'>������</td>"
    Response.Write "    <td width='180' align='center'>����</td>"
    Response.Write "  </tr>"
    
    Set rsPage = Server.CreateObject("Adodb.RecordSet")
    sqlPage = "select * from PE_Page Where ClassName='" & ClassName & "' Order by ID"
    rsPage.Open sqlPage, Conn, 1, 1
    If rsPage.BOF And rsPage.EOF Then
        rsPage.Close
        Set rsPage = Nothing
        Response.Write "<tr><td colspan='7' align='center'>��������δ����Զ���ҳ�棡</td></tr>"
        Exit Sub
    End If
    
    totalPut = rsPage.RecordCount
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
            rsPage.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If

    Do While Not rsPage.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td align='center'>" & rsPage("ID") & "</td>"
        Response.Write "    <td align='center'>" & rsPage("PageName") & "</td>"
        If Trim(rsPage("PageFileName") & "") = "" Then
            Response.Write "    <td align='center'><a href='" & InstallDir & "showpage.asp?id=" & rsPage("ID") & "' target='_blank'>" & InstallDir & "showpage.asp?id=" & rsPage("ID") & "</a></td>"
            Response.Write "    <td colspan='2'><a href='" & InstallDir & "showpage.asp?id=" & rsPage("ID") & "' target='_blank'>" & PE_HTMLEncode(rsPage("PageIntro")) & "</a></td>"
            Response.Write "    <td align='center'>"
        Else
            On Error Resume Next
            If ObjInstalled_FSO = True Then
                If fso.FileExists(Server.MapPath(Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName"))) Then
                    Response.Write "    <td align='center'><a href='" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "' target='_blank'>" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "</a></td>"
                        Response.Write "    <td><a href='" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "' target='_blank'>" & PE_HTMLEncode(rsPage("PageIntro")) & "</a></td>"
                    Response.Write "    <td align='center'><b>��</b></td>"
                Else
                    Response.Write "    <td align='center'><a href='Admin_Page.asp?Action=Modify&PageID=" & rsPage("ID") & "'>" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "</a></td>"
                    Response.Write "    <td><a href='Admin_Page.asp?Action=Modify&PageID=" & rsPage("ID") & "'>" & PE_HTMLEncode(rsPage("PageIntro")) & "</a></td>"
                    Response.Write "    <td align='center'><font color=red><b>��</b></font></td>"
                End If
            Else
                Response.Write "    <td align='center'><a href='Admin_Page.asp?Action=Modify&PageID=" & rsPage("ID") & "'>" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "</a></td>"
                Response.Write "    <td><a href='Admin_Page.asp?Action=Modify&PageID=" & rsPage("ID") & "'>" & PE_HTMLEncode(rsPage("PageIntro")) & "</a></td>"
                Response.Write "    <td align='center'><font color=red><b>��</b></font></td>"
            End If
            Response.Write "    <td align='center'>"
            If Err Then
                Response.Write "<font color=red>·����</font>&nbsp;&nbsp;"
            Else
                Response.Write "<a href='Admin_Page.asp?Action=CreateFile&PageID=" & rsPage("ID") & "'>����</a>&nbsp;&nbsp;"
            End If
        End If
        Response.Write "<a href='Admin_Page.asp?Action=ModifyPage&PageID=" & rsPage("ID") & "'>�޸�</a>&nbsp;&nbsp;"
        Response.Write "<a href='Admin_Page.asp?Action=DelPage&PageID=" & rsPage("ID") & "' onclick=""return confirm('���Ҫɾ����ҳ����');"">ɾ��</a>&nbsp;&nbsp;"
        If Trim(rsPage("PageFileName") & "") = "" Then
            Response.Write "<a href='Admin_Label.asp?Action=AddCai&PageUrl=" & InstallDir & "showpage.asp?id=" & rsPage("ID") & "'>�����ǩ</a>"
        ElseIf Not Err Then
            Response.Write "<a href='Admin_Label.asp?Action=AddCai&PageUrl=" & Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName") & "'>�����ǩ</a>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsPage.movenext
    Loop
    rsPage.Close
    Set rsPage = Nothing
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��ҳ��", True)
End Sub

Sub AddPage()
    Dim ClassName, strHead, Content
    
    ClassName = ReplaceBadChar(Trim(Request("ClassName")))
    Dim rsCheckOpti
    Set rsCheckOpti = Conn.Execute("select ClassName from PE_PageClass order by ID desc")
    If rsCheckOpti.BOF And rsCheckOpti.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��������Զ������,������Զ���ҳ�棡</li>"
		rsCheckOpti.Close
		set rsCheckOpti = Nothing
        Exit Sub
    End If
    '����ģ��Ԥ��ͷ�� �����ʱ�õ�
    strHead = "<html>" & vbCrLf
    strHead = strHead & "<head>" & vbCrLf
    strHead = strHead & "<title>��ģ�����</title>" & vbCrLf
    strHead = strHead & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
    strHead = strHead & "{$Skin_CSS} {$MenuJS}" & vbCrLf
    strHead = strHead & "</head>" & vbCrLf
    strHead = strHead & "<body leftmargin=0 topmargin=0 onmousemove='HideMenu()'>" & vbCrLf
    strHead = strHead & vbCrLf & "<!-- ��������Ҫ��ƵĴ��� -->" & vbCrLf
    strHead = strHead & vbCrLf & "</body>" & vbCrLf
    strHead = strHead & "</html>" & vbCrLf
        
    '�滻ͷ����ǩ Content Ϊ�滻��ͷ���ļ������ڱ༭����ʾcss
    
    Content = Replace(strHead, "{$Skin_CSS}", GetSkin_CSS(0))
    Content = Replace(Content, "{$MenuJS}", GetMenuJS("", False))
    Content = Replace(Content, "{$InstallDir}", InstallDir)
    
    Call ShowJSPage
      
    Response.Write "<form action='Admin_Page.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td align='center'><strong>�� �� �� �� �� ҳ ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>ҳ�����ƣ�</strong></td>"
    Response.Write "          <td><input name='PageName' type='text' id='PageName' size='30' maxlength='50'> <font color='#FF0000'>�����뱾ҳ�������</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>�������ࣺ</strong></td>"
    Response.Write "          <td><select name='ClassName' id='ClassName'>" & GetClassList(ClassName) & "</select><font color='#FF0000'>��ѡ����������</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>ҳ�����ͣ�</strong></td>"
    Response.Write "          <td><input name='PType' type='radio' value='0' onClick=""changetype(0);"" checked>��̬ҳ�� <input name='PType' type='radio' value='1' onClick=""changetype(1);"">��̬ҳ��</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tbody id='pathdiv'><tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>ҳ��·����</strong></td>"
    Response.Write "          <td><input name='PageUrl' type='text' id='PageUrl' size='30' maxlength='100'><font color='#FF0000'>����������·��(����дΪ��Ŀ¼)</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>�ļ����ƣ�</strong></td>"
    Response.Write "          <td><input name='PageFileName' type='text' id='PageFileName' size='30' maxlength='50' value=''><font color='#FF0000'>�����������ļ���(����дΪASP��ʽ)</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr></tbody>"
    Response.Write "    <tbody id='pathdiv2' style='display:none'><tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>����Ȩ�ޣ�</strong><br><font color=red>��ѡΪ����ҳ��</font></td>"
    Response.Write "          <td>" & GetUserGroup("", "") & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr></tbody>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><div id='itext'><strong>ҳ���飺</strong></div></td>"
    Response.Write "          <td><textarea name='PageIntro' cols='80' rows='5' id='PageIntro'></textarea></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td  align='center'><strong>ҳ �� �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "     <td >&nbsp;&nbsp;"
    Response.Write "        <textarea name='LabelContent' class='body2'   ROWS='10' COLS='108' onMouseUp=""setContent('get',1)"">" & strHead & "</textarea>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "     <td >&nbsp;"
    Response.Write "        <textarea name='LabelContent2'  style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
    Response.Write "        <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=1&TemplateType=0&tContentid=LabelContent2' frameborder='1' scrolling='no' width='780' height='400' ></iframe>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SavePage'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyPage()
    Dim PageID, sqlPage, rsPage, EditLabelContent, LabelContent, strTemp
    
    PageID = Trim(Request("PageID"))
    If PageID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
        Exit Sub
    Else
        PageID = PE_CLng(PageID)
    End If

    '�������Ȩ���ֶ��Ƿ����
    Dim i, dbrr
    Set rsPage = Conn.Execute("select top 1 * from PE_Page")
    For i = 0 To rsPage.Fields.Count - 1
        If rsPage.Fields(i).name = "arrGroupID" Then
            dbrr = True
        End If
    Next
    rsPage.Close
    Set rsPage = Nothing
    If dbrr <> True Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table PE_Page add arrGroupID nvarchar(255) null")
        Else
            Conn.Execute ("alter table PE_Page add arrGroupID varchar(255) null")
        End If
    End If

    sqlPage = "select * from PE_Page where ID=" & PageID
    Set rsPage = Conn.Execute(sqlPage)
    If rsPage.BOF And rsPage.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ���ָ���ı�ǩ��</li>"
        rsPage.Close
        Set rsPage = Nothing
        Exit Sub
    End If

    '����ı����ظ�����
    LabelContent = rsPage("PageContent")
    regEx.Pattern = "(\<\/textarea\>)"
    LabelContent = regEx.Replace(LabelContent, "[/textarea]")
    
    EditLabelContent = rsPage("PageContent")
    EditLabelContent = Replace(EditLabelContent, "<!--{$", "{$")
    EditLabelContent = Replace(EditLabelContent, "}-->", "}")
     
    'ͼƬ�滻JS
    regEx.Pattern = "(\<Script)(.[^\<]*)(\<\/Script\>)"
    Set Matches = regEx.Execute(EditLabelContent)
    For Each Match In Matches
        strTemp = Replace(Match.value, "<", "[!")
        strTemp = Replace(strTemp, ">", "!]")
        strTemp = Replace(strTemp, "'", """")
        strTemp = "<IMG alt='#" & strTemp & "#' src=""" & InstallDir & "editor/images/jscript.gif"" border=0 $>"
        EditLabelContent = Replace(EditLabelContent, Match.value, strTemp)
    Next
        
    'ͼƬ�滻������ǩ
    regEx.Pattern = "(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct)\((.*?)\)\}"
    EditLabelContent = regEx.Replace(EditLabelContent, "<IMG src=""" & InstallDir & "editor/images/label.gif"" border=0 zzz='$1($2)}'>")
    
    Call ShowJSPage
    
    Response.Write "<form action='Admin_Page.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td  align='center'><strong>�� �� �� �� �� ҳ ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>ҳ�����ƣ�</strong></td>"
    Response.Write "          <td><input name='PageName' type='text' id='PageName' size='30' maxlength='50' value='" & rsPage("PageName") & "'> <font color='#FF0000'>�����뱾ҳ������</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>�������ࣺ</strong></td>"
    Response.Write "<td><select name='ClassName' id='ClassName'>" & GetClassList(rsPage("ClassName")) & "</select><font color='#FF0000'>��ѡ����������</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>ҳ�����ͣ�</strong></td>"
    Response.Write "          <td><input name='PType' type='radio' value='0' onClick=""changetype(0);"""
    If Trim(rsPage("PageFileName") & "") <> "" Then Response.Write " checked"
    Response.Write ">��̬ҳ�� <input name='PType' type='radio' value='1' onClick=""changetype(1);"""
    If Trim(rsPage("PageFileName") & "") = "" Then Response.Write " checked"
    Response.Write ">��̬ҳ��</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If Trim(rsPage("PageFileName") & "") = "" Then
        Response.Write "    <tbody id='pathdiv' style='display:none'><tr class='tdbg'>"
    Else
        Response.Write "    <tbody id='pathdiv'><tr class='tdbg'>"
    End If
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>ҳ��·����</strong></td>"
    Response.Write "          <td><input name='PageUrl' type='text' id='PageUrl' size='30' maxlength='50' value='" & rsPage("PageUrl") & "'> <font color='#FF0000'>����������·��</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>�ļ����ƣ�</strong></td>"
    Response.Write "          <td><input name='PageFileName' type='text' id='PageFileName' size='30' maxlength='50' value='" & rsPage("PageFileName") & "'> <font color='#FF0000'>�����������ļ���</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr></tbody>"
    If Trim(rsPage("PageFileName") & "") = "" Then
        Response.Write "    <tbody id='pathdiv2'><tr class='tdbg'>"
    Else
        Response.Write "    <tbody id='pathdiv2' style='display:none'><tr class='tdbg'>"
    End If
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>����Ȩ�ޣ�</strong><br><font color=red>��ѡΪ����ҳ��</font></td>"
    Response.Write "          <td>" & GetUserGroup(rsPage("arrGroupID") & "", "") & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr></tbody>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    If Trim(rsPage("PageFileName") & "") = "" Then
        Response.Write "         <td width='100' align='center'><div id='itext'><strong>����������</strong></div></td>"
    Else
        Response.Write "         <td width='100' align='center'><div id='itext'><strong>ҳ���飺</strong></div></td>"
    End If
    Response.Write "         <td><textarea name='PageIntro' cols='80' rows='3' id='PageIntro'>" & rsPage("PageIntro") & "</textarea></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td  align='center'><strong>ҳ �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "     <td >&nbsp;&nbsp;"
    Response.Write "        <textarea name='LabelContent' class='body2'   ROWS='10' COLS='108' onMouseUp=""setContent('get',1)"">" & LabelContent & "</textarea>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "     <td >&nbsp;"
    Response.Write "        <textarea name='LabelContent2'  style='display:none' >" & Server.HTMLEncode(EditLabelContent) & "</textarea>"
    Response.Write "        <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=1&TemplateType=0&tContentid=LabelContent2' frameborder='1' scrolling='no' width='780' height='400' ></iframe>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'><input name='PageID' type='hidden' id='PageID' value='" & PageID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyPage'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' �����޸Ľ�� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsPage.Close
    Set rsPage = Nothing
End Sub

Sub SavePage()
    Dim PageID, PageName, PageUrl, PageFileName, PageIntro, ClassName, PType
    Dim rsPage, sqlPage, trs
    Dim PageContent, i
    
    '�������Ȩ���ֶ��Ƿ����
    Dim dbrr
    PType = PE_Clng(Trim(Request("PType")))
    Set rsPage = Conn.Execute("select top 1 * from PE_Page")
    For i = 0 To rsPage.Fields.Count - 1
        If rsPage.Fields(i).name = "arrGroupID" Then
            dbrr = True
        End If
    Next
    rsPage.Close
    Set rsPage = Nothing
    If dbrr <> True Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table PE_Page add arrGroupID nvarchar(255) null")
        Else
            Conn.Execute ("alter table PE_Page add arrGroupID varchar(255) null")
        End If
    End If

    PageID = Trim(Request.Form("PageID"))
    PageName = Trim(Request.Form("PageName"))
    ClassName = Trim(Request.Form("ClassName"))
    PageUrl = Trim(Request.Form("PageUrl"))
    PageFileName = Trim(Request.Form("PageFileName"))
    PageIntro = Trim(Request.Form("PageIntro"))
         
    For i = 1 To Request.Form("LabelContent").Count
        PageContent = PageContent & Request.Form("LabelContent")(i)
    Next
    
    If Action = "SaveModifyPage" Then
        If PageID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>��ָ��PageID</li>"
            Exit Sub
        Else
            PageID = PE_CLng(PageID)
        End If
    End If
    
    If PageName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>ҳ�����Ʋ���Ϊ�գ�</li>"
    Else
        PageName = ReplaceBadChar(PageName)
    End If
    
    If ClassName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>����ָ��һ�����࣡</li>"
    Else
        ClassName = ReplaceBadChar(ClassName)
    End If
    
    If PageUrl = "" Then
        PageUrl = "{$InstallDir}"
    Else
       If Left(PageUrl, 13) <> "{$InstallDir}" Then
            If Left(PageUrl, 1) <> "/" Then
                PageUrl = "{$InstallDir}" & PageUrl
            Else
                PageUrl = "{$InstallDir}" & Right(PageUrl, Len(PageUrl) - 1)
            End If
       End If
    End If
    
    If PageFileName <> "" Then
        PageFileName = ReplaceBadChar(PageFileName)
    End If
    
    If FoundErr = True Then Exit Sub
            
    If Action = "SaveModifyPage" Then
        sqlPage = "select * from PE_Page where ID=" & PageID
        Set rsPage = Server.CreateObject("ADODB.Recordset")
            rsPage.Open sqlPage, Conn, 1, 3
            rsPage("PageName") = PageName
            rsPage("ClassName") = ClassName
            rsPage("PageUrl") = PageUrl
            IF PType = 0 then
                If PageFileName <> "" Then rsPage("PageFileName") = PageFileName
            Else
                rsPage("PageFileName") = ""
            End IF
            rsPage("PageIntro") = PageIntro
            rsPage("PageContent") = PageContent
            rsPage("arrGroupID") = Trim(Request("GroupID"))
            rsPage.Update
        rsPage.Close
        Set rsPage = Nothing
        If ObjInstalled_FSO = True Then
            If fso.FileExists(Server.MapPath(PageUrl & PageFileName)) Then
                Call CreateFile(PageID)
            End If
        End If
        Call WriteSuccessMsg("�޸��Զ���ҳ��ɹ���", ComeUrl & "")
    Else
        sqlPage = "select top 1 * from PE_Page"
        Set rsPage = Server.CreateObject("ADODB.Recordset")
        rsPage.Open sqlPage, Conn, 1, 3
        rsPage.addnew
        rsPage("PageName") = PageName
        rsPage("ClassName") = ClassName
        rsPage("PageUrl") = PageUrl
        If PageFileName <> "" Then rsPage("PageFileName") = PageFileName
        rsPage("PageIntro") = PageIntro
        rsPage("PageContent") = PageContent
        rsPage("arrGroupID") = Trim(Request("arrGroupID"))
        rsPage.Update
        rsPage.Close
        Set rsPage = Nothing
        Call WriteSuccessMsg("�����Զ���ҳ��ɹ���", ComeUrl & "")
    End If
End Sub

Sub DelPage()
    Dim PageID, sqlPage, rsPage, tPageContent
    PageID = Trim(Request("PageID"))
    If PageID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
        Exit Sub
    Else
        PageID = PE_CLng(PageID)
    End If
    sqlPage = "select * from PE_Page where ID=" & PageID
    Set rsPage = Server.CreateObject("ADODB.Recordset")
    rsPage.Open sqlPage, Conn, 1, 3
    If rsPage.BOF And rsPage.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ���ָ����ҳ�棡</li>"
        rsPage.Close
        Set rsPage = Nothing
        Exit Sub
    Else
        If ObjInstalled_FSO = True Then
            If fso.FileExists(Server.MapPath(Replace(rsPage("PageUrl") & rsPage("PageFileName"),"{$InstallDir}",InstallDir))) Then
                fso.DeleteFile Server.MapPath(Replace(rsPage("PageUrl") & rsPage("PageFileName"),"{$InstallDir}",InstallDir))
            End If
        End If
    End If
    rsPage.Delete
    rsPage.Update
    rsPage.Close
    Set rsPage = Nothing
End Sub

Sub ShowJSClass()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.ClassName.value==''){" & vbCrLf
    Response.Write "     alert('�������Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.ClassName.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub AddClass()
    Call ShowJSClass

    Response.Write "<form action='Admin_Page.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "    <td align='center'><strong>�� �� �� �� �� �� ��</font></strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>�������ƣ�</strong></td>"
    Response.Write "          <td><input name='ClassName' type='text' id='ClassName' size='30' maxlength='50'> <font color='#FF0000'>�����뱾���������</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "         <td width='100' align='center'><strong>�����飺</strong></td>"
    Response.Write "         <td><textarea name='ClassIntro' cols='80' rows='3' id='ClassIntro'></textarea></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveClass'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyClass()
    Dim ClassID, sqlClass, rsClass
    
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    sqlClass = "select * from PE_PageClass where ID=" & ClassID
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ���ָ���ı�ǩ��</li>"
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    End If
        
    Call ShowJSClass
    
    Response.Write "<form action='Admin_Page.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td  align='center'><strong>�� �� �� �� �� �� Ŀ</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>�������ƣ�</strong></td>"
    Response.Write "          <td><input name='ClassName' type='text' id='ClassName' size='30' maxlength='50' value='" & rsClass("ClassName") & "'> <font color='#FF0000'>�����뱾��������</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "         <td width='100' align='center'><strong>�����飺</strong></td>"
    Response.Write "         <td><textarea name='ClassIntro' cols='80' rows='3' id='ClassIntro'>" & rsClass("ClassIntro") & "</textarea></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'><input name='ClassID' type='hidden' id='ClassID' value='" & ClassID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyClass'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' �����޸Ľ�� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub SaveClass()
    Dim ClassID, ClassName, ClassIntro, tempClassName
    Dim rsClass, sqlClass, trs
    ClassID = Trim(Request.Form("ClassID"))
    ClassName = Trim(Request.Form("ClassName"))
    ClassIntro = Trim(Request.Form("ClassIntro"))

    If Action = "SaveModifyClass" Then
        If ClassID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>��ָ��ClassID</li>"
            Exit Sub
        Else
            ClassID = PE_CLng(ClassID)
        End If
    End If
    
    If ClassName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�������Ʋ���Ϊ�գ�</li>"
    Else
        ClassName = ReplaceBadChar(ClassName)
    End If
    
    If FoundErr = True Then Exit Sub
            
    If Action = "SaveModifyClass" Then
        sqlClass = "select * from PE_PageClass where ID=" & ClassID
        Set rsClass = Server.CreateObject("ADODB.Recordset")
        rsClass.Open sqlClass, Conn, 1, 3
            tempClassName = rsClass("ClassName")
            rsClass("ClassName") = ClassName
            rsClass("ClassIntro") = ClassIntro
        rsClass.Update
        sqlClass = "select * from PE_Page where ClassName='" & tempClassName & "'"
        Set rsClass = Server.CreateObject("ADODB.Recordset")
        rsClass.Open sqlClass, Conn, 1, 3
        Do While Not rsClass.EOF
            rsClass("ClassName") = ClassName
            rsClass.Update
            rsClass.movenext
        Loop
        rsClass.Close
        Set rsClass = Nothing
        Call WriteSuccessMsg("�޸��Զ������ɹ���", ComeUrl & "")
    Else
        sqlClass = "select top 1 * from PE_PageClass where ClassName='" & ClassName & "'"
        Set rsClass = Server.CreateObject("ADODB.Recordset")
        rsClass.Open sqlClass, Conn, 1, 3
        If rsClass.BOF And rsClass.EOF Then
            rsClass.addnew
            rsClass("ClassName") = ClassName
            rsClass("ClassIntro") = ClassIntro
            rsClass("ClassType") = 1
            rsClass.Update
            rsClass.Close
            Set rsClass = Nothing
        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>��ָ��ClassID</li>"
            rsClass.Close
            Set rsClass = Nothing
            Exit Sub
        End If
        Call WriteSuccessMsg("�����Զ������ɹ���", ComeUrl & "")
    End If
End Sub

Sub DelClass()
    Dim ClassID, sqlClass, rsClass, rsPage
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    sqlClass = "select * from PE_PageClass where ID=" & ClassID
    Set rsClass = Server.CreateObject("ADODB.Recordset")
    rsClass.Open sqlClass, Conn, 1, 3
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ���ָ���ķ��࣡</li>"
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    Else
        sqlClass = "select * from PE_Page where ClassName='" & rsClass("ClassName") & "'"
        Set rsPage = Server.CreateObject("ADODB.Recordset")
        rsPage.Open sqlClass, Conn, 1, 3
        If Not (rsPage.BOF And rsPage.EOF) Then
            Do While Not rsPage.EOF
                If ObjInstalled_FSO = True Then
                    If fso.FileExists(Server.MapPath(rsPage("PageUrl") & rsPage("PageFileName"))) Then
                        fso.DeleteFile Server.MapPath(rsPage("PageUrl") & rsPage("PageFileName"))
                    End If
                End If
                rsPage.Delete
                rsPage.Update
                rsPage.movenext
            Loop
        End If
        rsPage.Close
        Set rsPage = Nothing
    End If
    rsClass.Delete
    rsClass.Update
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub CreateFile(iPageID)
    If ObjInstalled_FSO = True Then
        Dim PageID, sqlPage, rsPage, tPageContent

        If iPageID = "" Then
            PageID = Trim(Request("PageID"))
        Else
            PageID = iPageID
        End If
        If PageID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
            Exit Sub
        Else
            PageID = PE_CLng(PageID)
        End If
        sqlPage = "select PageName,PageUrl,PageFileName,PageContent from PE_Page where ID=" & PageID
        Set rsPage = Server.CreateObject("ADODB.Recordset")
        rsPage.Open sqlPage, Conn, 1, 1
        If rsPage.BOF And rsPage.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>�Ҳ���ָ����ҳ�棡</li>"
            rsPage.Close
            Set rsPage = Nothing
            Exit Sub
        Else
            If Trim(rsPage("PageFileName") & "") <> "" Then
                strHTML = rsPage("PageContent")
                Call ReplaceCommonLabel

                strHTML = Replace(strHTML, "{$ShowPath}", rsPage("PageName"))
                strHTML = Replace(strHTML, "{$MenuJS}", GetMenuJS("", False))
                strHTML = Replace(strHTML, "{$Skin_CSS}", GetSkin_CSS(0))
        
                Dim PE_Content
                Set PE_Content = New Article
                PE_Content.Init
                strHTML = PE_Content.GetCustomFromTemplate(strHTML)
                strHTML = PE_Content.GetPicFromTemplate(strHTML)
                strHTML = PE_Content.GetListFromTemplate(strHTML)
                strHTML = PE_Content.GetSlidePicFromTemplate(strHTML)
    
                Set PE_Content = New Soft
                PE_Content.Init
                strHTML = PE_Content.GetCustomFromTemplate(strHTML)
                strHTML = PE_Content.GetPicFromTemplate(strHTML)
                strHTML = PE_Content.GetListFromTemplate(strHTML)
                strHTML = PE_Content.GetSlidePicFromTemplate(strHTML)
    
                Set PE_Content = New Photo
                PE_Content.Init
                strHTML = PE_Content.GetPicFromTemplate(strHTML)
                strHTML = PE_Content.GetListFromTemplate(strHTML)
                strHTML = PE_Content.GetSlidePicFromTemplate(strHTML)
                Set PE_Content = Nothing

                Set PE_Content = New Product
                PE_Content.Init
                strHTML = PE_Content.GetPicFromTemplate(strHTML)
                strHTML = PE_Content.GetListFromTemplate(strHTML)
                strHTML = PE_Content.GetSlidePicFromTemplate(strHTML)
                strHTML = PE_Content.GetCustomFromTemplate(strHTML)

                Set PE_Content = Nothing

                If fso.FolderExists(Server.MapPath(Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir))) = False Then
                    fso.CreateFolder Server.MapPath(Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir))
                End If
                Call WriteToFile(Replace(rsPage("PageUrl"), "{$InstallDir}", InstallDir) & rsPage("PageFileName"), strHTML)
            End If
        End If
        rsPage.Close
        Set rsPage = Nothing
    End If
End Sub

Sub CreateClassFile()
    Dim ClassID, ClassName, PageID, rsClass, rsPage2
    ClassID = PE_CLng(Trim(Request("ClassID")))
    If ClassID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ʧ��</li>"
        Exit Sub
    End If
    Set rsClass = Conn.Execute("select ClassName from PE_PageClass where ID=" & ClassID)
    If Not (rsClass.BOF And rsClass.EOF) Then
        Set rsPage2 = Conn.Execute("select ID from PE_Page where ClassName='" & rsClass("ClassName") & "'")
        Do While Not rsPage2.EOF
            Call CreateFile(rsPage2("ID"))
            rsPage2.movenext
        Loop
        rsPage2.Close
        Set rsPage2 = Nothing
    End If
    rsClass.Close
    Set rsClass = Nothing
End Sub

Function GetClassList(iClassName)
    Dim optiTemp, rsOpti
    Set rsOpti = Conn.Execute("select ClassName from PE_PageClass order by ID desc")
    If rsOpti.BOF And rsOpti.EOF Then
        optiTemp = "<option value='0'>�������һ������</option>"
    Else
        Do While Not rsOpti.EOF
            optiTemp = optiTemp & "<option value='" & rsOpti("ClassName") & "'"
            If iClassName = rsOpti("ClassName") Then optiTemp = optiTemp & " selected"
            optiTemp = optiTemp & ">" & rsOpti("ClassName") & "</option>"
            rsOpti.movenext
        Loop
    End If
    GetClassList = optiTemp
End Function

Sub ShowJSPage()
    Dim TrueSiteUrl
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function changetype(itype){" & vbCrLf
    Response.Write " if (itype==0){" & vbCrLf
    Response.Write "     pathdiv.style.display='';" & vbCrLf
    Response.Write "     document.getElementById('itext').innerHTML='<strong>ҳ���飺</strong>';" & vbCrLf
    Response.Write "     if (document.myform.PageIntro.value=='����һ|0|false|0'){;" & vbCrLf
    Response.Write "         document.myform.PageIntro.value='';" & vbCrLf
    Response.Write "     }" & vbCrLf
    Response.Write "     pathdiv2.style.display='none';" & vbCrLf
    Response.Write " }else{" & vbCrLf
    Response.Write "     pathdiv.style.display='none';" & vbCrLf
    Response.Write "     document.getElementById('itext').innerHTML='<strong>����������</strong>';" & vbCrLf
    Response.Write "     if (document.myform.PageIntro.value==''){;" & vbCrLf
    Response.Write "         document.myform.PageIntro.value='����һ|0|false|0';" & vbCrLf
    Response.Write "     }" & vbCrLf
    Response.Write "     pathdiv2.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.PageName.value==''){" & vbCrLf
    Response.Write "     alert('ҳ�����Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.PageName.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.LabelContent2.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.LabelContent.value==''){" & vbCrLf
    Response.Write "     alert('ҳ�����ݲ���Ϊ�գ�');" & vbCrLf
    Response.Write "     document.myform.LabelContent.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (Strsave==""B""){" & vbCrLf
    Response.Write "      setContent (""get"",1);" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<script language=""VBScript"">" & vbCrLf
    Response.Write "    Dim regEx, Match, Matches, StrBody,strTemp,strMatch,arrMatch,i,Strsave" & vbCrLf
    Response.Write "    Dim Content,arrContent" & vbCrLf
    Response.Write "    Set regEx = New RegExp" & vbCrLf
    Response.Write "    regEx.IgnoreCase = True" & vbCrLf
    Response.Write "    regEx.Global = True" & vbCrLf
    Response.Write "    Strsave=""A""" & vbCrLf
    '=================================================
    '��  �ã�����html
    '=================================================
    Response.Write "Function  Resumeblank(byval Content)" & vbCrLf
    Response.Write " Dim strHtml,strHtml2,Num,Numtemp,Strtemp" & vbCrLf
    Response.Write "   strHtml=Replace(Content, ""<DIV"", ""<div"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DIV>"", ""</div>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TABLE"", ""<table"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TABLE>"", vbCrLf & ""</table>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TBODY>"", """")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TBODY>"","""" & vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TR"", ""<tr"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TR>"", vbCrLf & ""</tr>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TD"", ""<td"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TD>"", ""</td>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<!--"", vbCrLf & ""<!--"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<SELECT"",vbCrLf & ""<Select"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</SELECT>"",vbCrLf & ""</Select>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<OPTION"",vbCrLf & ""  <Option"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</OPTION>"",""</Option>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<INPUT"",vbCrLf & ""  <Input"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<script"",vbCrLf & ""<script"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""&amp;"",""&"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""{$--"",vbCrLf & ""<!--$"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""--}"",""$-->"")" & vbCrLf
    Response.Write "   arrContent = Split(strHtml,vbCrLf)" & vbCrLf
    Response.Write "    For i = 0 To UBound(arrContent)" & vbCrLf
    Response.Write "        Numtemp=false" & vbCrLf
    Response.Write "        if Instr(arrContent(i),""<table"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<table"" and Strtemp <>""</table>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<table""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<tr"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<tr"" and Strtemp<>""</tr>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<tr""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<td"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<td"" and Strtemp<>""</td>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<td""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</table>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</table>"" and Strtemp<>""<table"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</table>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</tr>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</tr>"" and Strtemp<>""<tr"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</tr>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</td>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</td>"" and Strtemp<>""<td"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</td>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<!--"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Num< 0 then Num = 0" & vbCrLf
    Response.Write "        if trim(arrContent(i))<>"""" then" & vbCrLf
    Response.Write "            if i=0 then" & vbCrLf
    Response.Write "                strHtml2= string(Num,"" "") & arrContent(i) " & vbCrLf
    Response.Write "            elseif Numtemp=True then" & vbCrLf
    Response.Write "                strHtml2= strHtml2 & vbCrLf & string(Num,"" "") & arrContent(i) " & vbCrLf
    Response.Write "            else" & vbCrLf
    Response.Write "                strHtml2= strHtml2 & vbCrLf & arrContent(i) " & vbCrLf
    Response.Write "            end if" & vbCrLf
    Response.Write "        end if" & vbCrLf
    Response.Write "      Next" & vbCrLf
    Response.Write "      Resumeblank=strHtml2" & vbCrLf
    Response.Write "    End function" & vbCrLf
    Response.Write "    function setContent(zhi,TpyeTemplate)" & vbCrLf
    Response.Write "      if zhi=""get"" then" & vbCrLf
    Response.Write "        if Strsave=""A"" then Exit Function" & vbCrLf
    Response.Write "        Strsave=""A""" & vbCrLf
    Response.Write "        TemplateContent= document.myform.LabelContent.value" & vbCrLf
    Response.Write "        TemplateContent2= editor.HtmlEdit.document.body.innerHTML" & vbCrLf
    Response.Write "        if TemplateContent="""" then " & vbCrLf
    Response.Write "            alert ""��ɾ���˴������ҳ�����������д��ҳ ��""" & vbCrLf
    Response.Write "            Exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body>"")=0 then" & vbCrLf
    Response.Write "            regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "            Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "            For Each Match In Matches" & vbCrLf
    Response.Write "                StrBody = Match.Value" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            StrBody=""<body>"" " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body"")>0 then" & vbCrLf
    Response.Write "            Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\}['|""""]\>""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(TemplateContent2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            regEx.Pattern = ""\{\$(.*)\}""" & vbCrLf
    Response.Write "            Set strTemp = regEx.Execute(replace(Match.Value,"" "",""""))" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
    Response.Write "                TemplateContent2 = Replace(TemplateContent2, Match.Value, ""<!--"" & strTemp2 & ""-->"")" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\$\>""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(TemplateContent2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            regEx.Pattern = ""\#(.*)\#""" & vbCrLf
    Response.Write "            Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""&amp;"", ""&"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""&13;&10;"",vbCrLf)" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""&9;"",vbTab)" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
    Response.Write "                TemplateContent2 = Replace(TemplateContent2, Match.Value, strTemp2)" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        TemplateContent2=Resumeblank(TemplateContent2)" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2,""{$InstallDir}{$rsClass_ClassUrl}"",""{$rsClass_ClassUrl}"")" & vbCrLf
    Response.Write "        regEx.Pattern = ""\{\$InstallDir\}editor.asp(.[^\<]*?)\#""" & vbCrLf
    Response.Write "        TemplateContent2 = regEx.Replace(TemplateContent2, ""#"")" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body"")=0 then" & vbCrLf
    Response.Write "            document.myform.LabelContent.value=TemplateContent2" & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            document.myform.LabelContent.value=Content1& vbCrLf &TemplateContent2& vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "    Else" & vbCrLf
    Response.Write "        if Strsave=""B"" then Exit Function" & vbCrLf
    Response.Write "        Strsave=""B""" & vbCrLf
    Response.Write "        TemplateContent= document.myform.LabelContent.value" & vbCrLf
    Response.Write "        if TemplateContent="""" then " & vbCrLf
    Response.Write "            alert ""��ɾ���˴������ҳ�����������д��ҳ ��""" & vbCrLf
    Response.Write "            Exit function" & vbCrLf
    Response.Write "           " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body>"")=0 then" & vbCrLf
    Response.Write "            regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "            Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "            For Each Match In Matches" & vbCrLf
    Response.Write "                StrBody = Match.Value" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            StrBody=""<body>"" " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body"")>0 then" & vbCrLf
    Response.Write "            Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "            Content2 = arrContent(1)" & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "            Content2 = arrContent(0)" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""<!--$"", ""{$--"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""$-->"", ""--}"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""<!--{$"", ""{$"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""}-->"", ""}"")" & vbCrLf
    Response.Write "        'ͼƬ�滻JS" & vbCrLf
    Response.Write "        regEx.Pattern = ""(\<Script)([\s\S]*?)(\<\/Script\>)""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(Content2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            strTemp = Replace(Match.Value, ""<"", ""[!"")" & vbCrLf
    Response.Write "            strTemp = Replace(strTemp, "">"", ""!]"")" & vbCrLf
    Response.Write "            strTemp = Replace(strTemp, ""'"", """""""")" & vbCrLf
    Response.Write "            strTemp = ""<IMG alt='#"" & strTemp & ""#' src=""""" & InstallDir & "editor/images/jscript.gif"""" border=0 $>""" & vbCrLf
    Response.Write "            Content2 = Replace(Content2, Match.Value, strTemp)" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        'ͼƬ�滻������ǩ" & vbCrLf
    Response.Write "        regEx.Pattern = ""(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct)\((.*?)\)\}""" & vbCrLf
    Response.Write "        Content2 = regEx.Replace(Content2, ""<IMG src=""""" & InstallDir & "editor/images/label.gif"""" border=0 zzz='$1($2)}'>"")" & vbCrLf
    Response.Write "        regEx.Pattern = ""\{\$InstallDir\}""" & vbCrLf
    Response.Write "        Content2 = regEx.Replace(Content2,""http://" & TrueSiteUrl & InstallDir & """)" & vbCrLf
    Response.Write "        editor.HtmlEdit.document.body.innerHTML=Content2" & vbCrLf
    Response.Write "        editor.showBorders()" & vbCrLf
    Response.Write "    End if" & vbCrLf
    Response.Write "    End function" & vbCrLf
    Response.Write "    function setstatus()" & vbCrLf 'Ϊ323 �����editor.asp ��Ч����
    Response.Write "    end function" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'��������Import
'��  �ã�����ģ���һ��
'=================================================
Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Page.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>ҳ�浼�루��һ����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ�����ģ�����ݿ���ļ����� "
    Response.Write "        <input name='PageMdb' type='text' id='PageMdb' value='../Temp/PE_Page.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='import2'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������import2
'��  �ã�����ģ��ڶ���
'=================================================
Sub import2()
    'On Error Resume Next

    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    
    '��õ���ģ�����ݿ�·��
    mdbname = Replace(Trim(Request.Form("PageMdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("PageMdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "��", "/") '��ֹ�ⲿ���Ӱ�ȫ����

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
        Exit Sub
    End If

    '��������ģ�����ݿ�
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    Response.Write "<form name='myform' method='post' action='Admin_Page.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>ҳ�浼�루�ڶ�����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>��Ҫ�����ҳ�����</strong></td>"
    Response.Write "          </tr>"
    Response.Write "           <tr>"
    Response.Write "            <td>"
    
    '��ʾģ��
    Response.Write "              <select name='ClassID' size='2' multiple style='height:300px;width:250px;'>"
    
    sql = "select ID,ClassName from PE_PageClass"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1
    If rs.BOF And rs.EOF Then
        'û��ģ��ʱָ���ر��ύ��ť
        Response.Write "                <option value='0'>û���κ�ҳ�����</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("ID") & "'>" & rs("ClassName") & "</option>"
            rs.movenext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"
    Response.Write "                  </td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='10'></td></tr>"
    Response.Write "                  <tr>"
    Response.Write "                    <td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='20'></td></tr>"
    Response.Write "                  <tr><td colspan='3' height='25' align='center'><input type='submit' name='Submit' value=' ����ҳ�� ' onClick=""document.myform.Action.value='Doimport';"""
    Response.Write "                 </td></tr>"
    Response.Write "               </table>"
    Response.Write "               <input name='PageMdb' type='hidden' id='PageMdb' value='" & mdbname & "'>"
    Response.Write "               <input name='Action' type='hidden' id='Action' value='Doimport'>"
    Response.Write "               <br>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "       </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������DoImport
'��  �ã�����ģ�屣��
'=================================================
Sub DoImport()
    On Error Resume Next
    
    Dim trs, crs, mdbname, tconn
    Dim ClassID, rs, sql, rsLabel, Table_PE_lable
    ClassID = ReplaceBadChar(Trim(Request.Form("ClassID")))
    mdbname = Replace(Trim(Request.Form("PageMdb")), "'", "")
    
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�������Ʋ���Ϊ�գ�</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
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

    Set trs = tconn.Execute(" select * from PE_PageClass where ID in (" & ClassID & ") order by ID")
    Do While Not trs.EOF
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open "select * from PE_PageClass where ClassName='" & trs("ClassName") & "'", Conn, 1, 3
        If rs.BOF And rs.EOF Then
            rs.addnew
            rs("ClassName") = trs("ClassName")
            rs("ClassIntro") = trs("ClassIntro")
            If trs("ClassType") <> "" And Not IsNull(trs("ClassType")) Then
                rs("ClassType") = trs("ClassType")
            Else
                rs("ClassType") = 0
            End If
            rs.Update
        Else
            ErrMsg = ErrMsg & "<li>ҳ�����" & trs("ClassName") & "�Ѿ�����,�޷�����"
            rs.Close
            Set rs = Nothing
            trs.Close
            Set trs = Nothing
            Err.Clear
            Exit Sub
        End If
        rs.Close
        
        Set crs = tconn.Execute(" select * from PE_Page where ClassName = '" & trs("ClassName") & "'")
            Set rs = Server.CreateObject("adodb.recordset")
            rs.Open "select * from PE_Page", Conn, 1, 3
            Do While Not crs.EOF
                rs.addnew
                rs("ClassName") = crs("ClassName")
                rs("PageName") = crs("PageName")
                rs("PageIntro") = crs("PageIntro")
                rs("PageUrl") = crs("PageUrl")
                rs("PageFileName") = crs("PageFileName")
                If crs("PageType") <> "" And Not IsNull(crs("PageType")) Then
                    rs("PageType") = crs("PageType")
                Else
                    rs("PageType") = 0
                End If
                rs("PageContent") = crs("PageContent")
                rs("arrGroupID") = crs("arrGroupID")
                rs.Update
                crs.movenext
            Loop
            rs.Close
        trs.movenext
    Loop
    Set rs = Nothing
    crs.Close
    Set crs = Nothing
    trs.Close
    Set trs = Nothing
   
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ���ָ�������ݿ��е���ѡ�е�ҳ�棡", ComeUrl & "?Action=Import2&PageMdb=" & Replace(mdbname, "/", "��") & "")
End Sub

'=================================================
'��������Export
'��  �ã�����ģ��
'=================================================
Sub Export()
    Dim rs, sql
    Dim trs, iCount
 
    Response.Write "<form name='myform' method='post' action='Admin_Page.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>ҳ�浼��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='ClassID' size='2' multiple style='height:300px;width:450px;'>"
    
    sql = "select ID,ClassName from PE_PageClass Order by ID"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>û���κ�ҳ��</option>"
        '�ر��ύ��ť
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("ID") & "'>" & rs("ClassName") & "</option>"
            rs.movenext
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
    Response.Write "        <td colspan='2'>Ŀ�����ݿ⣺<input name='PageMdb' type='text' id='PageMdb' value='../Temp/PE_Page.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> �����Ŀ�����ݿ�</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='ִ�е�������' onClick=""document.myform.Action.value='Doexport';"">"
    Response.Write "              <input name='Action' type='hidden' id='Action' value='Doexport'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.ClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.ClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'��������DoExport
'��  �ã�����ģ��
'=================================================
Sub DoExport()
    'On Error Resume Next
    Dim mdbname, tconn, trs, crs
    Dim ClassID, rs, rs2, FormatConn
    
    ClassID = Trim(Request.Form("ClassID"))
    FormatConn = Request.Form("FormatConn")
    mdbname = Replace(Trim(Request.Form("PageMdb")), "'", "")
    If IsValidID(ClassID) = False Then
        ClassID = ""
    End If

    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ������ҳ��</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ҳ�����ݿ���</li>"
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
      
    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_Page")
        tconn.Execute ("delete from PE_PageClass")
    End If

    Set rs = Conn.Execute("select * from PE_PageClass where ID in (" & ClassID & ")  order by ID")
    Do While Not rs.EOF
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_PageClass", tconn, 1, 3
        trs.addnew
        trs("ClassName") = rs("ClassName")
        trs("ClassIntro") = rs("ClassIntro")
        If rs("ClassType") <> "" And Not IsNull(rs("ClassType")) Then
            trs("ClassType") = rs("ClassType")
        Else
            trs("ClassType") = 0
        End If
        trs.Update
        trs.Close
        
        Set rs2 = Conn.Execute("select * from PE_Page where ClassName = '" & rs("ClassName") & "'")
        trs.Open "select * from PE_Page", tconn, 1, 3
        Do While Not rs2.EOF
            trs.addnew
                    trs("ClassName") = rs2("ClassName")
                    trs("PageName") = rs2("PageName")
                    trs("PageIntro") = rs2("PageIntro")
                    trs("PageUrl") = rs2("PageUrl")
                    trs("PageFileName") = rs2("PageFileName")
                    If rs2("PageType") <> "" And Not IsNull(rs2("PageType")) Then
                        trs("PageType") = rs2("PageType")
                    Else
                        trs("PageType") = 0
                    End If
                    trs("PageContent") = rs2("PageContent")
                    trs("arrGroupID") = rs2("arrGroupID")
            trs.Update
            rs2.movenext
        Loop
    rs.movenext
    Loop
    
    trs.Close
    Set trs = Nothing
    rs2.Close
    Set rs2 = Nothing
    rs.Close
    Set rs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ�����ѡ�е��Զ���ҳ�����õ�����ָ�������ݿ��У�", ComeUrl)
End Sub
%>
