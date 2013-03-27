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

Dim TypeSelect, SelectedName

TypeSelect = Trim(Request("TypeSelect"))
'������Ա����Ȩ��
If AdminPurview > 1 Then
    Select Case TypeSelect
    Case "Keyword", "AddKeyword", "ModifyKeyword", "SaveAddKeyword", "SaveModifyKeyword", "DelKeyword", "DelAllKeyword"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Keyword_" & ChannelDir)
    Case "Author", "AddAuthor", "ModifyAuthor", "SaveAddAuthor", "SaveModifyAuthor", "DelAuthor", "AuthorDis", "AuthorEn", "AuthorDTop", "AuthorTop", "AuthorDElite", "AuthorElite"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Author_" & ChannelDir)
    Case "CopyFrom", "AddCopyFrom", "ModifyCopyFrom", "SaveAddCopyFrom", "SaveModifyCopyFrom", "DelCopyFrom", "CopyFromDis", "CopyFromEn", "CopyFromDTop", "CopyFromTop", "CopyFromDElite", "CopyFromElite"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Copyfrom_" & ChannelDir)
    Case "Producer", "AddProducer", "ModifyProducer", "SaveAddProducer", "SaveModifyProducer", "DelProducer", "ProducerDis", "ProducerEn", "ProducerDTop", "ProducerTop", "ProducerDElite", "ProducerElite"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Producer_Shop")
    Case "Trademark", "AddTrademark", "ModifyTrademark", "SaveAddTrademark", "SaveModifyTrademark", "DelTrademark", "TrademarkDis", "TrademarkEn", "TrademarkDTop", "TrademarkTop", "TrademarkDElite", "TrademarkElite"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Trademark_Shop")
    Case "KeyLink", "AddKeyLink", "ModifyKeyLink", "SaveAddKeyLink", "SaveModifyKeyLink", "DelKeyLink", "runKeyLink", "disKeyLink"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "KeyLink")
    Case "Rtext", "AddRtext", "ModifyRtext", "SaveAddRtext", "SaveModifyRtext", "DelRtext", "runRtext", "disRtext"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Rtext")
    Case Else
        PurviewPassed = False
    End Select
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
        Call WriteEntry(6, AdminName, "ԽȨ����")
        Response.End
    End If
End If

XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

Dim ItemType
ItemType = Trim(Request("ItemType"))
If ItemType = "" Then
    ItemType = 999
Else
    ItemType = PE_CLng(ItemType)
End If

Select Case TypeSelect
Case "Keyword"
    SelectedName = "�ؼ��ֹ���"
Case "AddKeyword"
    SelectedName = "�����ؼ���"
Case "ModifyKeyword"
    SelectedName = "�޸Ĺؼ���"
Case "Author"
    SelectedName = "���߹���"
Case "AddAuthor"
    SelectedName = "����������Ϣ"
Case "ModifyAuthor"
    SelectedName = "�޸�������Ϣ"
Case "CopyFrom"
    SelectedName = "��Դ����"
Case "AddCopyFrom"
    SelectedName = "������Դ��Ϣ"
Case "ModifyCopyFrom"
    SelectedName = "�޸���Դ��Ϣ"
Case "KeyLink"
    SelectedName = "վ�����ӹ���"
Case "AddKeyLink"
    SelectedName = "����վ������"
Case "ModifyKeyLink"
    SelectedName = "�޸�վ������"
Case "Rtext"
    SelectedName = "�ַ��滻����"
Case "AddRtext"
    SelectedName = "�����ַ��滻"
Case "ModifyRtext"
    SelectedName = "�޸��ַ��滻"
Case "Producer"
    SelectedName = "���̹���"
Case "AddProducer"
    SelectedName = "����������Ϣ"
Case "ModifyProducer"
    SelectedName = "�޸ĳ�����Ϣ"
Case "Trademark"
    SelectedName = "Ʒ�ƹ���"
Case "AddTrademark"
    SelectedName = "����Ʒ��"
Case "ModifyTrademark"
    SelectedName = "�޸�Ʒ��"
Case Else
    SelectedName = "������վ����ϵͳ"
End Select
    
'ȡƵ���б�
Dim ChannelList, rsChannel
ChannelList = "<option value=0"
If ChannelID = 0 Then ChannelList = ChannelList & "selected"
ChannelList = ChannelList & ">ȫ��Ƶ��</option>"

Set rsChannel = Conn.Execute("select ChannelID,ChannelName,OrderID,ModuleType,Disabled from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID")
If Not (rsChannel.BOF And rsChannel.EOF) Then
    Do While Not rsChannel.EOF
        If rsChannel("ModuleType") <> 4 Then
            If rsChannel("ModuleType") = 5 Then
                If ModuleName = "Product" Or InStr(TypeSelect, "Keyword") > 0 Then
                    If rsChannel("ChannelID") = ChannelID Then
                        ChannelList = (ChannelList & "<option value=" & rsChannel("ChannelID") & " selected>" & rsChannel("ChannelName") & "</option>")
                    Else
                        ChannelList = (ChannelList & "<option value=" & rsChannel("ChannelID") & ">" & rsChannel("ChannelName") & "</option>")
                    End If
                End If
            Else
                If rsChannel("ChannelID") = ChannelID Then
                    ChannelList = (ChannelList & "<option value=" & rsChannel("ChannelID") & " selected>" & rsChannel("ChannelName") & "</option>")
                Else
                    ChannelList = (ChannelList & "<option value=" & rsChannel("ChannelID") & ">" & rsChannel("ChannelName") & "</option>")
                End If
            End If
        End If
        rsChannel.MoveNext
    Loop
End If
rsChannel.Close
Set rsChannel = Nothing
   
strFileName = "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect
If Keyword <> "" Then
    strFileName = strFileName & "&Field=" & strField & "&keyword=" & Keyword
End If
If ItemType < 999 Then
    strFileName = strFileName & "&ItemType=" & ItemType
End If

Response.Write "<html><head><title>" & SelectedName & "</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<SCRIPT language=javascript>" & vbCrLf
Response.Write "function unselectall(){" & vbCrLf
Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function CheckAll(form){" & vbCrLf
Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
Response.Write "    var e = form.elements[i];" & vbCrLf
Response.Write "    if (e.Name != 'chkAll'&&e.disabled!=true)" & vbCrLf
Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function CheckInput(){" & vbCrLf
Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
Response.Write "  if (CurrentMode==0){" & vbCrLf
Response.Write "    document.myform.Intro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "  else if(CurrentMode==1){" & vbCrLf
Response.Write "    document.myform.Intro.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function CheckKeyLink(){" & vbCrLf
Response.Write "  if(document.myform.Source.value==''){" & vbCrLf
Response.Write "      alert('�滻Ŀ�겻��Ϊ�գ�');" & vbCrLf
Response.Write "   document.myform.Source.focus();" & vbCrLf
Response.Write "      return false;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "  if(document.myform.Target.value==''){" & vbCrLf
Response.Write "      alert('�滻���ݲ���Ϊ�գ�');" & vbCrLf
Response.Write "   document.myform.Target.focus();" & vbCrLf
Response.Write "      return false;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf

Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
If InStr(TypeSelect, "Keyword") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10015)
ElseIf InStr(TypeSelect, "Author") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10016)
ElseIf InStr(TypeSelect, "CopyFrom") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10017)
ElseIf InStr(TypeSelect, "KeyLink") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10029)
ElseIf InStr(TypeSelect, "Rtext") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10030)
ElseIf InStr(TypeSelect, "Producer") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10018)
ElseIf InStr(TypeSelect, "Trademark") > 0 Then
    Call ShowPageTitle(ChannelName & "������" & SelectedName, 10019)
End If

Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='88' height='30'><strong>��������</strong></td>" & vbCrLf
If InStr(TypeSelect, "Keyword") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword'>�ؼ��ֹ�����ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=AddKeyword'>�����ؼ���</a></td>" & vbCrLf
ElseIf InStr(TypeSelect, "Author") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author'>���߹�����ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=AddAuthor'>�������</a></td>" & vbCrLf
ElseIf InStr(TypeSelect, "CopyFrom") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom'>��Դ������ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=AddCopyFrom'>������Դ</a></td>" & vbCrLf
ElseIf InStr(TypeSelect, "KeyLink") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?TypeSelect=KeyLink'>վ�����ӹ�����ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AddKeyLink'>����վ������</a></td>" & vbCrLf
ElseIf InStr(TypeSelect, "Rtext") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?TypeSelect=Rtext'>�ַ��滻������ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AddRtext'>�����ַ��滻</a></td>" & vbCrLf
ElseIf InStr(TypeSelect, "Producer") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer'>���̹�����ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=AddProducer'>��������</a></td>" & vbCrLf
ElseIf InStr(TypeSelect, "Trademark") > 0 Then
    Response.Write "    <td height='30'><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark'>Ʒ�ƹ�����ҳ</a>&nbsp;|&nbsp;<a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=AddTrademark'>����Ʒ��</a></td>" & vbCrLf
End If

Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case TypeSelect
Case "Keyword"
    Call KeywordManage
Case "AddKeyword"
    Call AddKeyword
Case "ModifyKeyword"
    Call ModifyKeyword
Case "SaveAddKeyword"
    Call SaveAddKeyword
Case "SaveModifyKeyword"
    Call SaveModifyKeyword
Case "DelKeyword"
    Call DelKeyword
Case "DelAllKeyword"
    Call DelAllKeyword
Case "Author"
    Call Author
Case "AddAuthor"
    Call AddAuthor
Case "ModifyAuthor"
    Call ModifyAuthor
Case "SaveAddAuthor"
    Call SaveAddAuthor
Case "SaveModifyAuthor"
    Call SaveModifyAuthor
Case "DelAuthor"
    Call DelAuthor
Case "AuthorDis"
    Call SetStat("Author", 1)
Case "AuthorEn"
    Call SetStat("Author", 2)
Case "AuthorDTop"
    Call SetStat("Author", 3)
Case "AuthorTop"
    Call SetStat("Author", 4)
Case "AuthorDElite"
    Call SetStat("Author", 5)
Case "AuthorElite"
    Call SetStat("Author", 6)
Case "CopyFrom"
    Call CopyFrom
Case "AddCopyFrom"
    Call AddCopyFrom
Case "ModifyCopyFrom"
    Call ModifyCopyFrom
Case "SaveAddCopyFrom"
    Call SaveAddCopyFrom
Case "SaveModifyCopyFrom"
    Call SaveModifyCopyFrom
Case "DelCopyFrom"
    Call DelCopyFrom
Case "CopyFromDis"
    Call SetStat("CopyFrom", 1)
Case "CopyFromEn"
    Call SetStat("CopyFrom", 2)
Case "CopyFromDTop"
    Call SetStat("CopyFrom", 3)
Case "CopyFromTop"
    Call SetStat("CopyFrom", 4)
Case "CopyFromDElite"
    Call SetStat("CopyFrom", 5)
Case "CopyFromElite"
    Call SetStat("CopyFrom", 6)
Case "Producer"
    Call Producer
Case "AddProducer"
    Call AddProducer
Case "ModifyProducer"
    Call ModifyProducer
Case "SaveAddProducer"
    Call SaveAddProducer
Case "SaveModifyProducer"
    Call SaveModifyProducer
Case "DelProducer"
    Call DelProducer
Case "ProducerDis"
    Call SetStat("Producer", 1)
Case "ProducerEn"
    Call SetStat("Producer", 2)
Case "ProducerDTop"
    Call SetStat("Producer", 3)
Case "ProducerTop"
    Call SetStat("Producer", 4)
Case "ProducerDElite"
    Call SetStat("Producer", 5)
Case "ProducerElite"
    Call SetStat("Producer", 6)
Case "Trademark"
    Call Trademark
Case "AddTrademark"
    Call AddTrademark
Case "ModifyTrademark"
    Call ModifyTrademark
Case "SaveAddTrademark"
    Call SaveAddTrademark
Case "SaveModifyTrademark"
    Call SaveModifyTrademark
Case "DelTrademark"
    Call DelTrademark
Case "TrademarkDis"
    Call SetStat("Trademark", 1)
Case "TrademarkEn"
    Call SetStat("Trademark", 2)
Case "TrademarkDTop"
    Call SetStat("Trademark", 3)
Case "TrademarkTop"
    Call SetStat("Trademark", 4)
Case "TrademarkDElite"
    Call SetStat("Trademark", 5)
Case "TrademarkElite"
    Call SetStat("Trademark", 6)
Case "KeyLink"
    Call KeyLink(0)
Case "AddKeyLink"
    Call AddKeyLink(0)
Case "ModifyKeyLink"
    Call ModifyKeyLink(0)
Case "SaveAddKeyLink"
    Call SaveAddKeyLink(0)
Case "SaveModifyKeyLink"
    Call SaveModifyKeyLink(0)
Case "DelKeyLink"
    Call DelKeyLink("KeyLink")
Case "runKeyLink"
    Call SetKeyLink(0, 1)
Case "disKeyLink"
    Call SetKeyLink(0, 0)
Case "Rtext"
    Call KeyLink(1)
Case "AddRtext"
    Call AddKeyLink(1)
Case "ModifyRtext"
    Call ModifyKeyLink(1)
Case "SaveAddRtext"
    Call SaveAddKeyLink(1)
Case "SaveModifyRtext"
    Call SaveModifyKeyLink(1)
Case "DelRtext"
    Call DelKeyLink("Rtext")
Case "runRtext"
    Call SetKeyLink(1, 1)
Case "disRtext"
    Call SetKeyLink(1, 0)
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Response.Write "������ʧ��"
    Exit Sub
End Sub

'**************
'�ؼ��ִ�����
'**************

Sub KeywordManage()
    Dim rsKeyList, sqlKeyList
    Dim iCount
    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�еĹؼ�����');"">"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='40' height='22'><strong>���</strong></td>"
    Response.Write "    <td height='22'><strong>�ؼ���</strong></td>"
    Response.Write "    <td width='80' height='22'><strong>ʹ��Ƶ��</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>���ʹ��ʱ��</strong></td>"
    Response.Write "    <td width='70' height='22'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"
    
    Set rsKeyList = Server.CreateObject("Adodb.RecordSet")
    sqlKeyList = "select * from PE_NewKeys Where ChannelID=" & ChannelID
    If Keyword <> "" Then
            sqlKeyList = sqlKeyList & " and KeyText like '%" & Keyword & "%' "
    End If
    sqlKeyList = sqlKeyList & " order by LastUseTime Desc"
    rsKeyList.Open sqlKeyList, Conn, 1, 1
    If rsKeyList.BOF And rsKeyList.EOF Then
        rsKeyList.Close
        Set rsKeyList = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='6' align='center'><br>û���κιؼ��֣�<br><br></td></tr></Table>"
        Exit Sub
    End If
    
    totalPut = rsKeyList.RecordCount
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
            rsKeyList.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Do While Not rsKeyList.EOF
        Response.Write " <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsKeyList("ID") & "'"
        Response.Write " onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsKeyList("ID") & "</td>"
        Response.Write "    <td>" & GetSubStr(rsKeyList("KeyText"), 40, True) & "</td>"
        Response.Write "    <td>" & rsKeyList("Hits") & "</td>"
        Response.Write "    <td>" & rsKeyList("LastUseTime") & "</td>"
        Response.Write "<td>"
        Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=ModifyKeyword&ChannelID=" & ChannelID & "&ID=" & rsKeyList("ID") & "'>�޸�</a>&nbsp;&nbsp;"
        Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=DelKeyword&ChannelID=" & ChannelID & "&ID=" & rsKeyList("ID") & "' onClick=""return confirm('ȷ��Ҫɾ���˹ؼ�����');"">ɾ��</a>"
        Response.Write "</td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsKeyList.MoveNext
    Loop
    rsKeyList.Close
    Set rsKeyList = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ�����йؼ���</td>"
    Response.Write "    <td><input name='TypeSelect' type='hidden' id='TypeSelect' value='DelKeyword'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�еĹؼ���'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='submit' name='Submit2' value='ɾ����Ƶ��ȫ���ؼ���' onClick=""document.myform.TypeSelect.value='DelAllKeyword'""></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "���ؼ���", True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>�ؼ���������</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SourceManage.asp'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='" & TypeSelect & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>�ؼ�����</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "�ؼ���"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub AddKeyword()
    Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� ��</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='100%' align='center' class='tdbg'><strong> �� �� �֣�</strong><input name='KeyText' type='text'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='100%' align='center' class='tdbg'><strong> ����Ƶ����</strong><select name='ChannelID'>" & ChannelList & "<option value=0"
    If ChannelID = 0 Then
        Response.Write " selected>ȫ��Ƶ��</option></select></td>"
    Else
        Response.Write ">ȫ��Ƶ��</option></select></td>"
    End If
    Response.Write "    </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
    Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddKeyword'>"
    Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword';"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
End Sub

Sub ModifyKeyword()
    Dim KeyID
    Dim rsKey, sqlKey
    KeyID = PE_CLng(Trim(Request("ID")))
    If KeyID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹؼ���ID</li>"
        Exit Sub
    End If
    sqlKey = "Select * from PE_NewKeys where ID=" & KeyID
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    rsKey.Open sqlKey, Conn, 1, 1
    If rsKey.BOF And rsKey.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˹ؼ��֣�</li>"
    Else
        Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� �� �� ��</strong></font></div></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr> "
        Response.Write "      <td width='100%' class='tdbg' align='center'><strong>�� �� �� ��</strong><input name='KeyText' type='text' value='" & rsKey("KeyText") & "'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td width='100%' class='tdbg' align='center'><strong>����Ƶ����</strong><select name='ChannelID'>" & ChannelList & "<option value=0"
        If ChannelID = 0 Then Response.Write " selected"
        Response.Write ">ȫ��Ƶ��</option></select>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td colspan='2' align='center' class='tdbg'>"
        Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyKeyword'>"
        Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsKey("ID") & ">"
        Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsKey.Close
    Set rsKey = Nothing
End Sub


Sub SaveAddKeyword()
    Dim KeyText
    Dim rsKey, sqlKey
    
    KeyText = Trim(Request("KeyText"))
    If KeyText = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ؼ��ֲ���Ϊ�գ�</li>"
    Else
        KeyText = ReplaceBadChar(KeyText)
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    sqlKey = "Select * from PE_NewKeys where ChannelID=" & ChannelID & " and KeyText='" & KeyText & "'"
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    rsKey.Open sqlKey, Conn, 1, 3
    If Not (rsKey.BOF And rsKey.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ˹ؼ��֣�</li>"
        rsKey.Close
        Set rsKey = Nothing
        Exit Sub
    End If
    rsKey.addnew
    rsKey("ChannelID") = ChannelID
    rsKey("KeyText") = KeyText
    rsKey("Hits") = 0
    rsKey("LastUseTime") = Now()
    rsKey.Update
    rsKey.Close
    Set rsKey = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword"
End Sub

Sub SaveModifyKeyword()
    Dim KeyText, KeyID
    Dim rsKey, sqlKey
    KeyText = Trim(Request("KeyText"))
    If KeyText = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�ؼ��ֲ���Ϊ�գ�</li>"
    End If
    KeyID = PE_CLng(Trim(Request("ID")))
    If KeyID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹؼ���ID</li>"
        Exit Sub
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    sqlKey = "Select ChannelID,KeyText from PE_NewKeys where ID=" & KeyID
    rsKey.Open sqlKey, Conn, 1, 3
    If Not (rsKey.BOF And rsKey.EOF) Then
        If rsKey("ChannelID") = ChannelID And ChannelID > 0 Then
            Conn.Execute ("update PE_" & ModuleName & " set Keyword='|" & KeyText & "|' where ChannelID=" & ChannelID & " and Keyword = '|" & rsKey("KeyText") & "|'")
        End If
        rsKey("ChannelID") = ChannelID
        rsKey("KeyText") = KeyText
        rsKey.Update
    End If
    rsKey.Close
    Set rsKey = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword"
End Sub

Sub DelKeyword()
    Dim KeyID
    KeyID = Trim(Request("ID"))
    If IsValidID(KeyID) = False Then
        KeyID = ""
    End If
    If KeyID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ���Ĺؼ���ID</li>"
        Exit Sub
    End If
    If InStr(KeyID, ",") > 0 Then
        Conn.Execute ("delete from PE_NewKeys where ID in (" & KeyID & ")")
    Else
        Conn.Execute ("delete from PE_NewKeys where ID=" & KeyID & "")
    End If
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword"
End Sub

Sub DelAllKeyword()
    Conn.Execute ("delete from PE_NewKeys where ChannelID=" & ChannelID)
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword"
End Sub

'**************
'���ߴ�����
'**************

Sub Author()
    Dim rsAuthor, sqlAuthor, rsChannelAuthor
    Dim iCount
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | "
    If ChannelID = 0 Then
        Response.Write "<font color='red'>ȫվ����</font>"
    Else
        Response.Write "<a href='Admin_SourceManage.asp?ChannelID=0&TypeSelect=Author&ItemType=" & ItemType & "'>ȫվ����</a>"
    End If
    Set rsChannelAuthor = Conn.Execute("select ChannelID,ChannelName from PE_Channel Where ModuleType in (1,2,3) and Disabled=" & PE_False & " order by OrderID")
    Do While Not rsChannelAuthor.EOF
        If rsChannelAuthor("ChannelID") = ChannelID Then
            Response.Write " | <font color='red'>" & rsChannelAuthor("ChannelName") & "</font>"
        Else
            Response.Write " | <a href='Admin_SourceManage.asp?ChannelID=" & rsChannelAuthor("ChannelID") & "&TypeSelect=Author&ItemType=" & ItemType & "'>" & rsChannelAuthor("ChannelName") & "</a>"
        End If
        rsChannelAuthor.MoveNext
    Loop
    Set rsChannelAuthor = Nothing
    Response.Write " |</td></tr></table>"

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author&ItemType=1'>"
    If ItemType = 1 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowAuthor/AuthorType1", "��½����") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowAuthor/AuthorType1", "��½����")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author&ItemType=2'>"
    If ItemType = 2 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowAuthor/AuthorType2", "��̨����") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowAuthor/AuthorType2", "��̨����")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author&ItemType=3'>"
    If ItemType = 3 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowAuthor/AuthorType3", "��������") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowAuthor/AuthorType3", "��������")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author&ItemType=4'>"
    If ItemType = 4 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowAuthor/AuthorType4", "��վ��Լ") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowAuthor/AuthorType4", "��վ��Լ")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author&ItemType=0'>"
    If ItemType = 0 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowAuthor/AuthorType5", "��������") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowAuthor/AuthorType5", "��������")
    End If

    Response.Write "</a> |</td></tr></table><br>"
    Response.Write "  <form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�е�������');"">"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title' height='22'>"
    Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='40'><strong>���</strong></td>"
    Response.Write "    <td width='80'><strong>����</strong></td>"
    Response.Write "    <td width='40' height='22'><strong>�Ա�</strong></td>"
    Response.Write "    <td height='22'><strong>���</strong></td>"
    Response.Write "    <td width='80' height='22'><strong>���߷���</strong></td>"
    Response.Write "    <td width='60' height='22'><strong>״̬</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"
 
    '�������ģ���ֶ��Ƿ����
    Dim dbrr, i
    dbrr = False
    Set rsAuthor = Conn.Execute("select top 1 * from PE_Author")
    For i = 0 To rsAuthor.Fields.Count - 1
        If rsAuthor.Fields(i).name = "TemplateID" Then
            dbrr = True
        End If
    Next
    rsAuthor.Close
    Set rsAuthor = Nothing
    If dbrr <> True Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table [PE_Author] add TemplateID int DEFAULT (0)")
        Else
            Conn.Execute ("alter table [PE_Author] add COLUMN TemplateID int 0")
        End If
    End If

    Set rsAuthor = Server.CreateObject("Adodb.RecordSet")
    sqlAuthor = "select * from PE_Author Where ChannelID=" & ChannelID
    If Keyword <> "" Then
        Select Case strField
        Case "name"
            sqlAuthor = sqlAuthor & " and AuthorName like '%" & Keyword & "%' "
        Case "address"
            sqlAuthor = sqlAuthor & " and Address like '%" & Keyword & "%' "
        Case "Phone"
            sqlAuthor = sqlAuthor & " and Tel like '%" & Keyword & "%' "
        Case "intro"
            sqlAuthor = sqlAuthor & " and Intro like '%" & Keyword & "%' "
        Case Else
            sqlAuthor = sqlAuthor & " and AuthorName like '%" & Keyword & "%' "
        End Select
    End If
    If ItemType < 999 Then
        sqlAuthor = sqlAuthor & " and AuthorType =" & ItemType
    End If
    sqlAuthor = sqlAuthor & " order by ID Desc"
    rsAuthor.Open sqlAuthor, Conn, 1, 1
    If rsAuthor.BOF And rsAuthor.EOF Then
        rsAuthor.Close
        Set rsAuthor = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='10' align='center'><br>û���κ����ߣ�<br><br></td></tr></table>"
        Exit Sub
    End If
    
    totalPut = rsAuthor.RecordCount
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
            rsAuthor.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
  
    Do While Not rsAuthor.EOF
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsAuthor("ID") & "'  onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsAuthor("ID") & "</td>"
        Response.Write "    <td>" & rsAuthor("AuthorName") & "</td>"
        Response.Write "    <td>" & GetSex(rsAuthor("Sex")) & "</td>"
        Response.Write "    <td>" & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), 24, False) & "</td>"
        Response.Write "    <td>" & GetAuthorType(rsAuthor("AuthorType")) & "</td><td>"
        If rsAuthor("Passed") = True Then
            Response.Write "<font color=""green"">��</font>"
        Else
            Response.Write "<font color=""red"">��</font>"
        End If
        If rsAuthor("onTop") = True Then
            Response.Write "&nbsp;<font color=""blue"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        If rsAuthor("isElite") = True Then
            Response.Write "&nbsp;<font color=""green"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td><td>"
        Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=ModifyAuthor&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>�޸�</a>"
        If rsAuthor("Passed") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AuthorDis&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>����</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AuthorEn&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>����</a>"
        End If
        If rsAuthor("onTop") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AuthorDTop&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AuthorTop&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>�̶�</a>"
        End If
        If rsAuthor("isElite") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AuthorDElite&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=AuthorElite&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "'>�Ƽ�</a>"
        End If
        Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=DelAuthor&ChannelID=" & ChannelID & "&ID=" & rsAuthor("ID") & "' onClick=""return confirm('ȷ��Ҫɾ������" & rsAuthor("AuthorName") & "��');"">ɾ��</a>"
        Response.Write "</td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsAuthor.MoveNext
    Loop
    rsAuthor.Close
   
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ����������</td>"
    Response.Write "    <td><input name='TypeSelect' type='hidden' id='TypeSelect' value='DelAuthor'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�е�����'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������", True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>����������</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SourceManage.asp'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='" & TypeSelect & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>������</option><option value='address'>���ߵ�ַ</option><option value='Phone'>���ߵ绰</option><option value='intro'>���߼��</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "�ؼ���"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub AddAuthor()
    Call PopCalendarInit
    Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
    Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� Ϣ</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong> ������</strong><input name='AuthorName' type='text' size='20' maxlength='20'> <font color='#FF0000'>*</font></td>"
    Response.Write "    <td rowspan='8' align='center' valign='top' class='tdbg'>"
    Response.Write "        <table width='180' height='200' border='1'>"
    Response.Write "            <tr><td width='100%' align='center'><img id='showphoto' src='" & InstallDir & "AuthorPic/default.gif' width='150' height='172'></td></tr>"
    Response.Write "        </table>"
    Response.Write "        <input name='Photo' type='text' size='25'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AuthorPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
    Response.Write "     </td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> Ƶ����</strong><select name='ChannelID'>" & ChannelList & "</select></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �Ա�</strong><input name='Sex' type='radio' value='1' checked>��&nbsp;&nbsp;<input type='radio' name='Sex' value='0'>Ů</td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ���գ�</strong><input name='BirthDay' type='text' size='20' maxlength='20' value='" & FormatDateTime(Date, 2) & "' maxlength='20'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BirthDay, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��ַ��</strong><input name='Address' type='text' size='20' maxlength='20'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �绰��</strong><input name='Tel' type='text' size='20' maxlength='20'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ���棺</strong><input name='Fax' type='text' size='20' maxlength='20'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��λ��</strong><input name='Company' type='text' size='20' maxlength='20'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ���ţ�</strong><input name='Department' type='text' size='20' maxlength='20'></td><td><strong> ģ�壺</strong><select name='TemplateID'>" & AuthorTemplateList(0) & "</select></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �ʱࣺ</strong><input name='ZipCode' type='text' size='20' maxlength='20'></td><td><strong> ��ҳ��</strong><input name='HomePage' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �ʼ���</strong><input name='Email' type='text' size='20' maxlength='20'></td><td><strong> �ѣѣ�</strong><input name='QQ' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>���߷��ࣺ</strong><input name='AuthorType' type='radio' value='1' checked>" & XmlText("ShowSource", "ShowAuthor/AuthorType1", "��½����") & "&nbsp;<input name='AuthorType' type='radio' value='2'>" & XmlText("ShowSource", "ShowAuthor/AuthorType2", "��̨����") & "&nbsp;<input name='AuthorType' type='radio' value='3'>" & XmlText("ShowSource", "ShowAuthor/AuthorType3", "��������") & "&nbsp;<input name='AuthorType' type='radio' value='4'>" & XmlText("ShowSource", "ShowAuthor/AuthorType4", "��վ��Լ") & "&nbsp;<input name='AuthorType' type='radio' value='0'>" & XmlText("ShowSource", "ShowAuthor/AuthorType5", "��������") & "&nbsp;</td></tr>"
    Response.Write "  <tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>���߼��</strong>��<br>"
    Response.Write "      <textarea name='Intro' id='Intro' cols='72' rows='9' style='display:none'></textarea>"
    Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='300' ></iframe>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
    Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddAuthor'>"
    Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author';"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
End Sub

Sub ModifyAuthor()
    Dim AuthorID
    Dim rsAuthor, sqlAuthor
    AuthorID = PE_CLng(Trim(Request("ID")))
    If AuthorID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�����ID</li>"
        Exit Sub
    End If
    sqlAuthor = "Select * from PE_Author where ID=" & AuthorID
    Set rsAuthor = Server.CreateObject("Adodb.RecordSet")
    rsAuthor.Open sqlAuthor, Conn, 1, 1
    If rsAuthor.BOF And rsAuthor.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ����ߣ�</li>"
    Else
        Call PopCalendarInit
        Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='return CheckInput();'>"
        Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� �� �� �� Ϣ</strong></font></div></td>"
        Response.Write "    </tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong> ������</strong><input name='AuthorName' type='text' value='" & rsAuthor("AuthorName") & "'> <font color='#FF0000'>*</font></td>"
        Response.Write "    <td rowspan='8' align='center' valign='top' class='tdbg'>"
        Response.Write "        <table width='180' height='200' border='1'>"
        Response.Write "            <tr><td width='100%' align='center'>"
        If IsNull(rsAuthor("Photo")) Then
            Response.Write "<img id='showphoto' src='" & InstallDir & "AuthorPic/default.gif' width='150' height='172'>"
        Else
            Response.Write "<img id='showphoto' src='" & rsAuthor("Photo") & "' width='150' height='172'>"
        End If
        Response.Write "        </td></tr></table>"
        Response.Write "        <input name='Photo' type='text' size='25' value='" & rsAuthor("Photo") & "'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AuthorPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
        Response.Write "     </td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> Ƶ����</strong><select name='ChannelID'>" & ChannelList & "</select></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �Ա�</strong><input name='Sex' type='radio' value='1'"
    If rsAuthor("Sex") = 1 Then Response.Write " Checked"
        Response.Write ">��&nbsp;&nbsp;<input type='radio' name='Sex' value='0'"
    If rsAuthor("Sex") = 0 Then Response.Write " Checked"
        Response.Write ">Ů</td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ���գ�</strong><input name='BirthDay' type='text'  value='" & rsAuthor("BirthDAy") & "' maxlength='20'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BirthDay, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��ַ��</strong><input name='Address' type='text'  value='" & rsAuthor("Address") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �绰��</strong><input name='Tel' type='text' value='" & rsAuthor("Tel") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ���棺</strong><input name='Fax' type='text' value='" & rsAuthor("Fax") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��λ��</strong><input name='Company' type='text' value='" & rsAuthor("Company") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ���ţ�</strong><input name='Department' type='text' value='" & rsAuthor("Department") & "'></td><td><strong> ģ�壺</strong><select name='TemplateID'>" & AuthorTemplateList(rsAuthor("TemplateID")) & "</select></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �ʱࣺ</strong><input name='ZipCode' type='text' value='" & rsAuthor("ZipCode") & "'></td><td><strong> ��ҳ��</strong><input name='HomePage' type='text' value='" & rsAuthor("HomePage") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �ʼ���</strong><input name='Email' type='text' value='" & rsAuthor("Email") & "'></td><td><strong> �ѣѣ�</strong><input name='QQ' type='text' value='" & rsAuthor("QQ") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>���߷��ࣺ</strong>"
        Response.Write "<input name='AuthorType' type='radio' value='1'"
    If rsAuthor("AuthorType") = 1 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType1", "��½����") & "&nbsp;<input name='AuthorType' type='radio' value='2'"
    If rsAuthor("AuthorType") = 2 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType2", "��̨����") & "&nbsp;<input name='AuthorType' type='radio' value='3'"
    If rsAuthor("AuthorType") = 3 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType3", "��������") & "&nbsp;<input name='AuthorType' type='radio' value='4'"
    If rsAuthor("AuthorType") = 4 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType4", "��վ��Լ") & "&nbsp;<input name='AuthorType' type='radio' value='0'"
    If rsAuthor("AuthorType") = 0 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType5", "��������") & "&nbsp;</td></tr>"
        Response.Write "  <tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>���߼��</strong>��<br>"
        Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>"
        If Trim(rsAuthor("Intro") & "") <> "" Then Response.Write Server.HTMLEncode(rsAuthor("Intro"))
        Response.Write "      </textarea>"
        Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td colspan='2' align='center' class='tdbg'>"
        Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyAuthor'>"
        Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsAuthor("ID") & ">"
        Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsAuthor.Close
    Set rsAuthor = Nothing
End Sub

Sub SaveAddAuthor()
    Dim AuthorName, Sex, Birthday, Address, Tel, Fax, Company, Department, ZipCode, Homepage, Email, QQ, Intro, Photo, AuthorType
    Dim rsAuthor, sqlAuthor
    AuthorName = Trim(Request("AuthorName"))

    If AuthorName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������������Ϊ�գ�</li>"
    Else
        AuthorName = ReplaceBadChar(AuthorName)
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    Sex = PE_CLng(Request("Sex"))
    Birthday = Trim(Request("BirthDay"))
    Photo = Trim(Request("Photo"))
    Address = Trim(Request("Address"))
    Tel = Trim(Request("Tel"))
    Fax = Trim(Request("Fax"))
    Company = Trim(Request("Company"))
    Department = Trim(Request("Department"))
    ZipCode = Trim(Request("ZipCode"))
    Homepage = Trim(Request("HomePage"))
    Email = Trim(Request("Email"))
    QQ = Trim(Request("QQ"))
    Intro = Trim(Request("Intro"))
    AuthorType = Trim(Request("AuthorType"))
    Set rsAuthor = Server.CreateObject("Adodb.RecordSet")
    sqlAuthor = "Select * from PE_Author where ChannelID=" & ChannelID & " and AuthorName='" & AuthorName & "'"
    rsAuthor.Open sqlAuthor, Conn, 1, 3
    If Not (rsAuthor.BOF And rsAuthor.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ����ߣ�</li>"
        rsAuthor.Close
        Set rsAuthor = Nothing
        Exit Sub
    End If
    rsAuthor.addnew
    rsAuthor("ChannelID") = ChannelID
    rsAuthor("AuthorName") = AuthorName
    rsAuthor("Sex") = Sex
    If Birthday <> "" Then rsAuthor("BirthDay") = Birthday
    If Address <> "" Then rsAuthor("Address") = Address
    If Tel <> "" Then rsAuthor("Tel") = Tel
    If Fax <> "" Then rsAuthor("Fax") = Fax
    If Company <> "" Then rsAuthor("Company") = Company
    If Department <> "" Then rsAuthor("Department") = Department
    If ZipCode <> "" Then rsAuthor("ZipCode") = ZipCode
    If Homepage <> "" Then rsAuthor("HomePage") = Homepage
    If Email <> "" Then rsAuthor("Email") = Email
    If QQ <> "" Then rsAuthor("QQ") = PE_CLng(QQ)
    If Intro <> "" Then rsAuthor("Intro") = Intro
    If Photo <> "" Then rsAuthor("Photo") = Photo
    rsAuthor("AuthorType") = PE_CLng(AuthorType)
    rsAuthor("LastUseTime") = Now()
    rsAuthor("Passed") = True
    rsAuthor("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
    rsAuthor.Update
    rsAuthor.Close
    Set rsAuthor = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author"
End Sub

Sub SaveModifyAuthor()
    Dim AuthorName, AuthorID, Sex, Birthday, Address, Tel, Fax, Company, Department, ZipCode, Homepage, Email, QQ, Intro, Photo, AuthorType
    Dim rsAuthor, sqlAuthor
    AuthorName = Trim(Request("AuthorName"))
    If AuthorName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�����������</li>"
    End If
    AuthorID = Trim(Request("ID"))
    If AuthorID <> "" Then
        If InStr(AuthorID, ",") > 0 Then
            AuthorID = ReplaceBadChar(AuthorID)
        Else
            AuthorID = PE_CLng(AuthorID)
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�����ID��</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    Sex = PE_CLng(Request("Sex"))
    Birthday = Trim(Request("BirthDay"))
    Photo = Trim(Request("Photo"))
    Address = Trim(Request("Address"))
    Tel = Trim(Request("Tel"))
    Fax = Trim(Request("Fax"))
    Company = Trim(Request("Company"))
    Department = Trim(Request("Department"))
    ZipCode = Trim(Request("ZipCode"))
    Homepage = Trim(Request("HomePage"))
    Email = Trim(Request("Email"))
    QQ = Trim(Request("QQ"))
    Intro = Trim(Request("Intro"))
    AuthorType = Trim(Request("AuthorType"))
    Set rsAuthor = Server.CreateObject("Adodb.RecordSet")
    sqlAuthor = "Select * from PE_Author where ID=" & AuthorID
    rsAuthor.Open sqlAuthor, Conn, 1, 3
    If Not (rsAuthor.BOF And rsAuthor.EOF) Then
        rsAuthor("ChannelID") = ChannelID
        If AuthorName <> "" Then rsAuthor("AuthorName") = AuthorName
        rsAuthor("Sex") = Sex
        If Birthday <> "" Then rsAuthor("BirthDay") = Birthday
        rsAuthor("Intro") = Intro
        rsAuthor("Address") = Address
        rsAuthor("Tel") = Tel
        rsAuthor("Fax") = Fax
        rsAuthor("Company") = Company
        rsAuthor("Department") = Department
        rsAuthor("ZipCode") = ZipCode
        rsAuthor("HomePage") = Homepage
        rsAuthor("Email") = Email
        rsAuthor("QQ") = PE_CLng(QQ)
        If Photo <> "" Then rsAuthor("Photo") = Photo
        rsAuthor("AuthorType") = PE_CLng(AuthorType)
        rsAuthor("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
        rsAuthor.Update
    End If
    rsAuthor.Close
    Set rsAuthor = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author"
End Sub

Sub DelAuthor()
    Dim AuthorID
    AuthorID = Trim(Request("ID"))
    If IsValidID(AuthorID) = False Then
        AuthorID = ""
    End If
    If AuthorID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ��������ID</li>"
        Exit Sub
    End If
    If InStr(AuthorID, ",") > 0 Then
        Conn.Execute ("delete from PE_Author where ID in (" & AuthorID & ")")
    Else
        Conn.Execute ("delete from PE_Author where ID=" & AuthorID & "")
    End If
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author"
End Sub

Function GetAuthorType(TypeID)
     '1Ϊ���� 2Ϊ��̨ 3Ϊ���� 4Ϊ��վ��Լ 0Ϊ����
    Select Case TypeID
    Case 1
        GetAuthorType = XmlText("ShowSource", "ShowAuthor/AuthorType1", "��½����")
    Case 2
        GetAuthorType = XmlText("ShowSource", "ShowAuthor/AuthorType2", "��̨����")
    Case 3
        GetAuthorType = XmlText("ShowSource", "ShowAuthor/AuthorType3", "��������")
    Case 4
        GetAuthorType = XmlText("ShowSource", "ShowAuthor/AuthorType4", "��վ��Լ")
    Case Else
        GetAuthorType = XmlText("ShowSource", "ShowAuthor/AuthorType5", "��������")
    End Select
End Function

Function GetSex(SexID)
    If SexID = "" Or SexID = 0 Then
        GetSex = XmlText("BaseText", "Girl", "Ů")
    Else
        GetSex = XmlText("BaseText", "Man", "��")
    End If
End Function


'**************
'��Դ������
'**************

Sub CopyFrom()
    Dim rsCopyFrom, sqlCopyFrom, rsChannelCopyFrom
    Dim iCount
   
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | "
    If ChannelID = 0 Then
        Response.Write "<font color='red'>ȫվ��Դ</font>"
    Else
        Response.Write "<a href='Admin_SourceManage.asp?ChannelID=0&TypeSelect=CopyFrom&ItemType=" & ItemType & "'>ȫվ��Դ</a>"
    End If
    Set rsChannelCopyFrom = Conn.Execute("select ChannelID,ChannelName from PE_Channel Where ModuleType in (1,2,3) and Disabled=" & PE_False & " order by OrderID")
    Do While Not rsChannelCopyFrom.EOF
        If rsChannelCopyFrom("ChannelID") = ChannelID Then
            Response.Write " | <font color='red'>" & rsChannelCopyFrom("ChannelName") & "</font>"
        Else
            Response.Write " | <a href='Admin_SourceManage.asp?ChannelID=" & rsChannelCopyFrom("ChannelID") & "&TypeSelect=CopyFrom&ItemType=" & ItemType & "'>" & rsChannelCopyFrom("ChannelName") & "</a>"
        End If
        rsChannelCopyFrom.MoveNext
    Loop
    Set rsChannelCopyFrom = Nothing
    Response.Write " |</td></tr></table>"

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom&ItemType=1'>"
    If ItemType = 1 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "����վ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "����վ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom&ItemType=2'>"
    If ItemType = 2 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "����վ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "����վ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom&ItemType=3'>"
    If ItemType = 3 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "����վ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "����վ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom&ItemType=0'>"
    If ItemType = 4 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "������Դ") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "������Դ")
    End If
    Response.Write "</a> |</td></tr></table><br>"
    Response.Write "  <form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�е���Դ��');"">"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='40' height='22'><strong>���</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>����</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>��ַ</strong></td>"
    Response.Write "    <td height='22'><strong>���</strong></td>"
    Response.Write "    <td width='80' height='22'><strong>��Դ����</strong></td>"
    Response.Write "    <td width='60' height='22'><strong>״̬</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"
    
    Set rsCopyFrom = Server.CreateObject("Adodb.RecordSet")
    sqlCopyFrom = "select * from PE_CopyFrom Where ChannelID=" & ChannelID
    If Keyword <> "" Then
        Select Case strField
        Case "name"
            sqlCopyFrom = sqlCopyFrom & " and SourceName like '%" & Keyword & "%' "
        Case "address"
            sqlCopyFrom = sqlCopyFrom & " and Address like '%" & Keyword & "%' "
        Case "Phone"
            sqlCopyFrom = sqlCopyFrom & " and Tel like '%" & Keyword & "%' "
        Case "intro"
            sqlCopyFrom = sqlCopyFrom & " and Intro like '%" & Keyword & "%' "
        Case "ContacterName"
            sqlCopyFrom = sqlCopyFrom & " and ContacterName like '%" & Keyword & "%' "
        Case Else
            sqlCopyFrom = sqlCopyFrom & " and SourceName like '%" & Keyword & "%' "
        End Select
    End If
    If ItemType < 999 Then
        sqlCopyFrom = sqlCopyFrom & " and SourceType =" & ItemType
    End If
    sqlCopyFrom = sqlCopyFrom & " order by ID Desc"
    
    rsCopyFrom.Open sqlCopyFrom, Conn, 1, 1
    If rsCopyFrom.BOF And rsCopyFrom.EOF Then
        rsCopyFrom.Close
        Set rsCopyFrom = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='8' align='center'><br>û���κ���Դ��<br><br></td></tr>"
        Response.Write "</Table>"
        Exit Sub
    End If
    
    totalPut = rsCopyFrom.RecordCount
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
            rsCopyFrom.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Do While Not rsCopyFrom.EOF
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsCopyFrom("ID") & "'"
        Response.Write " onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsCopyFrom("ID") & "</td>"
        Response.Write "    <td>" & GetSubStr(rsCopyFrom("SourceName"), 24, True) & "</td>"
        If rsCopyFrom("Address") <> "" Then
            Response.Write "    <td>" & GetSubStr(rsCopyFrom("Address"), 24, True) & "</td>"
        Else
            Response.Write "    <td>" & rsCopyFrom("Address") & "</td>"
        End If
        If rsCopyFrom("Intro") <> "" Then
            Response.Write "    <td> " & GetSubStr(nohtml(PE_HtmlDecode(rsCopyFrom("Intro"))), 30, False)
            If Len(rsCopyFrom("Intro")) > 32 Then Response.Write "��"
            Response.Write "</td>"
        Else
            Response.Write "    <td>" & rsCopyFrom("Intro") & "</td>"
        End If
        Response.Write "    <td>" & GetCopyFromType(rsCopyFrom("SourceType")) & "</td><td>"
        If rsCopyFrom("Passed") = True Then
            Response.Write "<font color=""green"">��</font>"
        Else
            Response.Write "<font color=""red"">��</font>"
        End If
        If rsCopyFrom("onTop") = True Then
            Response.Write "&nbsp;<font color=""blue"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        If rsCopyFrom("isElite") = True Then
            Response.Write "&nbsp;<font color=""green"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td><td>"
        Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=ModifyCopyFrom&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>�޸�</a>"
        If rsCopyFrom("Passed") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=CopyFromDis&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>����</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=CopyFromEn&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>����</a>"
        End If
        If rsCopyFrom("onTop") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=CopyFromDTop&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=CopyFromTop&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>�̶�</a>"
        End If
        If rsCopyFrom("isElite") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=CopyFromDElite&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=CopyFromElite&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "'>�Ƽ�</a>"
        End If
        Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=DelCopyFrom&ChannelID=" & ChannelID & "&ID=" & rsCopyFrom("ID") & "' onClick=""return confirm('ȷ��Ҫɾ����Դ" & rsCopyFrom("SourceName") & "��');"">ɾ��</a>"
        Response.Write "</td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsCopyFrom.MoveNext
    Loop
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ����������</td>"
    Response.Write "    <td><input name='TypeSelect' type='hidden' id='TypeSelect' value='DelCopyFrom'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�е���Դ'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����Դ", True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>��Դ������</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SourceManage.asp'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='" & TypeSelect & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>��Դ����</option><option value='address'>��Դ��ַ</option><option value='Phone'>��Դ�绰</option><option value='intro'>��Դ���</option><option value='ContacterName'>��ϵ��</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "�ؼ���"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub AddCopyFrom()
    Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
    Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� Դ �� Ϣ</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong> ��Դ���ƣ�</strong><input name='SourceName' type='text'> <font color='#FF0000'>*</font></td>"
    Response.Write "    <td rowspan='9' align='center' valign='top' class='tdbg'>"
    Response.Write "        <table width='180' height='200' border='1'>"
    Response.Write "            <tr><td width='100%' align='center'><img id='showphoto' src='" & InstallDir & "CopyFromPic/default.gif' width='150' height='172'></td></tr>"
    Response.Write "        </table>"
    Response.Write "        <input name='Photo' type='text' size='25'><strong>��ͼ Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=CopyFromPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
    Response.Write "     </td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ����Ƶ����</strong><select name='ChannelID'>" & ChannelList & "</select></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �� ϵ �ˣ�</strong><input name='ContacterName' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��λ��ַ��</strong><input name='Address' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �绰���룺</strong><input name='Tel' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������룺</strong><input name='Fax' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������䣺</strong><input name='Mail' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��λ���ƣ�</strong><input name='Company' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��ϵ���ţ�</strong><input name='Department' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������룺</strong><input name='ZipCode' type='text'></td><td><strong> ��λ��ҳ��</strong><input name='HomePage' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �����ʼ���</strong><input name='Email' type='text'></td><td><strong> ��ϵ�ѣѣ�</strong><input name='QQ' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>��Դ���ࣺ</strong><input name='SourceType' type='radio' value='1' checked>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "����վ��") & "&nbsp;<input name='SourceType' type='radio' value='2'>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "����վ��") & "&nbsp;<input name='SourceType' type='radio' value='3'>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "����վ��") & "&nbsp;<input name='SourceType' type='radio' value='0'>" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "������Դ") & "&nbsp;</td></tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>���</strong>��<br>"
    Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'></textarea>"
    Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
    Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddCopyFrom'>"
    Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom';"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
End Sub

Sub ModifyCopyFrom()
    Dim CopyFromID
    Dim rsCopyFrom, sqlCopyFrom
    CopyFromID = PE_CLng(Trim(Request("ID")))
    If CopyFromID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�����ID</li>"
        Exit Sub
    End If
    sqlCopyFrom = "Select * from PE_CopyFrom where ID=" & CopyFromID
    Set rsCopyFrom = Server.CreateObject("Adodb.RecordSet")
    rsCopyFrom.Open sqlCopyFrom, Conn, 1, 1
    If rsCopyFrom.BOF And rsCopyFrom.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ���Դ��</li>"
    Else
        Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='return CheckInput();'>"
        Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� �� Դ �� Ϣ</strong></font></div></td>"
        Response.Write "    </tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong> ��Դ���ƣ�</strong><input name='SourceName' type='text' value='" & rsCopyFrom("SourceName") & "'> <font color='#FF0000'>*</font></td>"
        Response.Write "    <td rowspan='9' align='center' valign='top' class='tdbg'>"
        Response.Write "        <table width='180' height='200' border='1'>"
        Response.Write "            <tr><td width='100%' align='center'>"
        If IsNull(rsCopyFrom("Photo")) Then
            Response.Write "<img id='showphoto' src='" & InstallDir & "CopyFromPic/default.gif' width='150' height='172'>"
        Else
            Response.Write "<img id='showphoto' src='" & rsCopyFrom("Photo") & "' width='150' height='172'>"
        End If
        Response.Write "        </td></tr></table>"
        Response.Write "        <input name='Photo' type='text' size='25'><strong>��ͼ Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=CopyFromPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
        Response.Write "     </td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ����Ƶ����</strong><select name='ChannelID'>" & ChannelList & "</select></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �� ϵ �ˣ�</strong><input name='ContacterName' type='text' value='" & rsCopyFrom("ContacterName") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��λ��ַ��</strong><input name='Address' type='text' value='" & rsCopyFrom("Address") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �绰���룺</strong><input name='Tel' type='text' value='" & rsCopyFrom("Tel") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������룺</strong><input name='Fax' type='text' value='" & rsCopyFrom("Fax") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������䣺</strong><input name='Mail' type='text' value='" & rsCopyFrom("Mail") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������룺</strong><input name='ZipCode' type='text' value='" & rsCopyFrom("ZipCode") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �����ʼ���</strong><input name='Email' type='text' value='" & rsCopyFrom("Email") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��λ��ҳ��</strong><input name='HomePage' type='text' value='" & rsCopyFrom("HomePage") & "'></td><td><strong> ��ϵ�ѣѣ�</strong><input name='QQ' type='text' value='" & rsCopyFrom("QQ") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>��Դ���ࣺ</strong>"
        Response.Write "<input name='SourceType' type='radio' value='1'"
        If rsCopyFrom("SourceType") = 1 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "����վ��") & "&nbsp;<input name='SourceType' type='radio' value='2'"
        If rsCopyFrom("SourceType") = 2 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "����վ��") & "&nbsp;<input name='SourceType' type='radio' value='3'"
        If rsCopyFrom("SourceType") = 3 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "����վ��") & "&nbsp;<input name='SourceType' type='radio' value='0'"
        If rsCopyFrom("SourceType") = 4 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "������Դ") & "&nbsp;</td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>���</strong>��<br>"
        Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>"
        If Trim(rsCopyFrom("Intro") & "") <> "" Then Response.Write Server.HTMLEncode(rsCopyFrom("Intro"))
        Response.Write "      </textarea>"
        Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td colspan='2' align='center' class='tdbg'>"
        Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyCopyFrom'>"
        Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsCopyFrom("ID") & ">"
        Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
End Sub


Sub SaveAddCopyFrom()
    Dim SourceName, Address, Tel, Fax, ContacterName, ZipCode, Homepage, Mail, Email, QQ, Intro, Photo, SourceType
    Dim rsCopyFrom, sqlCopyFrom
    
    SourceName = Trim(Request("SourceName"))
  
    If SourceName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Դ���Ʋ���Ϊ�գ�</li>"
    Else
        SourceName = ReplaceBadChar(SourceName)
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    
    Photo = Trim(Request("Photo"))
    Address = Trim(Request("Address"))
    Tel = Trim(Request("Tel"))
    Fax = Trim(Request("Fax"))
    ContacterName = Trim(Request("ContacterName"))
    Mail = Trim(Request("Mail"))
    Photo = Trim(Request("Photo"))
    ZipCode = Trim(Request("ZipCode"))
    Homepage = Trim(Request("HomePage"))
    Email = Trim(Request("Email"))
    QQ = Trim(Request("QQ"))
    Intro = Trim(Request("Intro"))
    SourceType = Trim(Request("SourceType"))
    
    Set rsCopyFrom = Server.CreateObject("Adodb.RecordSet")
    sqlCopyFrom = "Select * from PE_CopyFrom where ChannelID=" & ChannelID & " and SourceName='" & SourceName & "'"
    rsCopyFrom.Open sqlCopyFrom, Conn, 1, 3
    If Not (rsCopyFrom.BOF And rsCopyFrom.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ���Դ��</li>"
        rsCopyFrom.Close
        Set rsCopyFrom = Nothing
        Exit Sub
    End If
    rsCopyFrom.addnew
    rsCopyFrom("ChannelID") = ChannelID
    rsCopyFrom("SourceName") = SourceName
    If Photo <> "" Then rsCopyFrom("Photo") = Photo
    If Intro <> "" Then rsCopyFrom("Intro") = Intro
    If Address <> "" Then rsCopyFrom("Address") = Address
    If Tel <> "" Then rsCopyFrom("Tel") = Tel
    If Fax <> "" Then rsCopyFrom("Fax") = Fax
    If Mail <> "" Then rsCopyFrom("Mail") = Mail
    If ZipCode <> "" Then rsCopyFrom("ZipCode") = ZipCode
    If Homepage <> "" Then rsCopyFrom("HomePage") = Homepage
    If Email <> "" Then rsCopyFrom("Email") = Email
    If QQ <> "" Then rsCopyFrom("QQ") = PE_CLng(QQ)
    If ContacterName <> "" Then rsCopyFrom("ContacterName") = ContacterName
    rsCopyFrom("SourceType") = PE_CLng(SourceType)
    rsCopyFrom("LastUseTime") = Now()
    rsCopyFrom("Passed") = True
    rsCopyFrom.Update
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom"
End Sub

Sub SaveModifyCopyFrom()
    Dim SourceName, CopyFromID, Address, Tel, Fax, ContacterName, ZipCode, Homepage, Mail, Email, QQ, Intro, Photo, SourceType
    Dim rsCopyFrom, sqlCopyFrom
    SourceName = Trim(Request("SourceName"))
    If SourceName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Դ���Ʋ���Ϊ�գ�</li>"
    End If
    CopyFromID = Trim(Request("ID"))
    If CopyFromID <> "" Then
        If InStr(CopyFromID, ",") > 0 Then
            CopyFromID = ReplaceBadChar(CopyFromID)
        Else
            CopyFromID = PE_CLng(CopyFromID)
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ���ԴID��</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    Photo = Trim(Request("Photo"))
    Address = Trim(Request("Address"))
    Tel = Trim(Request("Tel"))
    Fax = Trim(Request("Fax"))
    ContacterName = Trim(Request("ContacterName"))
    Mail = Trim(Request("Mail"))
    Photo = Trim(Request("Photo"))
    ZipCode = Trim(Request("ZipCode"))
    Homepage = Trim(Request("HomePage"))
    Email = Trim(Request("Email"))
    QQ = Trim(Request("QQ"))
    Intro = Trim(Request("Intro"))
    SourceType = Trim(Request("SourceType"))
    
    Set rsCopyFrom = Server.CreateObject("Adodb.RecordSet")
    sqlCopyFrom = "Select * from PE_CopyFrom where ID=" & CopyFromID
    rsCopyFrom.Open sqlCopyFrom, Conn, 1, 3
    If Not (rsCopyFrom.BOF And rsCopyFrom.EOF) Then
        rsCopyFrom("ChannelID") = ChannelID
        If SourceName <> "" Then rsCopyFrom("SourceName") = SourceName
        If Photo <> "" Then rsCopyFrom("Photo") = Photo
        rsCopyFrom("Intro") = Intro
        rsCopyFrom("Address") = Address
        rsCopyFrom("Tel") = Tel
        rsCopyFrom("Fax") = Fax
        rsCopyFrom("Mail") = Mail
        rsCopyFrom("ZipCode") = ZipCode
        rsCopyFrom("HomePage") = Homepage
        rsCopyFrom("Email") = Email
        rsCopyFrom("QQ") = PE_CLng(QQ)
        rsCopyFrom("ContacterName") = ContacterName
        rsCopyFrom("SourceType") = PE_CLng(SourceType)
        rsCopyFrom.Update
    End If
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom"
End Sub

Sub DelCopyFrom()
    Dim CopyFromID
    CopyFromID = Trim(Request("ID"))
    If IsValidID(CopyFromID) = False Then
        CopyFromID = ""
    End If
    If CopyFromID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ������ԴID</li>"
        Exit Sub
    End If
    If InStr(CopyFromID, ",") > 0 Then
        Conn.Execute ("delete from PE_CopyFrom where ID in (" & CopyFromID & ")")
    Else
        Conn.Execute ("delete from PE_CopyFrom where ID=" & CopyFromID & "")
    End If
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom"
End Sub

Function GetCopyFromType(TypeID)
    Select Case TypeID
    Case 1
        GetCopyFromType = XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "����վ��")
    Case 2
        GetCopyFromType = XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "����վ��")
    Case 3
        GetCopyFromType = XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "����վ��")
    Case Else
        GetCopyFromType = XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "������Դ")
    End Select
End Function

'**************
'վ�����Ӽ��ַ��滻������
'**************
Sub KeyLink(iType)
    Dim rsKeylink, sqlKeylink, itext, LinkName, ReplaceType, iCount
    ReplaceType = Trim(Request("ReplaceType"))
    If ReplaceType = "" Then PE_CLng (ReplaceType)
    If iType = 0 Then
        itext = "վ������"
        LinkName = "KeyLink"
        Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr><td height='22'>�����ڵ�λ�ã���վ����&nbsp;&gt;&gt;&nbsp;" & itext & "����</td></tr>"
        Response.Write "</table>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
        Response.Write "  <form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�е�" & itext & "��');"">"
        Response.Write "     <td>"
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "  <tr align='center' class='title' height='22'>"
        Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
        Response.Write "    <td width='40'><strong>���</strong></td>"
        Response.Write "    <td width='200'><strong>����Ŀ��</strong></td>"
        Response.Write "    <td height='22'><strong>���ӵ�ַ</strong></td>"
        Response.Write "    <td width='40'><strong>���ȼ�</strong></td>"
        Response.Write "    <td width='40'><strong>״̬</strong></td>"
        Response.Write "    <td width='100'><strong>�� ��</strong></td>"
        Response.Write "  </tr>"
    Else
        itext = "�ַ��滻"
        LinkName = "Rtext"
        Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | <a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=0'>"
        If ReplaceType = "0" Then
            Response.Write "<font color=red>ȫ���滻</font>"
        Else
            Response.Write "ȫ���滻"
        End If
        Response.Write "</a> | <a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=1'>"
        If ReplaceType = "1" Then
            Response.Write "<font color=red>�滻����</font>"
        Else
            Response.Write "�滻����"
        End If
        Response.Write "</a> | <a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=2'>"
        If ReplaceType = "2" Then
            Response.Write "<font color=red>�滻����</font>"
        Else
            Response.Write "�滻����"
        End If
        Response.Write "</a> | <a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=3'>"
        If ReplaceType = "3" Then
            Response.Write "<font color=red>�滻����</font>"
        Else
            Response.Write "�滻����"
        End If
        Response.Write "</a> | <a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=4'>"
        If ReplaceType = "4" Then
            Response.Write "<font color=red>�滻����</font>"
        Else
            Response.Write "�滻����"
        End If
        Response.Write "</a> |</td></tr></table><br>"
        Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr><td height='22'>�����ڵ�λ�ã���վ����&nbsp;&gt;&gt;&nbsp;" & itext & "����</td></tr>"
        Response.Write "</table>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
        Response.Write "  <form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�е�" & itext & "��');"">"
        Response.Write "     <td>"
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "  <tr align='center' height='22' class='title'>"
        Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
        Response.Write "    <td width='40'><strong>���</strong></td>"
        Response.Write "    <td width='120'><strong>�滻Ŀ��</strong></td>"
        Response.Write "    <td height='22'><strong>�滻����</strong></td>"
        Response.Write "    <td width='80'><strong>����</strong></td>"
        Response.Write "    <td width='40'><strong>���ȼ�</strong></td>"
        Response.Write "    <td width='40'><strong>״̬</strong></td>"
        Response.Write "    <td width='100'><strong>�� ��</strong></td>"
        Response.Write "  </tr>"
    End If

    Set rsKeylink = Server.CreateObject("Adodb.RecordSet")
    If ReplaceType = "" Then
        sqlKeylink = "select * from PE_KeyLink Where LinkType=" & iType
    Else
        sqlKeylink = "select * from PE_KeyLink Where LinkType=" & iType & " and ReplaceType=" & ReplaceType
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "Source"
            sqlKeylink = sqlKeylink & " and Source like '%" & Keyword & "%' "
        Case "ReplaceText"
            sqlKeylink = sqlKeylink & " and ReplaceText like '%" & Keyword & "%' "
        Case Else
            sqlKeylink = sqlKeylink & " and Source like '%" & Keyword & "%' "
        End Select
    End If
    sqlKeylink = sqlKeylink & " order by ID Desc"
    strFileName = strFileName & "&ReplaceType=" & ReplaceType
    rsKeylink.Open sqlKeylink, Conn, 1, 1
    If rsKeylink.BOF And rsKeylink.EOF Then
        rsKeylink.Close
        Set rsKeylink = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='8' align='center'><br>û���κ�" & itext & "��<br><br></td></tr>"
        Response.Write "</Table>"
        Exit Sub
    End If
    
    totalPut = rsKeylink.RecordCount
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
            rsKeylink.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Do While Not rsKeylink.EOF
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsKeylink("ID") & "'"
        Response.Write " onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsKeylink("ID") & "</td>"
        Response.Write "    <td>" & rsKeylink("Source") & "</td>"
        Response.Write "    <td>" & GetSubStr(rsKeylink("ReplaceText"), 30, True) & "</td>"
        If iType > 0 Then
            Response.Write "<td>"
            Select Case rsKeylink("ReplaceType")
            Case 0
                Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=0'>ȫ���滻</a>"
            Case 1
                Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=1'>�����滻</a>"
            Case 2
                Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=2'>�����滻</a>"
            Case 3
                Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=3'>�����滻</a>"
            Case 4
                Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=Rtext&ReplaceType=4'>�����滻</a>"
            End Select
            Response.Write "</td>"
        End If
        Response.Write "    <td>" & rsKeylink("Priority") & "</td>"
        Response.Write "    <td>" & GetKeyLinkStatus(rsKeylink("isUse"), rsKeylink("ID"), LinkName) & "</td>"
        Response.Write "<td><a href='Admin_SourceManage.asp?TypeSelect=Modify" & LinkName & "&ID=" & rsKeylink("ID") & "'>�޸�</a>&nbsp;&nbsp;"
        Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=Del" & LinkName & "&ID=" & rsKeylink("ID") & "' onClick=""return confirm('ȷ��Ҫɾ�����" & itext & "��');"">ɾ��</a></td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsKeylink.MoveNext
    Loop
    rsKeylink.Close
    Set rsKeylink = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ������" & itext & "</td>"
    Response.Write "    <td><input name='TypeSelect' type='hidden' id='TypeSelect' value='Del" & LinkName & "'>"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�е�" & itext & "'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��" & itext & "", True)
    
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>������</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SourceManage.asp'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='" & TypeSelect & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    
    Response.Write "<select name='Field' size='1'><option value='Source' selected>Ŀ��</option><option value='ReplaceText'>����</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "�ؼ���"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub AddKeyLink(iType)
    Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='javascript:return CheckKeyLink();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    If iType = 0 Then
        Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� վ �� �� ��</strong></div></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='100' align='right' class='tdbg'><strong>����Ŀ�꣺</strong></td><td><input name='Source' type='text' size='80'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong>���ӵ�ַ��</strong></td><td><input name='Target' type='text'size='80' value='http://'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong>���ȼ���</strong></td><td><input name='Priority' type='text'size='5' value='1'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg''><strong>�滻������</strong></td><td><input name='ReplaceType' type='text'size='5' value='0'> <font color='#FF0000'>�滻Ŀ��Ĵ�����Ϊ0ʱ��ȫ���滻</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong>�򿪷�ʽ��</strong></td><td><input name='OpenType' type='radio' value='0' checked>ԭ����&nbsp;<input name='OpenType' type='radio' value='1'>�´���&nbsp;</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong>״̬��</strong></td><td><input name='Use' type='radio' value='1' checked>����&nbsp;<input name='Use' type='radio' value='0'>����&nbsp;</td>"
        Response.Write "    </tr>"
        Response.Write "  <tr>"
        Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
        Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddKeyLink'>"
        Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?TypeSelect=KeyLink';"" style='cursor:hand;'></td>"
    Else
        Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� ��</strong></div></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='100' align='right' class='tdbg'><strong> �滻Ŀ�꣺</strong></td><td><input name='Source' type='text' size='80'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong> �滻���ݣ�</strong></td><td><input name='Target' type='text'size='80'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong> ���ȼ���</strong></td><td><input name='Priority' type='text'size='5' value='1'> <font color='#FF0000'>*</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong> �滻��ʽ��</strong></td><td><input type='radio' name='ReplaceType' value='0' checked>ȫ���滻&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='1'>Ӧ��������&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='2'>Ӧ���ڱ���&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='3'>Ӧ��������&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='4'>Ӧ��������</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td align='right' class='tdbg'><strong> ״̬��</strong></td><td><input name='Use' type='radio' value='1' checked>����&nbsp;<input name='Use' type='radio' value='0'>����&nbsp;</td>"
        Response.Write "    </tr>"
        Response.Write "  <tr>"
        Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
        Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddRtext'>"
        Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?TypeSelect=Rtext';"" style='cursor:hand;'></td>"
    End If
    
    Response.Write "  </tr>"
    Response.Write "</table></form>"
End Sub

Sub ModifyKeyLink(iType)
    Dim KeyID
    Dim rsKey, sqlKey
    KeyID = PE_CLng(Trim(Request("ID")))
    If KeyID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ID</li>"
        Exit Sub
    End If
    sqlKey = "Select ID,Source,ReplaceText,isUse,LinkType,OpenType,ReplaceType,Priority from PE_KeyLink where ID=" & KeyID
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    rsKey.Open sqlKey, Conn, 1, 1
    If rsKey.BOF And rsKey.EOF Then
        FoundErr = True
        If iType = 1 Then
            ErrMsg = ErrMsg & "<li>�����ڴ�վ�����ӣ�</li>"
        Else
            ErrMsg = ErrMsg & "<li>�����ڴ��ַ��滻��</li>"
        End If
    Else
        Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='return CheckKeyLink();'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        If rsKey("LinkType") = 0 Then
            Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� վ �� �� ��</strong></font></div></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td width='100' align='right' class='tdbg'><strong>����Ŀ�꣺</strong></td><td><input name='Source' type='text' value='" & rsKey("Source") & "' size='80'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td align='right' class='tdbg'><strong>���ӵ�ַ��</strong></td><td><input name='Target' type='text' value='" & rsKey("ReplaceText") & "' size='80'> <font color='#FF0000'>*</font></select>"
            Response.Write "    </td></tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong>���ȼ���</strong></td><td><input name='Priority' type='text'size='5' value='" & rsKey("Priority") & "'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong>�滻������</strong></td><td><input name='ReplaceType' type='text'size='5' value='" & rsKey("ReplaceType") & "'> <font color='#FF0000'>�滻Ŀ��Ĵ�����Ϊ0ʱ��ȫ���滻</font></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong>�򿪷�ʽ��</strong></td><td><input name='OpenType' type='radio' value='0'"
            If rsKey("OpenType") = 0 Then Response.Write "checked"
            Response.Write ">ԭ����&nbsp;<input name='OpenType' type='radio' value='1'"
            If rsKey("OpenType") = 1 Then Response.Write "checked"
            Response.Write ">�´���&nbsp;</td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong>״̬��</strong></td><td><input name='Use' type='radio' value='1'"
            If rsKey("isUse") = 1 Then Response.Write "checked"
            Response.Write ">����&nbsp;<input name='Use' type='radio' value='0'"
            If rsKey("isUse") = 0 Then Response.Write "checked"
            Response.Write ">����&nbsp;</td>"
            Response.Write "    </tr>"
            Response.Write "    <tr>"
            Response.Write "      <td colspan='2' align='center' class='tdbg'>"
            Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsKey("ID") & ">"
            Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyKeyLink'>"
            Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?TypeSelect=KeyLink'"" style='cursor:hand;'></td>"
        Else
            Response.Write "      <td height='22' colspan='2'><div align='center'><font size='2'><strong>�� �� �� �� �� ��</strong></font></div></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td width='100' align='right' class='tdbg'><strong>�滻Ŀ�꣺</strong></td><td><input name='Source' type='text' value='" & rsKey("Source") & "' size='80'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td align='right' class='tdbg'><strong>�滻���ݣ�</strong></td><td><input name='Target' type='text' value='" & rsKey("ReplaceText") & "' size='80'> <font color='#FF0000'>*</font></select>"
            Response.Write "    </td></tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong> ���ȼ���</strong></td><td><input name='Priority' type='text'size='5' value='" & rsKey("Priority") & "'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong> �滻��ʽ��</strong></td><td><input type='radio' name='ReplaceType' value='0'"
            If rsKey("ReplaceType") = 0 Then Response.Write " checked"
            Response.Write ">ȫ���滻&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='1'"
            If rsKey("ReplaceType") = 1 Then Response.Write " checked"
            Response.Write ">Ӧ��������&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='2'"
            If rsKey("ReplaceType") = 2 Then Response.Write " checked"
            Response.Write ">Ӧ���ڱ���&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='3'"
            If rsKey("ReplaceType") = 3 Then Response.Write " checked"
            Response.Write ">Ӧ��������&nbsp;&nbsp;<input type='radio' name='ReplaceType' value='4'"
            If rsKey("ReplaceType") = 4 Then Response.Write " checked"
            Response.Write ">Ӧ��������</td></tr>"
            Response.Write "    <tr class='tdbg'> "
            Response.Write "      <td align='right' class='tdbg'><strong>״̬��</strong></td><td><input name='Use' type='radio' value='1'"
            If rsKey("isUse") = 1 Then Response.Write "checked"
            Response.Write ">����&nbsp;<input name='Use' type='radio' value='0'"
            If rsKey("isUse") = 0 Then Response.Write "checked"
            Response.Write ">����&nbsp;</td>"
            Response.Write "    </tr>"
            Response.Write "    <tr>"
            Response.Write "      <td colspan='2' align='center' class='tdbg'>"
            Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsKey("ID") & ">"
            Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyRtext'>"
            Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?TypeSelect=Rtext'"" style='cursor:hand;'></td>"
        End If
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsKey.Close
    Set rsKey = Nothing
End Sub

Sub SaveAddKeyLink(iType)
    Dim Source, RText, Use, ReplaceType, OpenType, Priority
    Dim rsKey, sqlKey
    
    Source = Trim(Request("Source"))
    RText = Trim(Request("Target"))
    Use = Trim(Request("Use"))
    ReplaceType = PE_CLng(Trim(Request("ReplaceType")))
    Priority = PE_CLng(Trim(Request("Priority")))
    If Priority = 0 Then Priority = 1
    OpenType = PE_CLng(Trim(Request("OpenType")))

    If Source = "" Or RText = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݲ���Ϊ�գ�</li>"
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    sqlKey = "Select Source,ReplaceText,isUse,LinkType,OpenType,ReplaceType,Priority from PE_KeyLink where LinkType=" & iType & " and Source='" & Source & "'"
    rsKey.Open sqlKey, Conn, 1, 3
    If Not (rsKey.BOF And rsKey.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ����ӣ�</li>"
        rsKey.Close
        Set rsKey = Nothing
        Exit Sub
    End If
    rsKey.addnew
    rsKey("Source") = Source
    rsKey("ReplaceText") = PE_HTMLEncode(RText)
    rsKey("isUse") = Use
    rsKey("LinkType") = iType
    rsKey("OpenType") = OpenType
    rsKey("ReplaceType") = ReplaceType
    rsKey("Priority") = Priority
    rsKey.Update
    rsKey.Close
    
    '���»���
    Dim arrKeyList
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=" & iType & " order by Priority"
    rsKey.Open sqlKey, Conn, 1, 1
    If Not (rsKey.BOF And rsKey.EOF) Then
        arrKeyList = rsKey.GetString(, , "|||", "@@@", "")
        If iType = 0 Then
            PE_Cache.SetValue "Site_KeyList", arrKeyList
        Else
            PE_Cache.SetValue "Site_ReplaceText", arrKeyList
        End If
    End If
    rsKey.Close
    Set rsKey = Nothing
    
    Call CloseConn
    If iType = 0 Then
        Response.Redirect "Admin_SourceManage.asp?TypeSelect=KeyLink"
    Else
        Response.Redirect "Admin_SourceManage.asp?TypeSelect=Rtext"
    End If
End Sub

Sub SaveModifyKeyLink(iType)
    Dim Source, RText, Use, KeyID, ReplaceType, OpenType, Priority
    Dim rsKey, sqlKey
    Source = Trim(Request("Source"))
    RText = Trim(Request("Target"))
    Use = Trim(Request("Use"))
    ReplaceType = PE_CLng(Trim(Request("ReplaceType")))
    Priority = PE_CLng(Trim(Request("Priority")))
    If Priority = 0 Then Priority = 1
    OpenType = PE_CLng(Trim(Request("OpenType")))

    KeyID = Trim(Request("ID"))
    If Source = "" Or RText = "" Or KeyID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݲ���Ϊ�գ�</li>"
    Else
        KeyID = PE_CLng(KeyID)
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    sqlKey = "Select ID,Source,ReplaceText,isUse,LinkType,OpenType,ReplaceType,Priority from PE_KeyLink where ID=" & KeyID
    rsKey.Open sqlKey, Conn, 1, 3
    If Not (rsKey.BOF And rsKey.EOF) Then
        rsKey("Source") = Source
        rsKey("ReplaceText") = PE_HTMLEncode(RText)
        rsKey("isUse") = Use
        rsKey("OpenType") = OpenType
        rsKey("ReplaceType") = ReplaceType
        rsKey("Priority") = Priority
        rsKey.Update
    End If
    rsKey.Close
    
    '���»���
    Dim arrKeyList
    Set rsKey = Server.CreateObject("Adodb.RecordSet")
    sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=" & iType & " order by Priority"
    rsKey.Open sqlKey, Conn, 1, 1
    If Not (rsKey.BOF And rsKey.EOF) Then
        arrKeyList = rsKey.GetString(, , "|||", "@@@", "")
        If iType = 0 Then
            PE_Cache.SetValue "Site_KeyList", arrKeyList
        Else
            PE_Cache.SetValue "Site_ReplaceText", arrKeyList
        End If
    End If
    rsKey.Close
    Set rsKey = Nothing
    
    Call CloseConn
    If iType = 0 Then
        Response.Redirect "Admin_SourceManage.asp?TypeSelect=KeyLink"
    Else
        Response.Redirect "Admin_SourceManage.asp?TypeSelect=Rtext"
    End If
End Sub

Sub DelKeyLink(iType)
    Dim KeyID
    Dim rsKey, sqlKey
    
    KeyID = Trim(Request("ID"))
    If IsValidID(KeyID) = False Then
        KeyID = ""
    End If
    If KeyID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ��������ID</li>"
        Exit Sub
    End If
    If InStr(KeyID, ",") > 0 Then
        Conn.Execute ("delete from PE_KeyLink where ID in (" & KeyID & ")")
    Else
        Conn.Execute ("delete from PE_KeyLink where ID=" & KeyID & "")
    End If

    '���»���
    PE_Cache.DelAllCache

    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?TypeSelect=" & iType
End Sub

Function GetKeyLinkStatus(StatID, LinkID, iType)
    If StatID = 0 Then
        GetKeyLinkStatus = "<a href='Admin_SourceManage.asp?TypeSelect=run" & iType & "&ID=" & LinkID & "'>����</a>"
    ElseIf StatID = 1 Then
        GetKeyLinkStatus = "<a href='Admin_SourceManage.asp?TypeSelect=dis" & iType & "&ID=" & LinkID & "'>����</a>"
    Else
        GetKeyLinkStatus = "<a href='Admin_SourceManage.asp?TypeSelect=run" & iType & "&ID=" & LinkID & "'>δ֪</a>"
    End If
End Function

Sub SetKeyLink(iType, Stat)
    Dim KeylinkID
    KeylinkID = Trim(Request("ID"))
    If KeylinkID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���õ�����ID</li>"
        Exit Sub
    Else
        KeylinkID = PE_CLng(KeylinkID)
    End If
    Conn.Execute ("update PE_KeyLink set isUse=" & Stat & " where ID=" & KeylinkID & "")

    '���»���
    PE_Cache.DelAllCache

    Call KeyLink(iType)
End Sub

'**************
'���̴�����
'**************

Sub Producer()
    Dim rsProducer, sqlProducer
    Dim iCount
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer&ItemType=1'>"
    If ItemType = 1 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowProducer/ProducerType1", "��½����") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowProducer/ProducerType1", "��½����")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer&ItemType=2'>"
    If ItemType = 2 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowProducer/ProducerType2", "��̨����") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowProducer/ProducerType2", "��̨����")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer&ItemType=3'>"
    If ItemType = 3 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowProducer/ProducerType3", "�պ�����") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowProducer/ProducerType3", "�պ�����")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer&ItemType=4'>"
    If ItemType = 4 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowProducer/ProducerType4", "ŷ������") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowProducer/ProducerType4", "ŷ������")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer&ItemType=0'>"
    If ItemType = 0 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowProducer/ProducerType5", "��������") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowProducer/ProducerType5", "��������")
    End If
    Response.Write "</a> |</td></tr></table><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�еĳ�����');"">"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='40' height='22'><strong>���</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>��������</strong></td>"
    Response.Write "    <td width='80' height='22'><strong>������д</strong></td>"
    Response.Write "    <td height='22'><strong>���</strong></td>"
    Response.Write "    <td width='80' height='22'><strong>���̷���</strong></td>"
    Response.Write "    <td width='60' height='22'><strong>״̬</strong></td>"
    Response.Write "    <td width='220' height='22'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"
    
    Set rsProducer = Server.CreateObject("Adodb.RecordSet")
    sqlProducer = "select * from PE_Producer Where ChannelID=" & ChannelID
    If Keyword <> "" Then
        Select Case strField
        Case "name"
            sqlProducer = sqlProducer & " and ProducerName like '%" & Keyword & "%' "
        Case "suoxie"
            sqlProducer = sqlProducer & " and ProducerShortName like '%" & Keyword & "%' "
        Case "address"
            sqlProducer = sqlProducer & " and Address like '%" & Keyword & "%' "
        Case "Postcode"
            sqlProducer = sqlProducer & " and Postcode like '%" & Keyword & "%' "
        Case "Phone"
            sqlProducer = sqlProducer & " and Address like '%" & Keyword & "%' "
        Case "intro"
            sqlProducer = sqlProducer & " and ProducerIntro like '%" & Keyword & "%' "
        Case Else
            sqlProducer = sqlProducer & " and ProducerName like '%" & Keyword & "%' "
        End Select
    End If
    If ItemType < 999 Then
        sqlProducer = sqlProducer & " and ProducerType =" & ItemType
    End If
    sqlProducer = sqlProducer & " order by ProducerID Desc"
    rsProducer.Open sqlProducer, Conn, 1, 1
    If rsProducer.BOF And rsProducer.EOF Then
        rsProducer.Close
        Set rsProducer = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='8' align='center'><br>û���κγ��̣�<br><br></td></tr>"
        Response.Write "</Table>"
        Exit Sub
    End If
    
    totalPut = rsProducer.RecordCount
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
            rsProducer.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Do While Not rsProducer.EOF
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsProducer("ProducerID") & "' onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsProducer("ProducerID") & "</td>"
        Response.Write "    <td><a href='Admin_SourceManage.asp?TypeSelect=Trademark&ChannelID=" & ChannelID & "&ProducerID=" & rsProducer("ProducerID") & "' title='�鿴�������������̱��б�'>" & rsProducer("ProducerName") & "</a></td>"
        Response.Write "    <td>" & rsProducer("ProducerShortName") & "</td>"
        Response.Write "    <td>" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), 30, False)
        If Len(rsProducer("ProducerIntro")) > 32 Then Response.Write "��"
        Response.Write "    </td>"
        Response.Write "    <td>" & GetProducerType(rsProducer("ProducerType")) & "</td><td>"
        If rsProducer("Passed") = True Then
            Response.Write "<font color=""green"">��</font>"
        Else
            Response.Write "<font color=""red"">��</font>"
        End If
        If rsProducer("onTop") = True Then
            Response.Write "&nbsp;<font color=""blue"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        If rsProducer("isElite") = True Then
            Response.Write "&nbsp;<font color=""green"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td><td>"
        Response.Write "      <a href='Admin_SourceManage.asp?TypeSelect=AddTrademark&ChannelID=" & ChannelID & "&ProducerID=" & rsProducer("ProducerID") & "'>�����̱�</a>&nbsp;"
        Response.Write "      <a href='Admin_SourceManage.asp?TypeSelect=ModifyProducer&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>�޸�</a>"
        If rsProducer("Passed") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=ProducerDis&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>����</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=ProducerEn&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>����</a>"
        End If
        If rsProducer("onTop") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=ProducerDTop&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=ProducerTop&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>�̶�</a>"
        End If
        If rsProducer("isElite") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=ProducerDElite&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=ProducerElite&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "'>�Ƽ�</a>"
        End If
        Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=DelProducer&ChannelID=" & ChannelID & "&ID=" & rsProducer("ProducerID") & "' onClick=""return confirm('ȷ��Ҫɾ������" & rsProducer("ProducerName") & "��');"">ɾ��</a>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsProducer.MoveNext
    Loop
    rsProducer.Close
    Set rsProducer = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ�����г���</td>"
    Response.Write "    <td><input name='TypeSelect' type='hidden' id='TypeSelect' value='DelProducer'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�еĳ���'></td>"
    Response.Write "  </tr>"
    Response.Write "</form></table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������", True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>����������</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SourceManage.asp'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='" & TypeSelect & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>��������</option><option value='suoxie'>������д</option><option value='address'>���̵�ַ</option><option value='Postcode'>�����ʱ�</option><option value='Phone'>���̵绰</option><option value='intro'>���̼��</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "�ؼ���"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub AddProducer()
    Call PopCalendarInit
    Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
    Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� �� Ϣ</strong></div></td>"
    Response.Write "    </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='300'class='tdbg'>&nbsp;<strong> �������ƣ�</strong><input name='ProducerName' type='text'> <font color='#FF0000'>*</font></td>"
    Response.Write "    <td rowspan='9' align='center' valign='top' class='tdbg'>"
    Response.Write "        <table width='180' height='200' border='1'>"
    Response.Write "            <tr><td width='100%' align='center'><img id='showphoto' src='" & InstallDir & "Shop/ProducerPic/default.gif' width='150' height='172'></td></tr>"
    Response.Write "        </table>"
    Response.Write "        <input name='Photo' type='text' size='25'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=ProducerPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
    Response.Write "     </td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������д��</strong><input name='ShortName' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������ڣ�</strong><input name='BirthDay' type='text' value='" & FormatDateTime(Date, 2) & "' maxlength='20'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BirthDay, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��˾��ַ��</strong><input name='Address' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��ϵ�绰��</strong><input name='Tel' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������룺</strong><input name='Fax' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������룺</strong><input name='Postcode' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������ҳ��</strong><input name='HomePage' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �����ʼ���</strong><input name='Email' type='text'></td></tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>���̷��ࣺ</strong><input name='ProducerType' type='radio' value='1' checked>" & XmlText("ShowSource", "ShowProducer/ProducerType1", "��½����") & "&nbsp;<input name='ProducerType' type='radio' value='2'>" & XmlText("ShowSource", "ShowProducer/ProducerType2", "��̨����") & "&nbsp;<input name='ProducerType' type='radio' value='3'>" & XmlText("ShowSource", "ShowProducer/ProducerType3", "�պ�����") & "&nbsp;<input name='ProducerType' type='radio' value='4'>" & XmlText("ShowSource", "ShowProducer/ProducerType4", "ŷ������") & "&nbsp;<input name='ProducerType' type='radio' value='0'>" & XmlText("ShowSource", "ShowProducer/ProducerType5", "��������") & "&nbsp;</td></tr>"
    Response.Write "  <tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>���̼��</strong>��<br>"
    Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'></textarea>"
    Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "  <td height='40' colspan='2' align='center' class='tdbg'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddProducer'>"
    Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer';"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
End Sub

Sub ModifyProducer()
    Dim ProducerID
    Dim rsProducer, sqlProducer
    ProducerID = PE_CLng(Trim(Request("ID")))
    If ProducerID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĳ���ID</li>"
        Exit Sub
    End If
    sqlProducer = "Select * from PE_Producer where ProducerID=" & ProducerID
    Set rsProducer = Server.CreateObject("Adodb.RecordSet")
    rsProducer.Open sqlProducer, Conn, 1, 1
    If rsProducer.BOF And rsProducer.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ˳��̣�</li>"
    Else
        Call PopCalendarInit
        Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='return CheckInput();'>"
        Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� �� �� �� �� Ϣ</strong></font></div></td>"
        Response.Write "    </tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong> �������ƣ�</strong><input name='ProducerName' type='text' value='" & rsProducer("ProducerName") & "'> <font color='#FF0000'>*</font></td>"
        Response.Write "    <td rowspan='9' align='center' valign='top' class='tdbg'>"
        Response.Write "        <table width='180' height='200' border='1'>"
        Response.Write "            <tr><td width='100%' align='center'>"
        If IsNull(rsProducer("ProducerPhoto")) Then
            Response.Write "<img id='showphoto' src='" & InstallDir & "Shop/ProducerPic/default.gif' width='150' height='172'>"
        Else
            Response.Write "<img id='showphoto' src='" & rsProducer("ProducerPhoto") & "' width='150' height='172'>"
        End If
        Response.Write "        </td></tr></table>"
        Response.Write "        <input name='Photo' type='text' size='25' value='" & rsProducer("ProducerPhoto") & "'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=ProducerPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
        Response.Write "     </td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������д��</strong><input name='ShortName' type='text' value='" & rsProducer("ProducerShortName") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������ڣ�</strong><input name='BirthDay' type='text'  value='" & rsProducer("BirthDAy") & "' maxlength='20'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BirthDay, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ��˾��ַ��</strong><input name='Address' type='text'  value='" & rsProducer("Address") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �绰���룺</strong><input name='Tel' type='text' value='" & rsProducer("Phone") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������룺</strong><input name='Fax' type='text' value='" & rsProducer("Fax") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������룺</strong><input name='Postcode' type='text' value='" & rsProducer("PostCode") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> ������ҳ��</strong><input name='HomePage' type='text' value='" & rsProducer("Homepage") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �����ʼ���</strong><input name='Email' type='text' value='" & rsProducer("Email") & "'></td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>���̷��ࣺ</strong>"
        Response.Write "<input name='ProducerType' type='radio' value='1'"
    If rsProducer("ProducerType") = 1 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType1", "��½����") & "&nbsp;<input name='ProducerType' type='radio' value='2'"
    If rsProducer("ProducerType") = 2 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType2", "��̨����") & "&nbsp;<input name='ProducerType' type='radio' value='3'"
    If rsProducer("ProducerType") = 3 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType3", "�պ�����") & "&nbsp;<input name='ProducerType' type='radio' value='4'"
    If rsProducer("ProducerType") = 4 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType4", "ŷ������") & "&nbsp;<input name='ProducerType' type='radio' value='0'"
    If rsProducer("ProducerType") = 0 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType5", "��������") & "&nbsp;</td></tr>"
        Response.Write "  <tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>���̼��</strong>��<br>"
        Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>"
        If Trim(rsProducer("ProducerIntro") & "") <> "" Then Response.Write Server.HTMLEncode(rsProducer("ProducerIntro"))
        Response.Write "      </textarea>"
        Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td colspan='2' align='center' class='tdbg'>"
        Response.Write "      <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyProducer'>"
        Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsProducer("ProducerID") & ">"
        Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsProducer.Close
    Set rsProducer = Nothing
End Sub

Sub SaveAddProducer()
    Dim ProducerName, ShortName, Birthday, Address, Tel, Fax, PostCode, Homepage, Email, Intro, Photo, ProducerType
    Dim rsProducer, sqlProducer
    ProducerName = Trim(Request("ProducerName"))
    Birthday = Trim(Request("BirthDay"))

    If ProducerName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������Ʋ���Ϊ�գ�</li>"
    Else
        ProducerName = ReplaceBadChar(ProducerName)
    End If

    If IsDate(Birthday) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ڸ�ʽ����ȷ��</li>"
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    ShortName = Trim(Request("ShortName"))
    Photo = Trim(Request("Photo"))
    Address = Trim(Request("Address"))
    Tel = Trim(Request("Tel"))
    Fax = Trim(Request("Fax"))
    PostCode = Trim(Request("PostCode"))
    Homepage = Trim(Request("HomePage"))
    Email = Trim(Request("Email"))
    Intro = Trim(Request("Intro"))
    ProducerType = Trim(Request("ProducerType"))
    
    Set rsProducer = Server.CreateObject("Adodb.RecordSet")
    sqlProducer = "Select * from PE_Producer where ChannelID=" & ChannelID & " and ProducerName='" & ProducerName & "'"
    rsProducer.Open sqlProducer, Conn, 1, 3
    If Not (rsProducer.BOF And rsProducer.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ˳��̣�</li>"
        rsProducer.Close
        Set rsProducer = Nothing
        Exit Sub
    End If
    rsProducer.addnew
    rsProducer("ChannelID") = ChannelID
    rsProducer("ProducerName") = ProducerName
    If ShortName <> "" Then rsProducer("ProducerShortName") = ShortName
    If Photo <> "" Then rsProducer("ProducerPhoto") = Photo
    If Birthday <> "" Then rsProducer("BirthDay") = Birthday
    If Address <> "" Then rsProducer("Address") = Address
    If Tel <> "" Then rsProducer("Phone") = Tel
    If Fax <> "" Then rsProducer("Fax") = Fax
    If PostCode <> "" Then rsProducer("Postcode") = PostCode
    If Homepage <> "" Then rsProducer("HomePage") = Homepage
    If Email <> "" Then rsProducer("Email") = Email
    If Intro <> "" Then rsProducer("ProducerIntro") = Intro
    rsProducer("ProducerType") = PE_CLng(ProducerType)
    rsProducer("LastUseTime") = Now()
    rsProducer("Passed") = True
    rsProducer.Update
    rsProducer.Close
    Set rsProducer = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer"
End Sub

Sub SaveModifyProducer()
    Dim ProducerID, ProducerName, ShortName, Birthday, Address, Tel, Fax, PostCode, Homepage, Email, Intro, Photo, ProducerType
    Dim rsProducer, sqlProducer
    ProducerName = Trim(Request("ProducerName"))
    If ProducerName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĳ������ƣ�</li>"
    End If
    ProducerID = Trim(Request("ID"))
    If ProducerID <> "" Then
        If InStr(ProducerID, ",") > 0 Then
            ProducerID = ReplaceBadChar(ProducerID)
        Else
            ProducerID = PE_CLng(ProducerID)
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĳ���ID��</li>"
    End If

    Birthday = Trim(Request("BirthDay"))
    If IsDate(Birthday) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ڸ�ʽ����ȷ��</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    ShortName = Trim(Request("ShortName"))
    Photo = Trim(Request("Photo"))
    Address = Trim(Request("Address"))
    Tel = Trim(Request("Tel"))
    Fax = Trim(Request("Fax"))
    PostCode = Trim(Request("PostCode"))
    Homepage = Trim(Request("HomePage"))
    Email = Trim(Request("Email"))
    Intro = Trim(Request("Intro"))
    ProducerType = Trim(Request("ProducerType"))
    
    Set rsProducer = Server.CreateObject("Adodb.RecordSet")
    sqlProducer = "Select * from PE_Producer where ProducerID=" & ProducerID
    rsProducer.Open sqlProducer, Conn, 1, 3
    If Not (rsProducer.BOF And rsProducer.EOF) Then
        rsProducer("ChannelID") = ChannelID
        rsProducer("ProducerName") = ProducerName
        If ShortName <> "" Then rsProducer("ProducerShortName") = ShortName
        If Photo <> "" Then rsProducer("ProducerPhoto") = Photo
        If Birthday <> "" Then rsProducer("BirthDay") = Birthday
        If Address <> "" Then rsProducer("Address") = Address
        If Tel <> "" Then rsProducer("Phone") = Tel
        If Fax <> "" Then rsProducer("Fax") = Fax
        If PostCode <> "" Then rsProducer("Postcode") = PostCode
        If Homepage <> "" Then rsProducer("HomePage") = Homepage
        If Email <> "" Then rsProducer("Email") = Email
        If Intro <> "" Then rsProducer("ProducerIntro") = Intro
        rsProducer("ProducerType") = PE_CLng(ProducerType)
        rsProducer("LastUseTime") = Now()
        rsProducer.Update
    End If
    rsProducer.Close
    Set rsProducer = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer"
End Sub

Sub DelProducer()
    Dim ProducerID
    ProducerID = Trim(Request("ID"))
    If IsValidID(ProducerID) = False Then
        ProducerID = ""
    End If
    If ProducerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ���ĳ���ID</li>"
        Exit Sub
    End If
    If InStr(ProducerID, ",") > 0 Then
        Conn.Execute ("delete from PE_Producer where ProducerID in (" & ProducerID & ")")
    Else
        Conn.Execute ("delete from PE_Producer where ProducerID=" & ProducerID & "")
    End If
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer"
End Sub

Function GetProducerType(TypeID)
    Select Case TypeID
    Case 1
        GetProducerType = XmlText("ShowSource", "ShowProducer/ProducerType1", "��½����")
    Case 2
        GetProducerType = XmlText("ShowSource", "ShowProducer/ProducerType2", "��̨����")
    Case 3
        GetProducerType = XmlText("ShowSource", "ShowProducer/ProducerType3", "�պ�����")
    Case 4
        GetProducerType = XmlText("ShowSource", "ShowProducer/ProducerType4", "ŷ������")
    Case Else
        GetProducerType = XmlText("ShowSource", "ShowProducer/ProducerType5", "��������")
    End Select
End Function


'**************
'Ʒ�ƴ�����
'**************

Sub Trademark()
    Dim rsTrademark, sqlTrademark, TrademarkID
    Dim iCount
    TrademarkID = PE_CLng(Trim(Request("TrademarkID")))
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title'><td height='22'> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark&ItemType=1'>"
    If ItemType = 1 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowTrademark/TrademarkType1", "��½Ʒ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowTrademark/TrademarkType1", "��½Ʒ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark&ItemType=2'>"
    If ItemType = 2 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowTrademark/TrademarkType2", "��̨Ʒ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowTrademark/TrademarkType2", "��̨Ʒ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark&ItemType=3'>"
    If ItemType = 3 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowTrademark/TrademarkType3", "�պ�Ʒ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowTrademark/TrademarkType3", "�պ�Ʒ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark&ItemType=4'>"
    If ItemType = 4 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowTrademark/TrademarkType4", "ŷ��Ʒ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowTrademark/TrademarkType4", "ŷ��Ʒ��")
    End If
    Response.Write "</a> | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark&ItemType=0'>"
    If ItemType = 0 Then
        Response.Write "<font color=red>" & XmlText("ShowSource", "ShowTrademark/TrademarkType5", "����Ʒ��") & "</font>"
    Else
        Response.Write XmlText("ShowSource", "ShowTrademark/TrademarkType5", "����Ʒ��")
    End If
    Response.Write "</a> |</td></tr></table><br>"
    Response.Write "  <form name='myform' method='Post' action='Admin_SourceManage.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ�е�Ʒ����');"">"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='40' height='22'><strong>���</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>Ʒ������</strong></td>"
    Response.Write "    <td height='22'><strong>���</strong></td>"
    Response.Write "    <td width='80' height='22'><strong>Ʒ�Ʒ���</strong></td>"
    Response.Write "    <td width='60' height='22'><strong>״̬</strong></td>"
    Response.Write "    <td width='150' height='22'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"
    
    Set rsTrademark = Server.CreateObject("Adodb.RecordSet")
    sqlTrademark = "select * from PE_Trademark Where ChannelID=" & ChannelID
    If TrademarkID <> 0 Then sqlTrademark = sqlTrademark & " and TrademarkID=" & TrademarkID
    If Keyword <> "" Then
        Select Case strField
        Case "name"
            sqlTrademark = sqlTrademark & " and TrademarkName like '%" & Keyword & "%' "
        Case "intro"
            sqlTrademark = sqlTrademark & " and TrademarkIntro like '%" & Keyword & "%' "
        Case Else
            sqlTrademark = sqlTrademark & " and TrademarkName like '%" & Keyword & "%' "
        End Select
    End If
    If ItemType < 999 Then
        sqlTrademark = sqlTrademark & " and TrademarkType =" & ItemType
    End If
    
    sqlTrademark = sqlTrademark & " order by IsElite,TrademarkID Desc"
    rsTrademark.Open sqlTrademark, Conn, 1, 1
    If rsTrademark.BOF And rsTrademark.EOF Then
        rsTrademark.Close
        Set rsTrademark = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='7' align='center'><br>û���κ�Ʒ�ƣ�<br><br></td></tr>"
        Response.Write "</Table>"
        Exit Sub
    End If
    
    totalPut = rsTrademark.RecordCount
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
            rsTrademark.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Do While Not rsTrademark.EOF
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ID' type='checkbox' id='ID' value='" & rsTrademark("TrademarkID") & "' onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsTrademark("TrademarkID") & "</td>"
        Response.Write "    <td>" & rsTrademark("TrademarkName") & "</td>"
        Response.Write "    <td> " & GetSubStr(nohtml(PE_HtmlDecode(rsTrademark("TrademarkIntro"))), 30, False)
        If Len(rsTrademark("TrademarkIntro")) > 32 Then Response.Write "��"
        Response.Write "    </td>"
        Response.Write "    <td>" & GetTrademarkType(rsTrademark("TrademarkType")) & "</td><td>"
        If rsTrademark("Passed") = True Then
            Response.Write "<font color=""green"">��</font>"
        Else
            Response.Write "<font color=""red"">��</font>"
        End If
        If rsTrademark("onTop") = True Then
            Response.Write "&nbsp;<font color=""blue"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        If rsTrademark("isElite") = True Then
            Response.Write "&nbsp;<font color=""green"">��</font>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td><td>"
        Response.Write "<a href='Admin_SourceManage.asp?TypeSelect=ModifyTrademark&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>�޸�</a>"
        If rsTrademark("Passed") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=TrademarkDis&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>����</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=TrademarkEn&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>����</a>"
        End If
        If rsTrademark("onTop") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=TrademarkDTop&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=TrademarkTop&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>�̶�</a>"
        End If
        If rsTrademark("isElite") = True Then
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=TrademarkDElite&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>���</a>"
        Else
            Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=TrademarkElite&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "'>�Ƽ�</a>"
        End If
        Response.Write "&nbsp;<a href='Admin_SourceManage.asp?TypeSelect=DelTrademark&ChannelID=" & ChannelID & "&ID=" & rsTrademark("TrademarkID") & "' onClick=""return confirm('ȷ��Ҫɾ��Ʒ��" & rsTrademark("TrademarkName") & "��');"">ɾ��</a>"
        Response.Write "</td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsTrademark.MoveNext
    Loop
    rsTrademark.Close
    Set rsTrademark = Nothing
    
    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ������Ʒ��</td>"
    Response.Write "    <td><input name='TypeSelect' type='hidden' id='TypeSelect' value='DelTrademark'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�е�Ʒ��'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��Ʒ��", True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='80' align='right'><strong>Ʒ��������</strong></td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='Admin_SourceManage.asp'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='" & TypeSelect & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>Ʒ������</option><option value='intro'>Ʒ�Ƽ��</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If Keyword <> "" Then
        Response.Write Keyword
    Else
        Response.Write "�ؼ���"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub AddTrademark()
    Dim ProducerID
    ProducerID = Trim(Request("ProducerID"))
    Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='javascript:return CheckInput();'>"
    Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� Ʒ �� �� Ϣ</strong></font></div></td>"
    Response.Write "    </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='300' height='22' class='tdbg'>&nbsp;<strong> Ʒ�����ƣ�</strong><input name='TrademarkName' type='text'> <font color='#FF0000'>*</font></td>"
    Response.Write "    <td rowspan='4' align='center' valign='top' class='tdbg'>"
    Response.Write "        <table width='180' height='200' border='1'>"
    Response.Write "            <tr><td width='100%' align='center'><img id='showphoto' src='" & InstallDir & "Shop/TrademarkPic/default.gif' width='150' height='172'></td></tr>"
    Response.Write "        </table>"
    Response.Write "        <input name='Photo' type='text' size='25'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=TrademarkPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
    Response.Write "     </td></tr>"
    Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������̣�</strong>"
    If ProducerID = "" Then
        Response.Write "<select name='ProducerID'>" & GetProducerList(ChannelID, 0) & "</select>"
    Else
        Response.Write GetProducerName(ProducerID)
        Response.Write "<input name='ProducerID' type='hidden' id='ProducerID' value='" & ProducerID & "'>"
    End If
    Response.Write "  </td></tr>"
    Response.Write "  <tr class='tdbg'><td height='22'>&nbsp;<strong> �Ƿ��Ƽ���</strong><input type=checkbox name='Elite' value='Yes'></td></tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td><table width='100%'><tr><td rowspan='5' width='80'>&nbsp;<strong>Ʒ�Ʒ��ࣺ</strong></td><td>"
    Response.Write "<input name='TrademarkType' type='radio' value='1' checked>" & XmlText("ShowSource", "ShowTrademark/TrademarkType1", "��½Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='2'>" & XmlText("ShowSource", "ShowTrademark/TrademarkType2", "��̨Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='3'>" & XmlText("ShowSource", "ShowTrademark/TrademarkType3", "�պ�Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='4'>" & XmlText("ShowSource", "ShowTrademark/TrademarkType4", "ŷ��Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='0'>" & XmlText("ShowSource", "ShowTrademark/TrademarkType5", "����Ʒ��") & "</td></tr></table></td></tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td colspan='2'>&nbsp;<strong>Ʒ�Ƽ��</strong>��<br>"
    Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'></textarea>"
    Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr><td height='40' colspan='2' align='center' class='tdbg'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveAddTrademark'>"
    Response.Write "    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark';"" style='cursor:hand;'></td>"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
End Sub

Sub ModifyTrademark()
    Dim TrademarkID
    Dim rsTrademark, sqlTrademark
    TrademarkID = PE_CLng(Trim(Request("ID")))
    If TrademarkID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�Ʒ��ID</li>"
        Exit Sub
    End If
    sqlTrademark = "Select * from PE_Trademark where TrademarkID=" & TrademarkID
    Set rsTrademark = Server.CreateObject("Adodb.RecordSet")
    rsTrademark.Open sqlTrademark, Conn, 1, 1
    If rsTrademark.BOF And rsTrademark.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ�Ʒ�ƣ�</li>"
    Else
        Response.Write "<form method='post' action='Admin_SourceManage.asp' name='myform' onsubmit='return CheckInput();'>"
        Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� Ʒ �� �� Ϣ</strong></font></div></td>"
        Response.Write "    </tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td width='300' class='tdbg'>&nbsp;<strong> Ʒ�����ƣ�</strong><input name='TrademarkName' type='text' value='" & rsTrademark("TrademarkName") & "'> <font color='#FF0000'>*</font></td>"
        Response.Write "    <td rowspan='4' align='center' valign='top' class='tdbg'>"
        Response.Write "        <table width='180' height='200' border='1'>"
        Response.Write "            <tr><td width='100%' align='center'>"
        If IsNull(rsTrademark("TrademarkPhoto")) Then
            Response.Write "<img id='showphoto' src='" & InstallDir & "Shop/TrademarkPic/default.gif' width='150' height='172'>"
        Else
            Response.Write "<img id='showphoto' src='" & rsTrademark("TrademarkPhoto") & "' width='150' height='172'>"
        End If
        Response.Write "        </td></tr></table>"
        Response.Write "        <input name='Photo' type='text' size='25' value='" & rsTrademark("TrademarkPhoto") & "'><strong>���� Ƭ �� ַ</strong><br><iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=TrademarkPic' frameborder=0 scrolling=no width='285' height='25'></iframe>"
        Response.Write "     </td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �������̣�</strong><select name='ProducerID'>" & GetProducerList(ChannelID, rsTrademark("ProducerID")) & "</select></td></tr>"
        Response.Write "  <tr class='tdbg'><td>&nbsp;<strong> �Ƿ��Ƽ���</strong><input type=checkbox name='Elite' value='Yes'"
        If rsTrademark("isElite") = True Then Response.Write " checked"
        Response.Write "></td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td><table width='100%'><tr><td rowspan='5' width='80'>&nbsp;<strong>Ʒ�Ʒ��ࣺ</strong></td><td>"
        Response.Write "<input name='TrademarkType' type='radio' value='1'"
    If rsTrademark("TrademarkType") = 1 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType1", "��½Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='2'"
    If rsTrademark("TrademarkType") = 2 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType2", "��̨Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='3'"
    If rsTrademark("TrademarkType") = 3 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType3", "�պ�Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='4'"
    If rsTrademark("TrademarkType") = 4 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType4", "ŷ��Ʒ��") & "</td></tr><tr><td><input name='TrademarkType' type='radio' value='0'"
    If rsTrademark("TrademarkType") = 0 Then Response.Write " checked"
        Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType5", "����Ʒ��") & "</td></tr></table></td></tr>"
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2'>&nbsp;<strong>Ʒ�Ƽ��</strong>��<br>"
        Response.Write "      <textarea name='Intro' cols='72' rows='9' style='display:none'>"
        If Trim(rsTrademark("TrademarkIntro") & "") <> "" Then Response.Write Server.HTMLEncode(rsTrademark("TrademarkIntro"))
        Response.Write "      </textarea>"
        Response.Write "      <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Intro' frameborder='1' scrolling='no' width='550' height='250' ></iframe>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "    <tr>"
        Response.Write "      <td colspan='2' align='center' class='tdbg'>"
        Response.Write "      <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "      <input name='TypeSelect' type='hidden' id='TypeSelect' value='SaveModifyTrademark'>"
        Response.Write "      <input name='ID' type='hidden' id='ID' value=" & rsTrademark("TrademarkID") & ">"
        Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsTrademark.Close
    Set rsTrademark = Nothing
End Sub

Sub SaveAddTrademark()
    Dim TrademarkName, ProducerID, Intro, Photo, TrademarkType, Elite
    Dim rsTrademark, sqlTrademark
    TrademarkName = Trim(Request("TrademarkName"))

    If TrademarkName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Ʒ�����Ʋ���Ϊ�գ�</li>"
    Else
        TrademarkName = ReplaceBadChar(TrademarkName)
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    ProducerID = Trim(Request("ProducerID"))
    If IsNull(ProducerID) Then
        ProducerID = 0
    Else
        ProducerID = PE_CLng(ProducerID)
    End If
    Photo = Trim(Request("Photo"))
    Intro = Trim(Request("Intro"))
    TrademarkType = Trim(Request("TrademarkType"))
    Elite = Trim(Request("Elite"))
    
    Set rsTrademark = Server.CreateObject("Adodb.RecordSet")
    sqlTrademark = "Select * from PE_Trademark where ChannelID=" & ChannelID & " and TrademarkName='" & TrademarkName & "'"
    rsTrademark.Open sqlTrademark, Conn, 1, 3
    If Not (rsTrademark.BOF And rsTrademark.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����Ѿ����ڴ�Ʒ�ƣ�</li>"
        rsTrademark.Close
        Set rsTrademark = Nothing
        Exit Sub
    End If
    rsTrademark.addnew
    rsTrademark("ChannelID") = ChannelID
    rsTrademark("ProducerID") = ProducerID
    rsTrademark("TrademarkName") = TrademarkName
    If Photo <> "" Then rsTrademark("TrademarkPhoto") = Photo
    If Intro <> "" Then rsTrademark("TrademarkIntro") = Intro
    rsTrademark("TrademarkType") = PE_CLng(TrademarkType)
    If Elite = "Yes" Then
        rsTrademark("IsElite") = True
    Else
        rsTrademark("IsElite") = False
    End If
    rsTrademark("Passed") = True
    rsTrademark.Update
    rsTrademark.Close
    Set rsTrademark = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark"
End Sub

Sub SaveModifyTrademark()
    Dim TrademarkName, TrademarkID, ProducerID, Intro, Photo, TrademarkType, Elite
    Dim rsTrademark, sqlTrademark
    TrademarkName = Trim(Request("TrademarkName"))
    If TrademarkName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�Ʒ�����ƣ�</li>"
    End If
    TrademarkID = Trim(Request("ID"))
    If TrademarkID <> "" Then
        If InStr(TrademarkID, ",") > 0 Then
            TrademarkID = ReplaceBadChar(TrademarkID)
        Else
            TrademarkID = PE_CLng(TrademarkID)
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�Ʒ��ID��</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    ProducerID = Trim(Request("ProducerID"))
    If IsNull(ProducerID) Then
        ProducerID = 0
    Else
        ProducerID = PE_CLng(ProducerID)
    End If
    Photo = Trim(Request("Photo"))
    Intro = Trim(Request("Intro"))
    TrademarkType = Trim(Request("TrademarkType"))
    Elite = Trim(Request("Elite"))
    
    Set rsTrademark = Server.CreateObject("Adodb.RecordSet")
    sqlTrademark = "Select * from PE_Trademark where TrademarkID=" & TrademarkID
    rsTrademark.Open sqlTrademark, Conn, 1, 3
    If Not (rsTrademark.BOF And rsTrademark.EOF) Then
        rsTrademark("ChannelID") = ChannelID
        rsTrademark("ProducerID") = ProducerID
        rsTrademark("TrademarkName") = TrademarkName
        If Photo <> "" Then rsTrademark("TrademarkPhoto") = Photo
        If Intro <> "" Then rsTrademark("TrademarkIntro") = Intro
        rsTrademark("TrademarkType") = PE_CLng(TrademarkType)
        If Elite = "Yes" Then
            rsTrademark("IsElite") = True
        Else
            rsTrademark("IsElite") = False
        End If
        rsTrademark.Update
    End If
    rsTrademark.Close
    Set rsTrademark = Nothing
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark"
End Sub

Sub DelTrademark()
    Dim TrademarkID
    TrademarkID = Trim(Request("ID"))
    If IsValidID(TrademarkID) = False Then
        TrademarkID = ""
    End If
    If TrademarkID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ����Ʒ��ID</li>"
        Exit Sub
    End If
    If InStr(TrademarkID, ",") > 0 Then
        Conn.Execute ("delete from PE_Trademark where TrademarkID in (" & TrademarkID & ")")
    Else
        Conn.Execute ("delete from PE_Trademark where TrademarkID=" & TrademarkID & "")
    End If
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark"
End Sub

Function GetTrademarkType(TypeID)
    Select Case TypeID
    Case 1
        GetTrademarkType = XmlText("ShowSource", "ShowTrademark/TrademarkType1", "��½Ʒ��")
    Case 2
        GetTrademarkType = XmlText("ShowSource", "ShowTrademark/TrademarkType2", "��̨Ʒ��")
    Case 3
        GetTrademarkType = XmlText("ShowSource", "ShowTrademark/TrademarkType3", "�պ�Ʒ��")
    Case 4
        GetTrademarkType = XmlText("ShowSource", "ShowTrademark/TrademarkType4", "ŷ��Ʒ��")
    Case Else
        GetTrademarkType = XmlText("ShowSource", "ShowTrademark/TrademarkType5", "����Ʒ��")
    End Select
End Function

Function GetProducerName(ProduceID)
    Dim rsProducer
    Set rsProducer = Conn.Execute("Select ProducerID,ProducerName from PE_Producer where ProducerID=" & ProduceID)
    If Not (rsProducer.BOF And rsProducer.EOF) Then
        GetProducerName = rsProducer("ProducerName")
    Else
        GetProducerName = "��"
    End If
    rsProducer.Close
    Set rsProducer = Nothing
End Function

Function GetProducerList(iChannelID, iProducerID)
    Dim rsProducer, strtmp
    Set rsProducer = Conn.Execute("Select ProducerID,ChannelID,ProducerName from PE_Producer where ChannelID=" & iChannelID)
    If Not (rsProducer.BOF And rsProducer.EOF) Then
        Do While Not rsProducer.EOF
            If rsProducer("ProducerID") = iProducerID Then
                            strtmp = strtmp & "<option value=" & rsProducer("ProducerID") & " selected>" & rsProducer("ProducerName") & "</option>"
                        Else
                            strtmp = strtmp & "<option value=" & rsProducer("ProducerID") & ">" & rsProducer("ProducerName") & "</option>"
                        End If
            rsProducer.MoveNext
        Loop
    Else
        strtmp = "<option value=''>��δ��ӳ���</option>"
    End If
    rsProducer.Close
    Set rsProducer = Nothing
    GetProducerList = strtmp
End Function

Sub SetStat(imodetype, istat)
    Dim ItemID, idname
    ItemID = PE_CLng(Trim(Request("ID")))
    If ItemID = 0 Or IsNull(imodetype) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����Ķ���</li>"
        Exit Sub
    End If
    If imodetype = "Producer" Or imodetype = "Trademark" Then
        idname = imodetype & "ID"
    Else
        idname = "ID"
    End If
    Select Case istat
    Case 1
        Conn.Execute ("update PE_" & imodetype & " set Passed=" & PE_False & " where " & idname & "=" & ItemID & "")
    Case 2
        Conn.Execute ("update PE_" & imodetype & " set Passed=" & PE_True & " where " & idname & "=" & ItemID & "")
    Case 3
        Conn.Execute ("update PE_" & imodetype & " set onTop=" & PE_False & " where " & idname & "=" & ItemID & "")
    Case 4
        Conn.Execute ("update PE_" & imodetype & " set onTop=" & PE_True & " where " & idname & "=" & ItemID & "")
    Case 5
        Conn.Execute ("update PE_" & imodetype & " set IsElite=" & PE_False & " where " & idname & "=" & ItemID & "")
    Case 6
        Conn.Execute ("update PE_" & imodetype & " set IsElite=" & PE_True & " where " & idname & "=" & ItemID & "")
    End Select
    Call CloseConn
    Response.Redirect "Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=" & imodetype
End Sub

Function AuthorTemplateList(iTempid)
    Dim rsTemplate, strtmp
    Set rsTemplate = Conn.Execute("select * from PE_Template where ChannelID=0 and TemplateType=10 ")
    If rsTemplate.BOF And rsTemplate.EOF Then
        strtmp = "<option value=0>��δ���ģ��!</option>"
    Else
        If iTempid = 0 Then
            strtmp = "<option value=0 selected>Ĭ��ģ��!</option>"
        Else
            strtmp = "<option value=0>Ĭ��ģ��!</option>"
        End If
        Do While Not rsTemplate.EOF
            strtmp = strtmp & "<option value=" & rsTemplate("TemplateID")
                If rsTemplate("TemplateID") = iTempid Then strtmp = strtmp & " selected"
            strtmp = strtmp & ">" & rsTemplate("TemplateName") & "</option>"
            rsTemplate.MoveNext
        Loop
    End If
    rsTemplate.Close
    Set rsTemplate = Nothing
    AuthorTemplateList = strtmp
End Function
%>
