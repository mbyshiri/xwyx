<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
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

Dim HtmlDir
Dim ManageType, InfoShortName

FileExt_SiteSpecial = arrFileExt(FileExt_SiteSpecial)

HtmlDir = InstallDir & ChannelDir
ManageType = Trim(Request("ManageType"))

Response.Write "<html><head><title>" & ChannelShortName & "ר�����</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
If ChannelID > 0 Then
     Call ShowPageTitle(ChannelName & "����----ר�����", 10004)
Else
     Call ShowPageTitle("ȫվר�����", 10004)
End If
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>"
Response.Write "    <td>"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "ר�������ҳ</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Add'>���" & ChannelShortName & "ר��</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Order'>" & ChannelShortName & "ר������</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Unite'>�ϲ�" & ChannelShortName & "ר��</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Batch'>��������</a>"
If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&ManageType=HTML'><b>����HTML����</b></a>"
End If
Response.Write "</td></tr></table>"

Action = Trim(Request("Action"))
Select Case Action
Case "Add"
    Call AddSpecial
Case "SaveAdd"
    Call SaveAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Del"
    Call DelSpecial
Case "Clear"
    Call ClearSpecial
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "Unite"
    Call ShowUniteForm
Case "UniteSpecial"
    Call UniteSpecial
Case "Batch"
    Call ShowBatch
Case "DoBatch"
    Call DoBatch
Case "Order"
    Call ShowOrder
Case "CreateSpecialDir"
    Call CreateSpecialDir1
Case "CreateAllSpecialDir"
    Call CreateAllSpecialDir
Case "DelSpecialDir"
    Call DelSpecialDir1
Case "DelAllSpecialDir"
    Call DelAllSpecialDir
Case "CreateJS"
    Call CreateJS_Special
    Call WriteSuccessMsg("�Ѿ��ɹ�����ר��JS�ļ���", ComeUrl)
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 20
    strFileName = "Admin_Special.asp?ChannelID=" & ChannelID
    If Request("page") <> "" Then
        CurrentPage = PE_CLng(Request("page"))
    Else
        CurrentPage = 1
    End If

    Dim arrOpenType
    arrOpenType = Array("ԭ���ڴ�", "�´��ڴ�")

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>ר������</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>ר��Ŀ¼</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>�򿪷�ʽ</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>�Ƽ�ר��</strong></td>"
    Response.Write "    <td width='200' align='center'><strong>ר����ʾ</strong></td>"
    Response.Write "    <td width='100' height='22' align='center'><strong>�������</strong></td>"
    Response.Write "  </tr>"
    Dim rsSpecial, sql
    sql = "select * from PE_Special where ChannelID=" & ChannelID & " order by OrderID"
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 1
    If rsSpecial.BOF And rsSpecial.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>û���κ�ר��</td></tr>"
        totalPut = 0
    Else
        totalPut = rsSpecial.RecordCount
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
                rsSpecial.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim i
        i = 0
        Do While Not rsSpecial.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If ChannelID > 0 Then
                Response.Write "    <td align='center'><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ManageType=Special&SpecialID=" & rsSpecial("SpecialID") & "' title='�����������ר���" & InfoShortName & "'>" & rsSpecial("SpecialName") & "</a></td>"
            Else
                Response.Write "    <td align='center'>" & rsSpecial("SpecialName") & "</td>"
            End If
            Response.Write "    <td width='80' align='center'>" & rsSpecial("SpecialDir") & "</td>"
            Response.Write "    <td width='80' align='center'>" & arrOpenType(rsSpecial("OpenType")) & "</td>"
            Response.Write "    <td width='80' align='center'>"
            If rsSpecial("IsElite") = True Then
                Response.Write "<font color=green>��</font>"
            Else
                Response.Write "��"
            End If
            Response.Write "</td>"
            Response.Write "    <td width='200'>" & PE_HTMLEncode(rsSpecial("Tips")) & "</td>"
            If ManageType = "HTML" Then
                Response.Write "    <td width='240' align='center'>"
                Response.Write "<a href='Admin_Create" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=CreateSpecial&SpecialID=" & rsSpecial("SpecialID") & "' title='���ɱ�ר���" & InfoShortName & "�б�HTMLҳ��'>�����б�ҳ</a>&nbsp;|&nbsp;"
                Response.Write "<a href='" & HtmlDir & "/Special/" & rsSpecial("SpecialDir") & "/Index.html' title='�鿴��ר���" & InfoShortName & "�б�HTMLҳ��' target='_blank'>�鿴�б�ҳ</a>"
                If Not fso.FolderExists(Server.MapPath(HtmlDir & "/Special/" & rsSpecial("SpecialDir"))) Then
                    Response.Write "&nbsp;|&nbsp;<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=CreateSpecialDir&SpecialID=" & rsSpecial("SpecialID") & "' title='���ɱ�ר���Ŀ¼'>����ר��Ŀ¼</a>"
                Else
                    Response.Write "&nbsp;|&nbsp;<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=DelSpecialDir&SpecialID=" & rsSpecial("SpecialID") & "' title='�˲�����ɾ����ר���Ŀ¼' onclick=""return confirm('�˲�����ɾ����ר���Ŀ¼���������������Ŀ¼��');"">ɾ��ר��Ŀ¼</a>"
                End If
            Else
                Response.Write "    <td width='100' align='center'>"
                Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&action=Modify&SpecialID=" & rsSpecial("SpecialID") & "'>�޸�</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Del&SpecialID=" & rsSpecial("SpecialID") & "' onClick=""return confirm('ȷ��Ҫɾ����ר����ɾ����ר���ԭ���ڴ�ר���" & InfoShortName & "���������κ�ר�⡣');"">ɾ��</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Clear&SpecialID=" & rsSpecial("SpecialID") & "' onClick=""return confirm('ȷ��Ҫ��մ�ר���е�" & InfoShortName & "�𣿱�������ԭ���ڴ�ר���" & InfoShortName & "��Ϊ�������κ�ר�⡣');"">���</a>"
            End If
            Response.Write "</td></tr>"
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsSpecial.MoveNext
        Loop
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��ר��", True)

    If ManageType = "HTML" Then
        Response.Write "<br><table align='center'><tr><form name='form1' action='Admin_Special.asp' method='post'><td>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='CreateAllSpecialDir'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='��������ר���Ŀ¼' style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form><form name='form2' action='Admin_Create" & ModuleName & ".asp' method='post'><td><input name='CreateType' type='hidden' value='2'>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='CreateSpecial'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='��������ר���" & InfoShortName & "�б�ҳ' style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form><form name='form4' action='Admin_Special.asp' method='post'><td><input name='ManageType' type='hidden' value='HTML'>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='DelAllSpecialDir'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='ɾ������ר���Ŀ¼' onclick=""return confirm('�˲�����ɾ������ר���Ŀ¼���������������Ŀ¼��������ϵͳ�е�ר���б��ļ��������ң�����ʹ�ô˹�����ɾ������Ŀ¼��Ȼ���������ɡ�');"" style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form></tr></table><br>"
        Response.Write "<b>ע�⣺</b><br>&nbsp;&nbsp;&nbsp;&nbsp;1����������HTML����֮ǰ������ȷ���Ѿ���������ר���Ŀ¼��������ܻᵼ�����ɳ�����ר��Ŀ¼Ϊ��ɫ����ʾ��ר�⻹û�д�����ص�Ŀ¼����ʹ�á�����ר��Ŀ¼���������´�����ר���Ŀ¼��"
        Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;2����Ϊ���ɲ�����ķѴ����ķ�������Դ����������Ҫ�൱��ʱ�䡣<font color=red>�����ɹ�����ǧ��Ҫˢ��ҳ�棡����</font>ͬʱ�����Ҿ�������վ�������Ƚ�Сʱ���С���������Ҫʹ���������ɹ��ܡ�"
    Else
        Response.Write "<table width='100%'><tr><form name='form1' action='Admin_Special.asp' method='post'><td align='center'>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='CreateJS'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='ˢ��ר��JS' style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form></tr></table>"
        If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
            Response.Write "<br><b>ע�⣺</b><br>&nbsp;&nbsp;&nbsp;&nbsp;��ר��Ŀ¼Ϊ��ɫ����ʾ��ר�⻹û�д�����ص�Ŀ¼���뵽������HTML����ҳ��ʹ�á�����ר��Ŀ¼���������´�����ר���Ŀ¼��<br>"
        End If
    End If
End Sub

Sub ShowOrder()
    Dim iCount, i, j
    Dim rsSpecial, sql
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    sql = "select * from PE_Special where ChannelID=" & ChannelID & " Order by OrderID"
    rsSpecial.Open sql, Conn, 1, 1
    iCount = rsSpecial.RecordCount
    j = 1
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4' align='center'><strong>" & ChannelShortName & "ר������</strong></td>"
    Response.Write "  </tr>"
    Do While Not rsSpecial.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""> "
        Response.Write "    <td align='center'>" & rsSpecial("SpecialName") & "</td>"
        Response.Write "    <form action='Admin_Special.asp?Action=UpOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If j > 1 Then
            Response.Write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=SpecialID value=" & rsSpecial("SpecialID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsSpecial("OrderID") & ">&nbsp;<input type=submit name=Submit value=�޸�>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form>"
        Response.Write "    <form action='Admin_Special.asp?Action=DownOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If iCount > j Then
            Response.Write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=SpecialID value=" & rsSpecial("SpecialID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsSpecial("OrderID") & ">&nbsp;<input type=submit name=Submit value=�޸�>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "      <td width='200' align='center'>&nbsp;</td>"
        Response.Write "    </form>"
        Response.Write "  </tr>"
        j = j + 1
        rsSpecial.MoveNext
    Loop
    Response.Write "</table> "
    rsSpecial.Close
    Set rsSpecial = Nothing
End Sub

Sub AddSpecial()

    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Special.asp?ChannelID=" & ChannelID & "'>ר�����</a>&nbsp;&gt;&gt;&nbsp;���ר��</td></tr></table>"
    Response.Write "<form method='post' action='Admin_Special.asp' name='form1'>"

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��������</td>" & vbCrLf
    End If
    Response.Write "   <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "   <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>ר�����ƣ�</strong></td>"
    Response.Write "      <td class='tdbg'><input name='SpecialName' type='text' id='SpecialName' size='49' maxlength='30'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>ר��Ŀ¼��</strong><br>ֻ����Ӣ�ģ����ܴ��ո��\������/���ȷ��š�<br>��������Ҫ������֧��FSO������ʹ��ķ�������֧��FSO��Ҳ������¼�룬��Ϊ�����ڻ��˿ռ����������ɡ�</td>"
    Response.Write "      <td class='tdbg'><input name='SpecialDir' type='text' id='SpecialDir' size='49' maxlength='30'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>ר��ͼƬ��</strong></td>"
    Response.Write "      <td class='tdbg'><input name='SpecialPicUrl' type='text' id='SpecialPicUrl' size='49' maxlength='200'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>�򿪷�ʽ��</strong></td>"
    Response.Write "      <td><input name='OpenType' type='radio' value='0' checked>��ԭ���ڴ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1'>���´��ڴ�</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>�Ƿ�Ϊ�Ƽ�ר�⣺</strong></td>"
    Response.Write "      <td><input name='IsElite' type='radio' value='True' checked>��&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>ר����ʾ��</strong><br>�������ר��������ʱ����ʾ�趨����ʾ���֣���֧��HTML��</td>"
    Response.Write "      <td class='tdbg'><textarea name='Tips' cols='60' rows='3' id='Tips'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>ר��˵����</strong><br>����ר��ҳ��ר�����˵����֧��HTML��</td>"
    Response.Write "      <td class='tdbg'><textarea name='Readme' cols='60' rows='3' id='Readme'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>ÿҳ��ʾ��" & InfoShortName & "����</strong></td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, 20) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='350' class='tdbg5'><strong>Ĭ����ɫ���</strong><br>���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "      <td class='tdbg'><select name='SkinID'>" & GetSkin_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>�������ģ�壺</strong><br>���ģ���а����˰�����Ƶİ�ʽ����Ϣ�������������ӵ����ģ�壬���ܻᵼ�¡�ר����ɫ���ʧЧ��</td>"
    Response.Write "      <td class='tdbg'><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 4, 0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write " </tbody>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Call EditCustom_Content("Add", "", "Special")
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input  type='submit' name='Submit' value=' �� �� '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Special.asp'"" style='cursor:hand;'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim SpecialID, rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ר��ID��</li>"
        Exit Sub
    End If
    sql = "Select * from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ר�⣬�����Ѿ���ɾ����</li>"
    Else
        Response.Write "<script language='javascript'>" & vbCrLf
        Response.Write "var tID=0;" & vbCrLf
        Response.Write "function ShowTabs(ID){" & vbCrLf
        Response.Write "  if(ID!=tID){" & vbCrLf
        Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
        Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
        Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
        Response.Write "    Tabs[ID].style.display='';" & vbCrLf
        Response.Write "    tID=ID;" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "</script>" & vbCrLf
        Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Special.asp?ChannelID=" & ChannelID & "'>ר�����</a>&nbsp;&gt;&gt;&nbsp;�޸�ר�����ã�<font color='red'>" & rsSpecial("SpecialName") & "</td></tr></table>"
        Response.Write "<form method='post' action='Admin_Special.asp' name='form1'>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
        Response.Write "  <tr align='center' height='24'>" & vbCrLf
        Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
        If IsCustom_Content = True And ModuleType <> 6 Then
            Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��������</td>" & vbCrLf
        End If
        Response.Write "   <td>&nbsp;</td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
        Response.Write "   <tbody id='Tabs' style='display:'>" & vbCrLf
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>ר�����ƣ�</strong></td>"
        Response.Write "      <td class='tdbg'><input name='SpecialName' type='text' id='SpecialName' value='" & rsSpecial("SpecialName") & "' size='49' maxlength='30'><input name='SpecialID' type='hidden' id='SpecialID' value='" & rsSpecial("SpecialID") & "'></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>ר��Ŀ¼��</strong></td>"
        Response.Write "      <td class='tdbg'><input name='SpecialDir' type='text' id='SpecialDir' value='" & rsSpecial("SpecialDir") & "' size='49' maxlength='30' disabled></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>ר��ͼƬ��</strong></td>"
        Response.Write "      <td class='tdbg'><input name='SpecialPicUrl' type='text' id='SpecialPicUrl' value='" & rsSpecial("SpecialPicUrl") & "' size='49' maxlength='200'>&nbsp;</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>�򿪷�ʽ��</strong></td>"
        Response.Write "      <td><input name='OpenType' type='radio' value='0'  " & RadioValue(rsSpecial("OpenType"), 0) & ">��ԭ���ڴ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1' " & RadioValue(rsSpecial("OpenType"), 1) & ">���´��ڴ�</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>�Ƿ�Ϊ�Ƽ�ר�⣺</strong></td>"
        Response.Write "      <td><input name='IsElite' type='radio' value='True' " & RadioValue(rsSpecial("IsElite"), True) & ">��&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'" & RadioValue(rsSpecial("IsElite"), False) & ">��</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>ר����ʾ��</strong><br>�������ר��������ʱ����ʾ�趨����ʾ���֣���֧��HTML��</td>"
        Response.Write "      <td class='tdbg'><textarea name='Tips' cols='60' rows='3' id='Tips'>" & rsSpecial("Tips") & "</textarea></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>ר��˵����</strong><br>����ר��ҳ��ר�����˵����֧��HTML��</td>"
        Response.Write "      <td class='tdbg'><textarea name='Readme' cols='60' rows='3' id='Readme'>" & rsSpecial("Readme") & "</textarea></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>ÿҳ��ʾ��" & InfoShortName & "����</strong></td>"
        Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, rsSpecial("MaxPerPage")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>Ĭ����ɫ���</strong><br>���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
        Response.Write "      <td class='tdbg'><select name='SkinID'>" & GetSkin_Option(rsSpecial("SkinID")) & "</select></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>�������ģ�壺</strong><br>���ģ���а����˰�����Ƶİ�ʽ����Ϣ�������������ӵ����ģ�壬���ܻᵼ�¡�ר����ɫ���ʧЧ��</td>"
        Response.Write "      <td class='tdbg'><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 4, rsSpecial("TemplateID")) & "</select></td>"
        Response.Write "    </tr>"
        Response.Write " </tbody>" & vbCrLf
        If IsCustom_Content = True And ModuleType <> 6 Then
            Call EditCustom_Content("Modify", rsSpecial("Custom_Content"), "Special")
        End If
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModify'>"
        Response.Write "        <input  type='submit' name='Submit' value='�����޸Ľ��'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp; "
        Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Special.asp'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
End Sub

Sub ShowUniteForm()
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'> "
    Response.Write "    <td height='22' colspan='3' align='center'><strong>�ϲ�" & ChannelShortName & "ר��</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='100'><form name='myform' method='post' action='Admin_Special.asp' onSubmit='return ConfirmUnite();'>"
    Response.Write "        &nbsp;&nbsp;��ר�� <select name='SpecialID' id='SpecialID'>" & GetSpecial_Option(0) & "</select> �ϲ��� <select name='TargetSpecialID' id='TargetSpecialID'>" & GetSpecial_Option(0) & "</select>"
    Response.Write "        <br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='UniteSpecial'>"
    Response.Write "        <input type='submit' name='Submit' value=' �ϲ�ר�� ' style='cursor:hand;'>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Special.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "      </form></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='60'><strong>ע�����</strong><br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;���в��������棬�����ز���������<br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;������ͬһ��ר���ڽ��в�����<br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;�ϲ�������ָ����ר�⽫��ɾ��������" & InfoShortName & "��ת�Ƶ�Ŀ��ר���С�</td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function ConfirmUnite(){" & vbCrLf
    Response.Write "  if (document.myform.SpecialID.value==document.myform.TargetSpecialID.value){" & vbCrLf
    Response.Write "    alert('�벻Ҫ����ͬר���ڽ��в�����');" & vbCrLf
    Response.Write " document.myform.TargetSpecialID.focus();" & vbCrLf
    Response.Write " return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub ShowBatch()
    Response.Write "<form name='form1' method='post' action='Admin_Special.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='3' align='center'><strong>��������" & ChannelShortName & "ר������</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' valign='top'><font color='red'>��ʾ��</font>���԰�ס��Shift��<br>��Ctrl�������ж��ר���ѡ��<br>"
    Response.Write "      <select name='SpecialID' size='2' multiple style='height:200px;width:200px;'>" & GetSpecial_Option(0) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  ѡ������ר��  ' onclick='SelectAll()'><br>"
    Response.Write "      <input type='button' name='Submit' value='ȡ��ѡ������ר��' onclick='UnSelectAll()'></div></td>"
    Response.Write "      <td>"
    Response.Write "     <table id='SpecialSettings' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF' style='display:'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyOpenType' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>�򿪷�ʽ��</strong></td>"
    Response.Write "      <td><input name='OpenType' type='radio' value='0' checked>��ԭ���ڴ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1'>���´��ڴ�</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyIsElite' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>�Ƿ�Ϊ�Ƽ�ר�⣺</strong></td>"
    Response.Write "      <td><input name='IsElite' type='radio' value='True' checked>��&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyMaxPerPage' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>ÿҳ��ʾ��" & InfoShortName & "����</strong></td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, 20) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifySkinID' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>ר����ɫ���</strong><br>���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "      <td><select name='SkinID' id='SkinID'>" & GetSkin_Option(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyTemplateID' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>�������ģ�壺</strong><br>���ģ���а�����ר����Ƶİ�ʽ����Ϣ�������������ӵ����ģ�壬���ܻᵼ�¡�ר����ɫ���ʧЧ�� </td>"
    Response.Write "      <td><select name='TemplateID' id='TemplateID'>" & GetTemplate_Option(ChannelID, 4, 0) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='3' align='center'><input name='Action' type='hidden' id='Action' value='DoBatch'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Submit' type='submit' value=' ִ�������� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Special.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'></td></tr>"
    Response.Write "  </table>"
    Response.Write "</td></tr></table>"
    Response.Write "</form>" & vbCrLf
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.form1.SpecialID.length;i++){" & vbCrLf
    Response.Write "    document.form1.SpecialID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.form1.SpecialID.length;i++){" & vbCrLf
    Response.Write "    document.form1.SpecialID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub SaveAdd()
    Dim SpecialName, SpecialDir, SpecialID, OrderID
    Dim rsSpecial, sql
    SpecialName = ReplaceBadChar(Trim(Request("SpecialName")))
    SpecialDir = Trim(Request("SpecialDir"))
    If SpecialName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ר�����Ʋ���Ϊ�գ�</li>"
    End If
    If SpecialDir = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ר��Ŀ¼����Ϊ�գ�</li>"
    Else
        If IsValidStr(SpecialDir) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ר��Ŀ¼��ֻ����Ӣ�ģ�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    SpecialID = GetNewID("PE_Special", "SpecialID")
    OrderID = GetMinID("PE_Special", "OrderID")
    Conn.Execute ("update PE_Special set OrderID=OrderID+1 where ChannelID=" & ChannelID & "")
    
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open "Select * from PE_Special Where ChannelID=" & ChannelID & " and (SpecialName='" & SpecialName & "' or SpecialDir='" & SpecialDir & "')", Conn, 1, 3
    If Not (rsSpecial.BOF And rsSpecial.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ר�����ƻ�ר��Ŀ¼�Ѿ����ڣ�</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
        Exit Sub
    End If
    
    rsSpecial.addnew
    rsSpecial("SpecialID") = SpecialID
    rsSpecial("ChannelID") = ChannelID
    rsSpecial("OrderID") = OrderID
    rsSpecial("SpecialName") = SpecialName
    rsSpecial("SpecialDir") = SpecialDir
    rsSpecial("SpecialPicUrl") = Trim(Request("SpecialPicUrl"))
    rsSpecial("IsElite") = PE_CBool(Trim(Request("IsElite")))
    rsSpecial("OpenType") = PE_CLng(Trim(Request("OpenType")))
    rsSpecial("Tips") = Trim(Request("Tips"))
    rsSpecial("ReadMe") = Trim(Request("ReadMe"))
    rsSpecial("MaxPerPage") = PE_CLng(Trim(Request("MaxPerPage")))
    rsSpecial("SkinID") = PE_CLng(Trim(Request("SkinID")))
    rsSpecial("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
    '��������
    Dim Custom_Num, Custom_Content, i
    Custom_Num = PE_CLng(Request.Form("Custom_Num"))
    If Custom_Num <> 0 Then
        For i = 1 To Custom_Num
            If i <> 1 Then
                Custom_Content = Custom_Content & "{#$$$#}"
            End If
            Custom_Content = Custom_Content & Trim(Request("Custom_Content" & i))
        Next
    End If
    rsSpecial("Custom_Content") = Custom_Content
    rsSpecial.Update
    rsSpecial.Close
    Set rsSpecial = Nothing
    Conn.Execute ("update PE_Channel set SpecialCount=SpecialCount+1 where ChannelID=" & ChannelID & "")
    Call CreateJS_Special
    If UseCreateHTML > 0 Then
        Call CreateSpecialDir(SpecialDir)
    End If
    Call ClearSiteCache(ChannelID)
    Call CloseConn
    Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
End Sub

Sub CreateSpecialDir(DirName)
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim tmpDir
    tmpDir = InstallDir & ChannelDir & "/Special"
    If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.CreateFolder Server.MapPath(tmpDir)
    End If
    tmpDir = tmpDir & "/" & DirName
    If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.CreateFolder Server.MapPath(tmpDir)
    End If
End Sub

Sub CreateSpecialDir1()
    Dim SpecialID, rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ר��ID��</li>"
        Exit Sub
    End If
    sql = "Select SpecialDir from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ר�⣬�����Ѿ���ɾ����</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
    Else
        Call CreateSpecialDir(rsSpecial(0))
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    If FoundErr = False Then Call WriteSuccessMsg("����ר��Ŀ¼�ɹ���", ComeUrl)
End Sub

Sub SaveModify()
    Dim SpecialID, SpecialName
    Dim rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ר��ID��</li>"
        Exit Sub
    End If
    SpecialName = ReplaceBadChar(Trim(Request.Form("SpecialName")))
    If SpecialName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ר�����Ʋ���Ϊ�գ�</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    sql = "Select * from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ר�⣬�����Ѿ���ɾ����</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
    Else
        rsSpecial("SpecialName") = SpecialName
        rsSpecial("SpecialPicUrl") = Trim(Request("SpecialPicUrl"))
        rsSpecial("IsElite") = PE_CBool(Trim(Request("IsElite")))
        rsSpecial("OpenType") = PE_CLng(Trim(Request("OpenType")))
        rsSpecial("Tips") = Trim(Request("Tips"))
        rsSpecial("ReadMe") = Trim(Request("ReadMe"))
        rsSpecial("MaxPerPage") = PE_CLng(Trim(Request("MaxPerPage")))
        rsSpecial("SkinID") = PE_CLng(Trim(Request("SkinID")))
        rsSpecial("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
        '��������
        Dim Custom_Num, Custom_Content, i
        Custom_Num = PE_CLng(Request.Form("Custom_Num"))
        If Custom_Num <> 0 Then
            For i = 1 To Custom_Num
                If i <> 1 Then
                    Custom_Content = Custom_Content & "{#$$$#}"
                End If
                Custom_Content = Custom_Content & Trim(Request("Custom_Content" & i))
            Next
        End If
        rsSpecial("Custom_Content") = Custom_Content
        rsSpecial.Update
        rsSpecial.Close
        Set rsSpecial = Nothing
        Call CreateJS_Special
        Call ClearSiteCache(ChannelID)
        Call CloseConn
        Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
    End If
End Sub

Sub DelSpecial()
    Dim SpecialID
    SpecialID = Trim(Request("SpecialID"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ר��ID��</li>"
        Exit Sub
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If UseCreateHTML > 0 Then
        Dim trs, SpecialDir
        Set trs = Conn.Execute("select SpecialDir from PE_Special where SpecialID=" & SpecialID)
        SpecialDir = trs(0)
        Set trs = Nothing
        Call DelSpecialDir(SpecialDir)
    End If
    If FoundErr = True Then Exit Sub

    Dim rsInfo
    Set rsInfo = Conn.Execute("select * from PE_InfoS where SpecialID=" & SpecialID & " order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & rsInfo("ModuleType") & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 1 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=0 where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing

    Conn.Execute ("delete from PE_Special where SpecialID=" & SpecialID)
    Conn.Execute ("update PE_Channel set SpecialCount=SpecialCount-1 where ChannelID=" & ChannelID & "")
    Call CreateJS_Special
    Call CloseConn
    Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
End Sub

Sub DelSpecialDir(DirName)
    On Error Resume Next
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim tmpDir
    tmpDir = InstallDir & ChannelDir & "/Special/" & DirName
    If fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.DeleteFolder Server.MapPath(tmpDir)
    End If
    If Err Then
        Error.Clear
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ר��Ŀ¼�޷�ɾ�����������ļ�����ʹ���С����Ժ����ԣ�</li>"
    End If
End Sub

Sub DelSpecialDir1()
    Dim SpecialID, rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ר��ID��</li>"
        Exit Sub
    End If
    sql = "Select SpecialDir from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ר�⣬�����Ѿ���ɾ����</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
    Else
        Call DelSpecialDir(rsSpecial(0))
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing

    If FoundErr = False Then Call WriteSuccessMsg("ɾ��ר��Ŀ¼�ɹ���", ComeUrl)
End Sub

Sub ClearSpecial()
    Dim SpecialID
    SpecialID = Trim(Request("SpecialID"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�ר��ID��</li>"
        Exit Sub
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If UseCreateHTML > 0 Then
        Dim trs, SpecialDir
        Set trs = Conn.Execute("select SpecialDir from PE_Special where SpecialID=" & SpecialID)
        SpecialDir = trs(0)
        Set trs = Nothing
        Call ClearSpecialDir(SpecialDir)
    End If
    If FoundErr = True Then Exit Sub

    Dim rsInfo
    Set rsInfo = Conn.Execute("select * from PE_InfoS where SpecialID=" & SpecialID & " order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & rsInfo("ModuleType") & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 1 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=0 where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing
    
    Call CloseConn
    Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
End Sub

Sub ClearSpecialDir(DirName)
    On Error Resume Next
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim tmpDir
    tmpDir = InstallDir & ChannelDir & "/Special/" & DirName
    If fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.DeleteFile Server.MapPath(tmpDir) & "\*.*"
    End If
    If Err Then
        Error.Clear
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�޷���ȫ�����ר��Ŀ¼�µ��ļ����������ļ�����ʹ���С����Ժ����ԣ�</li>"
    End If
End Sub

Sub UpOrder()
    Dim SpecialID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsSpecial
    SpecialID = Trim(Request("SpecialID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        cOrderID = CInt(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        MoveNum = CInt(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ��Ҫ���������֣�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_Special")
    MaxOrderID = mrs(0) + 1
    '�Ƚ���ǰר��������󣬰�����ר��
    Conn.Execute ("update PE_Special set OrderID=" & MaxOrderID & " where SpecialID=" & SpecialID)
    
    'Ȼ��λ�ڵ�ǰר�����ϵ�ר���OrderID���μ�һ����ΧΪҪ����������
    sqlOrder = "select * from PE_Special where OrderID<" & cOrderID & " order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '�����ǰר���Ѿ��������棬�������ƶ�
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '�õ�Ҫ����λ�õ�OrderID��������ר��
        Conn.Execute ("update PE_Special set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    'Ȼ���ٽ���ǰר�������Ƶ���Ӧλ�ã�������ר��
    Conn.Execute ("update PE_Special set OrderID=" & tOrderID & " where SpecialID=" & SpecialID)
    Call CreateJS_Special
    Call CloseConn
    Response.Redirect "Admin_Special.asp?Action=Order&ChannelID=" & ChannelID
End Sub

Sub DownOrder()
    Dim SpecialID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsSpecial, PrevID, NextID
    SpecialID = Trim(Request("SpecialID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        cOrderID = CInt(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        MoveNum = CInt(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ��Ҫ���������֣�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_Special")
    MaxOrderID = mrs(0) + 1
    '�Ƚ���ǰר��������󣬰�����ר��
    Conn.Execute ("update PE_Special set OrderID=" & MaxOrderID & " where SpecialID=" & SpecialID)
    
    'Ȼ��λ�ڵ�ǰר�����µ�ר���OrderID���μ�һ����ΧΪҪ�½�������
    sqlOrder = "select * from PE_Special where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '�����ǰר���Ѿ��������棬�������ƶ�
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '�õ�Ҫ����λ�õ�OrderID��������ר��
        Conn.Execute ("update PE_Special set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    'Ȼ���ٽ���ǰר�������Ƶ���Ӧλ�ã�������ר��
    Conn.Execute ("update PE_Special set OrderID=" & tOrderID & " where SpecialID=" & SpecialID)
    Call CreateJS_Special
    Call CloseConn
    Response.Redirect "Admin_Special.asp?Action=Order&ChannelID=" & ChannelID
End Sub

Sub UniteSpecial()
    Dim SpecialID, TargetSpecialID, SuccessMsg
    SpecialID = Trim(Request("SpecialID"))
    TargetSpecialID = Trim(Request("TargetSpecialID"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�ϲ���ר�⣡</li>"
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If TargetSpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ŀ��ר�⣡</li>"
    Else
        TargetSpecialID = PE_CLng(TargetSpecialID)
    End If
    If SpecialID = TargetSpecialID Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�벻Ҫ����ͬר���ڽ��в���</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim rsInfo
    Set rsInfo = Conn.Execute("select * from PE_InfoS where SpecialID=" & SpecialID & " order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & rsInfo("ModuleType") & " and SpecialID=" & TargetSpecialID & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 0 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=" & TargetSpecialID & " where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing
    

    'ɾ�����ϲ�ר��
    Conn.Execute ("delete from PE_Special where SpecialID=" & SpecialID)
    Conn.Execute ("update PE_Channel set SpecialCount=SpecialCount-1 where ChannelID=" & ChannelID & "")
    SuccessMsg = "ר��ϲ��ɹ����Ѿ������ϲ�ר�����������ת��Ŀ��ר���С�"
    Call CreateJS_Special
    Call WriteSuccessMsg(SuccessMsg, ComeUrl)
End Sub

Sub CreateAllSpecialDir()
    'On Error Resume Next
    If Not fso.FolderExists(Server.MapPath(HtmlDir & "/Special")) Then
        fso.CreateFolder Server.MapPath(HtmlDir & "/Special")
    End If

    Dim sqlSpecial, rsSpecial, i, iDepth
    sqlSpecial = "select * from PE_Special where ChannelID=" & ChannelID & " order by OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    Do While Not rsSpecial.EOF
        If Not fso.FolderExists(Server.MapPath(HtmlDir & "/Special/" & rsSpecial("SpecialDir"))) Then
            fso.CreateFolder Server.MapPath(HtmlDir & "/Special/" & rsSpecial("SpecialDir"))
        End If
        rsSpecial.MoveNext
    Loop
    rsSpecial.Close
    Set rsSpecial = Nothing
    Call WriteSuccessMsg("��������ר���Ŀ¼�ɹ���", ComeUrl)
End Sub

Sub DelAllSpecialDir()
    On Error Resume Next
    Dim theFolder, theSubFolder, strFolderName
    Set theFolder = fso.GetFolder(Server.MapPath(HtmlDir & "/Special"))
    For Each theSubFolder In theFolder.SubFolders
        strFolderName = theSubFolder.name
        theSubFolder.Delete
        If Err Then
            Err.Clear
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ɾ��Ŀ¼��" & strFolderName & "ʧ�ܣ����ܵ�ǰĿ¼����ʹ���С����Ժ����ԣ�</li>"
        End If
    Next
    If FoundErr <> True Then
        Call WriteSuccessMsg("ɾ��������Ŀ��Ŀ¼�ɹ���", ComeUrl)
    End If
End Sub

Sub CreateJS_Special()

    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    Dim hf, strSpecial, SpecialPath
    'ȫվר��
    If ChannelID = 0 Then
        SpecialPath = InstallDir & "js"
        JS_SpecialNum = 10
    Else
        SpecialPath = InstallDir & ChannelDir & "/js"
    End If

    If Not fso.FolderExists(Server.MapPath(SpecialPath)) Then
        fso.CreateFolder (Server.MapPath(SpecialPath))
    End If
    
    strSpecial = GetSpecialList(JS_SpecialNum)
    Call WriteToFile(SpecialPath & "/ShowSpecialList.js", "document.write(""" & strSpecial & """);")
End Sub

Sub DoBatch()
    Dim SpecialID
    Dim sql, rsSpecial, i, trs
    SpecialID = Trim(Request("SpecialID"))
	If IsValidID(SpecialID) = False Then
		SpecialID = ""
	End If

    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��Ҫ�����޸����õ�ר�⣡</li>"
    End If


    If FoundErr = True Then
        Exit Sub
    End If
    
    If InStr(SpecialID, ",") > 0 Then
        SpecialID = ReBuild(SpecialID)
        sql = "select * from PE_Special where SpecialID in (" & SpecialID & ")"
    Else
        sql = "select * from PE_Special where SpecialID=" & SpecialID
    End If
    Set rsSpecial = Server.CreateObject("Adodb.recordset")
    rsSpecial.Open sql, Conn, 1, 3
    Do While Not rsSpecial.EOF
        If Trim(Request("ModifyIsElite")) = "Yes" Then rsSpecial("IsElite") = PE_CBool(Trim(Request("IsElite")))
        If Trim(Request("ModifyOpenType")) = "Yes" Then rsSpecial("OpenType") = PE_CLng(Trim(Request("OpenType")))
        If Trim(Request("ModifyMaxPerPage")) = "Yes" Then rsSpecial("MaxPerPage") = PE_CLng(Trim(Request("MaxPerPage")))
        If Trim(Request("ModifySkinID")) = "Yes" Then rsSpecial("SkinID") = PE_CLng(Trim(Request("SkinID")))
        If Trim(Request("ModifyTemplateID")) = "Yes" Then rsSpecial("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
        rsSpecial.Update
        rsSpecial.MoveNext
        Set trs = Nothing
    Loop
    rsSpecial.Close
    Set rsSpecial = Nothing
    Call ClearSiteCache(ChannelID)
    Call CreateJS_Special
    Call WriteSuccessMsg("��������ר�����Գɹ���", ComeUrl)
End Sub

Function ReBuild(ByVal iSpecialID)
    Dim arrSpecialID, SpecialArr, i
    arrSpecialID = Split(iSpecialID, ",")
    SpecialArr = ""
    For i = 0 To UBound(arrSpecialID)
        If Trim(arrSpecialID(i)) <> "" And Trim(arrSpecialID(i)) <> "0" Then
            If SpecialArr = "" Then
                SpecialArr = arrSpecialID(i)
            Else
                SpecialArr = SpecialArr & "," & arrSpecialID(i)
            End If
        End If
    Next
    ReBuild = SpecialArr
End Function



'=================================================
'��������GetSpecialList
'��  �ã��������б�ʽ��ʾר������
'��  ����SpecialNum  ------�����ʾ���ٸ�ר������
'=================================================
Function GetSpecialList(SpecialNum)
    Dim sqlSpecial, rsSpecial, strSpecial, i
    If SpecialNum <= 0 Or SpecialNum > 100 Then
        SpecialNum = 10
    End If
    sqlSpecial = "select SpecialID,SpecialName,SpecialDir,Tips from PE_Special where ChannelID=" & ChannelID & " and IsElite=" & PE_True & " order by OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    If rsSpecial.BOF And rsSpecial.EOF Then
        strSpecial = "&nbsp;û���κ�ר����Ŀ"
    Else

        i = 0
        Do While Not rsSpecial.EOF
            If ChannelID = 0 Then
                If FileExt_SiteSpecial <> ".asp" Then
                    strSpecial = strSpecial & "<li><a href='" & InstallDir & "Special/" & rsSpecial(2) & "/Index" & FileExt_SiteSpecial & "' title='" & rsSpecial(3) & "'>" & rsSpecial(1) & "</a></li>"
                Else
                    strSpecial = strSpecial & "<li><a href='" & InstallDir & "ShowSpecial.asp?SpecialID=" & rsSpecial(0) & "' title='" & Trim(nohtml(rsSpecial(3))) & "'>" & rsSpecial(1) & "</a></li>"
                End If
            Else
                If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
                    strSpecial = strSpecial & "<li><a href='" & ChannelUrl & "/Special/" & rsSpecial(2) & "/Index" & FileExt_List & "' title='" & rsSpecial(3) & "'>" & rsSpecial(1) & "</a></li>"
                Else
                    strSpecial = strSpecial & "<li><a href='" & ChannelUrl & "/ShowSpecial.asp?SpecialID=" & rsSpecial(0) & "' title='" & Trim(nohtml(rsSpecial(3))) & "'>" & rsSpecial(1) & "</a></li>"
                End If
            End If

            rsSpecial.MoveNext
            i = i + 1
            If i >= SpecialNum Then Exit Do
        Loop
    End If
    If Not rsSpecial.EOF Then
        strSpecial = strSpecial & "<p align='right'><a href='" & ChannelUrl & "/SpecialList.asp'>����ר��</a></p>"
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    GetSpecialList = strSpecial
End Function
%>
