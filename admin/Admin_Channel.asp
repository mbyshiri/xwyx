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
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Channel"   '����Ȩ��

rsGetAdmin.Close
Set rsGetAdmin = Nothing

Response.Write "<html><head><title>Ƶ������</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("Ƶ �� �� ��", 10002)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Channel.asp'>Ƶ��������ҳ</a>&nbsp;|&nbsp;<a href='Admin_Channel.asp?Action=Add'>�����Ƶ��</a>&nbsp;|&nbsp;<a href='Admin_Channel.asp?Action=Order'>Ƶ������</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Action = Trim(Request("Action"))
Select Case Action
Case "Add"
    Call AddChannel
Case "SaveAdd"
    Call SaveAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Disabled"
    Call DisabledChannel(0)
Case "UnDisabled"
    Call DisabledChannel(1)
Case "Del"
    Call DelChannel
Case "Order"
    Call order
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "UpdateData"
    Call UpdateData
Case "UpdateChannelFiles"
    Call UpdateChannelFiles
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "Ƶ���������ʧ�ܣ�ʧ��ԭ��" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsChannelList, sqlChannelList
    sqlChannelList = "select * from PE_Channel Where 1=1"
    If Not (FoundInArr(AllModules, "Supply", ",")) Then
        sqlChannelList = sqlChannelList & " And ModuleType<>6"
    End If
    If Not (FoundInArr(AllModules, "Job", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>8"
    End If
    If Not (FoundInArr(AllModules, "House", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>7"
    End If
    sqlChannelList = sqlChannelList & " order by OrderID "
    Set rsChannelList = Conn.Execute(sqlChannelList)
    
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "    <td align='center'><strong>Ƶ������</strong></td>"
    Response.Write "    <td width='54' align='center'><strong>�򿪷�ʽ</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>Ƶ������</strong></td>"
    Response.Write "    <td width='120' align='center'><strong>Ƶ��Ŀ¼/���ӵ�ַ</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>��Ŀ����</strong></td>"
    Response.Write "    <td width='54' align='center'><strong>����ģ��</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>����HTML��ʽ</strong></td>"
    Response.Write "    <td width='54' align='center'><strong>Ƶ��״̬</strong></td>"
    Response.Write "    <td width='110' align='center'><strong>����</strong></td>"
    Response.Write "    <td width='65' align='center'><strong>Ƶ������</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    Do While Not rsChannelList.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td align='center'>" & rsChannelList("ChannelID") & "</td>"
        Response.Write "    <td align='center'><a href='Admin_Channel.asp?Action=Modify&iChannelID=" & rsChannelList("ChannelID") & "' title='" & rsChannelList("ReadMe") & "'>" & rsChannelList("ChannelName") & "</a></td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("OpenType") = 0 Then
            Response.Write "<font color=green>ԭ����</font>"
        Else
            Response.Write "�´���"
        End If
        Response.Write "</td>"
        Response.Write "<td width='60' align='center'>"
        Select Case rsChannelList("ChannelType")
        Case 0
            Response.Write "<font color=blue>ϵͳƵ��</font>"
        Case 1
            Response.Write "<font color=green>�ڲ�Ƶ��</font>"
        Case 2
            Response.Write "<font color=red>�ⲿƵ��</font>"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='120' style='word-wrap:break-word'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write "Ŀ¼��" & rsChannelList("ChannelDir")
        Else
            Response.Write "<font color=red>���ӣ�" & rsChannelList("LinkUrl") & "</font>"
        End If
        Response.Write "</td>"
        Response.Write "    <td width='60' align='center'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write rsChannelList("ChannelShortName")
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write GetModuleTypeName(rsChannelList("ModuleType"))
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='60' align='center'>"
        Select Case rsChannelList("UseCreateHTML")
        Case 0
            Response.Write "������"
        Case 1
            Response.Write "ȫ������"
        Case 2
            Response.Write "��������1"
        Case 3
            Response.Write "��������2"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("Disabled") = True Then
            Response.Write "<font color=red>�ѽ���</font>"
        Else
            Response.Write "����"
        End If
        Response.Write "</td>"
        Response.Write "<td width='110' align='center'>"
        Response.Write "<a href='Admin_Channel.asp?Action=Modify&iChannelID=" & rsChannelList("ChannelID") & "'>�޸�</a>&nbsp;&nbsp;"
        If rsChannelList("Disabled") = True Then
            Response.Write "<a href='Admin_Channel.asp?Action=UnDisabled&iChannelID=" & rsChannelList("ChannelID") & "'>����</a>&nbsp;&nbsp;"
        Else
            Response.Write "<a href='Admin_Channel.asp?Action=Disabled&iChannelID=" & rsChannelList("ChannelID") & "'>����</a>&nbsp;&nbsp;"
        End If
        If rsChannelList("ChannelType") > 0 Then
            Response.Write "<a href='Admin_Channel.asp?Action=Del&iChannelID=" & rsChannelList("ChannelID") & "' onClick=""return confirm('ȷ��Ҫɾ����Ƶ����');"">ɾ��</a>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='65' align='center'>"
        If rsChannelList("ChannelType") < 2 And rsChannelList("ModuleType") <> 4 And rsChannelList("ModuleType") <> 8 Then
            Response.Write "<a href='Admin_Channel.asp?Action=UpdateData&iChannelID=" & rsChannelList("ChannelID") & "'>����</a>&nbsp;"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        If rsChannelList("ChannelType") = 1 And rsChannelList("ModuleType") <> 4 Then
            Response.Write "<a href='Admin_Channel.asp?Action=UpdateChannelFiles&iChannelID=" & rsChannelList("ChannelID") & "'>�ļ�</a>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "</tr>"
        rsChannelList.MoveNext
    Loop
    Response.Write "</table>"
    rsChannelList.Close
    Set rsChannelList = Nothing
    Response.Write "<form name='form1' method='post' action='Admin_Channel.asp'><div align='center'>"
    Response.Write "<input type='hidden' name='Action' value='UpdateData'>"
    Response.Write "<input type='submit' name='submit' value='��������Ƶ��������' onclick=""document.form1.Action.value='UpdateData'""> "
    Response.Write "<input type='submit' name='submit' value='��������Ƶ�����ļ�' onclick=""document.form1.Action.value='UpdateChannelFiles'"">"
    Response.Write "</div></form>"
End Sub

Sub order()
    Dim rsChannelList, sqlChannelList, iCount, i, j
    'sqlChannelList = "select * from PE_Channel order by OrderID"
    sqlChannelList = "select * from PE_Channel Where 1=1"
    If Not (FoundInArr(AllModules, "Supply", ",")) Then
        sqlChannelList = sqlChannelList & " And ModuleType<>6"
    End If
    If Not (FoundInArr(AllModules, "Job", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>8"
    End If
    If Not (FoundInArr(AllModules, "House", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>7"
    End If
    sqlChannelList = sqlChannelList & " order by OrderID "
    Set rsChannelList = Server.CreateObject("Adodb.RecordSet")
    rsChannelList.Open sqlChannelList, Conn, 1, 1
    iCount = rsChannelList.RecordCount
    j = 1
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write " <td width='32' align='center'><strong>���</strong></td>"
    Response.Write " <td height='22' align='center'><strong> Ƶ������</strong></td>"
    Response.Write " <td width='54' align='center'><strong>�򿪷�ʽ</strong></td>"
    Response.Write " <td width='80' align='center'><strong>Ƶ������</strong></td>"
    Response.Write " <td width='120' align='center'><strong>Ƶ��Ŀ¼/</strong><strong>���ӵ�ַ</strong></td>"
    Response.Write " <td width='80' align='center'><strong>����ģ��</strong></td>"
    Response.Write " <td width='240' colspan='2' align='center'><strong>����</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    Do While Not rsChannelList.EOF
        Response.Write "<tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "<td width='32' align='center'>" & rsChannelList("OrderID") & "</td>"
        Response.Write "    <td align='center'><a href='Admin_Channel.asp?Action=Modify&iChannelID=" & rsChannelList("ChannelID") & "' title='" & nohtml(rsChannelList("ReadMe")) & "'>" & rsChannelList("ChannelName") & "</a></td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("OpenType") = 0 Then
            Response.Write "<font color=green>ԭ����</font>"
        Else
            Response.Write "�´���"
        End If
        Response.Write "</td>"
        Response.Write "<td width='80' align='center'>"
        Select Case rsChannelList("ChannelType")
        Case 0
            Response.Write "<font color=blue>ϵͳƵ��</font>"
        Case 1
            Response.Write "<font color=green>�ڲ�Ƶ��</font>"
        Case 2
            Response.Write "<font color=red>�ⲿƵ��</font>"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='120'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write "Ŀ¼��" & rsChannelList("ChannelDir")
        Else
            Response.Write "<font color=red>���ӣ�" & rsChannelList("LinkUrl") & "</font>"
        End If
        Response.Write "</td>"
        Response.Write "<td width='80' align='center'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write GetModuleTypeName(rsChannelList("ModuleType"))
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<form action='Admin_Channel.asp?Action=UpOrder' method='post'>"
        Response.Write "  <td width='120' align='center'>"
        If j > 1 Then
            Response.Write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select>"
            Response.Write "<input type=hidden name=iChannelID value=" & rsChannelList("ChannelID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsChannelList("OrderID") & ">&nbsp;<input type=submit name=Submit value=�޸�>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form>"
        Response.Write "<form action='Admin_Channel.asp?Action=DownOrder' method='post'>"
        Response.Write "  <td width='120' align='center'>"
        If iCount > j Then
            Response.Write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select>"
            Response.Write "<input type=hidden name=iChannelID value=" & rsChannelList("ChannelID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsChannelList("OrderID") & ">&nbsp;<input type=submit name=Submit value=�޸�>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form></tr>"
        j = j + 1
        rsChannelList.MoveNext
    Loop
    Response.Write "</table>"
    rsChannelList.Close
    Set rsChannelList = Nothing
End Sub

Sub AddChannel()
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Channel.asp'>Ƶ������</a>&nbsp;&gt;&gt;&nbsp;�����Ƶ��</td></tr></table>"
    Response.Write "<form method='post' action='Admin_Channel.asp' name='myform' onSubmit='return CheckForm();'>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>Ƶ������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>ǰ̨��ʽ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>�ϴ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>����ѡ��</td>" & vbCrLf
    If IsCustom_Content = True Then
        Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>��������</td>" & vbCrLf
    End If
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong> Ƶ�����ƣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelName' type='text' id='ChannelName' size='49' maxlength='30'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��ͼƬ��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelPicUrl' type='text' id='ChannelPicUrl' size='49' maxlength='200'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��˵����</strong><br>�������Ƶ��������ʱ����ʾ�趨��˵�����֣���֧��HTML��</td>" & vbCrLf
    Response.Write "      <td valign='middle'><textarea name='ReadMe' cols='40' rows='3' id='ReadMe'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ�����ͣ�</strong><br><font color=red>������ѡ��Ƶ��һ����Ӻ�Ͳ����ٸ���Ƶ�����͡�</font></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function HideTabTitle(displayValue,tempType){" & vbCrLf
    Response.Write "  for (var i = 1; i < TabTitle.length; i++) {" & vbCrLf
    Response.Write "    if(tempType==0&&i==2) {" & vbCrLf
    Response.Write "        TabTitle[i].style.display='none';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    else{" & vbCrLf
    Response.Write "        TabTitle[i].style.display=displayValue;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<input name='ChannelType' type='radio' value='2'  onclick=""HideTabTitle('none');ChannelSetting.style.display='none'"" ><font color=blue><b>�ⲿƵ��</b></font>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�ⲿƵ��ָ���ӵ���ϵͳ����ĵ�ַ�С�����Ƶ��׼�����ӵ���վ�е�����ϵͳʱ����ʹ�����ַ�ʽ��<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�ⲿƵ�������ӵ�ַ��<input name='LinkUrl' type='text' id='LinkUrl' value='' size='40' maxlength='200'>"
    Response.Write "   <br><br>" & vbCrLf
    Response.Write "   <input name='ChannelType' type='radio' value='1' checked"
    If ObjInstalled_FSO = False Then Response.Write " disabled "
    Response.Write " onclick=""HideTabTitle('',1);ChannelSetting.style.display=''"">"
    Response.Write "<font color=blue><b>ϵͳ�ڲ�Ƶ��</b></font>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ�ڲ�Ƶ��ָ�����ڱ�ϵͳ���й���ģ�飨���š����¡�ͼƬ�ȣ�����������µ�Ƶ������Ƶ���߱�����ʹ�ù���ģ����ȫ��ͬ�Ĺ��ܡ����磬���һ����Ϊ������ѧԺ������Ƶ������Ƶ��ʹ�á����¡�ģ��Ĺ��ܣ�������ӵġ�����ѧԺ��Ƶ������ԭ����Ƶ�������й��ܡ�<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�˹�����Ҫ������֧��FSO�ſ��á�<br>" & vbCrLf
    Response.Write "      <table id='ChannelSetting' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF' style='display:'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>Ƶ��ʹ�õĹ���ģ�飺</strong></td>"
    Response.Write "          <td><select name='ModuleType' id='ModuleType'>"
    Response.Write "          <option value='1' selected>����</option>"
    Response.Write "          <option value='2'>����</option>"
    Response.Write "          <option value='3'>ͼƬ</option>"
    Response.Write "          </select>&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>������ѡ��Ƶ��һ����Ӻ�Ͳ����޸Ĵ��</font></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>Ƶ��Ŀ¼��</strong>��Ƶ��Ӣ������<br>"
    Response.Write "          <font color='#FF0000'>ֻ����Ӣ�ģ����ܴ��ո��\������/���ȷ��š�</font><br>"
    Response.Write "          <font color='#0000FF'>������</font>News��Article��Soft</td>"
    Response.Write "          <td><input name='ChannelDir' type='text' id='ChannelDir' size='20' maxlength='50'>  <font color='#FF0000'>*&nbsp;&nbsp;&nbsp;&nbsp;������¼�룬Ƶ��һ����Ӻ�Ͳ����޸Ĵ��</font></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong><font color=red>Ƶ�����ӵ�ַ������������</font></strong><br>�����Ҫ��ǰ̨����Ƶ����Ϊ��վ��һ��<font color='red'>������վ��</font>�����ʣ���������������ַ���磺http://news.powereasy.net���������뱣��Ϊ�ա�</td>"
    Response.Write "          <td><input name='ChannelUrl' type='input' value='' size='30' maxlength='100'"
    If SiteUrlType = 0 Then Response.Write " disabled"
    Response.Write "> <font color='red'>* ���ܴ�Ŀ¼</font><br>���Ҫ���ô˹��ܣ������ڡ���վѡ��н������ӵ�ַ��ʽ����Ϊ������·����</td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>��Ŀ���ƣ�</strong><br>���磺Ƶ������Ϊ������ѧԺ��������Ŀ����Ϊ�����¡��򡰽̡̳�</td>"
    Response.Write "          <td><input name='ChannelShortName' type='text' id='ChannelShortName' size='20' maxlength='30'> <font color='#FF0000'>*</font></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>��Ŀ��λ��</strong><br>���磺��ƪ������������������</td>"
    Response.Write "          <td><input name='ChannelItemUnit' type='text' id='ChannelItemUnit' size='10' maxlength='30'></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>�򿪷�ʽ��</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='OpenType' value='0'>��ԭ���ڴ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='OpenType' type='radio' value='1' checked>���´��ڴ�" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>���ñ�Ƶ����</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='Disabled' value='True'>�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='Disabled' type='radio' value='False' checked>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��Ȩ�ޣ�</strong><br><font color='red'>Ƶ��Ȩ��Ϊ�̳й�ϵ����Ƶ����Ϊ����֤Ƶ����ʱ�����µ���Ŀ��Ϊ��������Ŀ��Ҳ��Ч���෴�����Ƶ����Ϊ������Ƶ���������µ���Ŀ������������Ȩ�ޡ�</font></td>"
    Response.Write "      <td>"
    Response.Write "        <table>"
    Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' value='0' checked>����Ƶ��</td><td>�κ��ˣ������οͣ����������Ƶ���µ���Ϣ����������Ŀ��������ָ���������ĿȨ�ޡ�</td></tr>"
    Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' value='1'>��֤Ƶ��</td><td>�οͲ����������������ָ����������Ļ�Ա�顣���Ƶ������Ϊ��֤Ƶ�������Ƶ���ġ�����HTML��ѡ��ֻ����Ϊ��������HTML����</td></tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>���������Ƶ���Ļ�Ա�飺</strong><br>���Ƶ��Ȩ������Ϊ����֤Ƶ���������ڴ��������������Ƶ���Ļ�Ա��</td>"
    Response.Write "      <td>" & GetUserGroup("", "") & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ������˼���</strong><br>�趨Ϊ��Ҫ���ʱ�����ĳ���Ա�С�������Ϣ������ˡ�����Ȩ����˻�Ա�鲻�ܴ��ޡ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='CheckLevel' type='radio' value='0'>�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>�ڱ�Ƶ��������Ϣ����Ҫ����Ա���</font><br>"
    Response.Write "        <input name='CheckLevel' type='radio' value='1' checked>һ�����&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>��Ҫ��Ŀ���Ա������ˣ�ע���˼���Ϊ��С������ͬ��</font><br>"
    Response.Write "        <input name='CheckLevel' type='radio' value='2'>�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>������Ҫ��Ŀ���Ա��ˣ�������Ҫ��Ŀ�ܱ�������</font><br>"
    Response.Write "        <input name='CheckLevel' type='radio' value='3'>�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>������Ҫ��Ŀ���Ա��ˣ�������Ҫ��Ŀ�ܱ���ˣ�������ҪƵ������Ա���</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf	
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>�������κ���Ŀ������Ȩ�ޣ�</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input name='EnableComment' type='checkbox' value='True' checked>����������<br>"
    Response.Write "        <input name='CheckComment' type='checkbox' value='True' checked>������Ҫ���"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>�˸�ʱվ�ڶ���/Email֪ͨ���ݣ�</strong><br>��֧��HTML����</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfReject' cols='60' rows='4'>�ǳ���Ǹ�ĸ�����������{$ChannelShortName}��{$Title}����Ϊ���¼���ԭ��δ��¼�ã�" & vbCrLf & "1��" & vbCrLf & "2��" & vbCrLf & "3��" & vbCrLf & vbCrLf & "�ڴ��������ٴ�Ͷ�壡</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>���������ʱվ�ڶ���/Email֪ͨ���ݣ�</strong><br>��֧��HTML����</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfPassed' cols='60' rows='4'>��ϲ��������{$ChannelShortName}��{$Title}���Ѿ���¼�ã�" & vbCrLf & "�ǳ���л����Ͷ�壡</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��META�ؼ��ʣ�</strong><br>��������������õĹؼ���<br>����ؼ�������,�ŷָ�</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��META��ҳ������</strong><br>��������������õ���ҳ����<br>�����������,�ŷָ�</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>���/�޸���Ϣʱ�Ľ������ã�</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input name='arrEnabledTabs' type='checkbox' value='Charge' checked>��ʾ���շ�ѡ���ǩ<br>"
    Response.Write "        <input name='arrEnabledTabs' type='checkbox' value='Vote' checked>��ʾ���������á���ǩ<br>"
    Response.Write "        <input name='arrEnabledTabs' type='checkbox' value='SoftParameter' checked>��ʾ�������������ǩ����������ģ����Ч��<br>"
    Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Recieve' checked>��ʾ��ǩ�����á���ǩ����������ģ����Ч��<br>"
    Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Copyfee' checked>��ʾ��������á���ǩ����������ģ����Ч��<br>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ĭ�ϸ�ѱ�׼��</strong>(��λ��Ԫ/ǧ��)</td>"
    Response.Write "      <td><input name='MoneyPerKw' type='text'id='MoneyPerKw' size='10' maxlength='10'> <font color=red>Ԫ/ǧ��</font></Td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ���Ƶ������ʾƵ�����ƣ�</strong></td>"
    Response.Write "      <td><input name='ShowName' type='radio' value='True' checked>�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��ڵ�������ʾƵ�����ƣ�</strong></td>"
    Response.Write "      <td><input name='ShowNameOnPath' type='radio' value='True' checked>�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��ڱ�Ƶ����ʾ��״�����˵���</strong></td>"
    Response.Write "      <td><input name='ShowClassTreeGuide' type='radio' value='True' checked>�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowClassTreeGuide' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>ǰ̨�Ƿ���ʾ������ʡ�Ժţ�</strong><br>��ģ����ָ�����ⳤ��С�ڱ���ʵ�ʳ���ʱ�����Ծ����Ƿ��ڱ��������ʾʡ�Ժ�</td>"
    Response.Write "      <td><input name='ShowSuspensionPoints' type='radio' value='True' checked>�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowSuspensionPoints' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ���ȵ�ĵ������Сֵ��</strong></td>"
    Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='500' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�������ڸ��µ���ϢΪ����Ϣ��</strong></td>"
    Response.Write "      <td><input name='DaysOfNew' type='text' id='DaysOfNew' value='7' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>����������ÿ����ʾ����Ŀ����</strong></td>"
    Response.Write "      <td><input name='MaxPerLine' type='text' id='MaxPerLine' value='10' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>��Ϣ�б����������Ƶĳ��ȣ�</strong></td>"
    Response.Write "      <td><input name='AuthorInfoLen' type='text' id='AuthorInfoLen' value='8' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ����ҳר���б��������</strong></td>"
    Response.Write "      <td><input name='JS_SpecialNum' type='text' id='JS_SpecialNum' value='10' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ����ҳ��ÿҳ��Ϣ����</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_Index' type='text' id='MaxPerPage_Index' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�������ҳ��ÿҳ��Ϣ����</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>������Ϣҳ��ÿҳ��Ϣ����</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_New' type='text' id='MaxPerPage_New' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>������Ϣҳ��ÿҳ��Ϣ����</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_Hot' type='text' id='MaxPerPage_Hot' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�Ƽ���Ϣҳ��ÿҳ��Ϣ����</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_Elite' type='text' id='MaxPerPage_Elite' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>ר���б�ҳ��ÿҳ��Ϣ����</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_SpecialList' type='text' id='MaxPerPage_SpecialList' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ����Ĭ�Ϸ��</strong></td>"
    Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(0) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>������Ŀ�˵�����ʾ��ʽ��</strong><br>���Ĵ˲�������Ҫˢ����ĿJS����Ч��</td>"
    Response.Write "      <td><select name='TopMenuType' id='TopMenuType'>" & GetMenuType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�ײ���Ŀ��������ʾ��ʽ��</strong><br>���Ĵ˲�������Ҫˢ����ĿJS����Ч��</td>"
    Response.Write "      <td><select name='ClassGuideType' id='ClassGuideType'>" & GetGuideType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf



    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ������ڱ�Ƶ���ϴ��ļ���</strong></td>"
    Response.Write "      <td><input name='EnableUploadFile' type='radio' value='True' checked>�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='EnableUploadFile' type='radio' value='False'>��</td>"
    Response.Write "    </tr>" & vbCrLf
    Dim UploadDir
    Randomize
    UploadDir = "UploadFiles_" & CInt(Rnd * 8999 + 1000)
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�ϴ��ļ��ı���Ŀ¼��</strong><br><font color='red'>����Զ��ڻ򲻶��ڵĸ����ϴ�Ŀ¼���Է�������վ����</font></td>"
    Response.Write "      <td><input name='UploadDir' type='text' id='UploadDir' value='" & UploadDir & "' size='20' maxlength='20'>&nbsp;&nbsp;<font color='red'>ֻ����Ӣ�ĺ����֣����ܴ��ո��\������/���ȷ��š�</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�����ϴ�������ļ���С��</strong></td>"
    Response.Write "      <td><input name='MaxFileSize' type='text' id='MaxFileSize' value='1024' size='10' maxlength='10'> KB&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>��ʾ��1 KB = 1024 Byte��1 MB = 1024 KB</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>�����ϴ����ļ����ͣ�</strong><br>�����ļ�����֮���ԡ�|���ָ�</td>"
    Response.Write "      <td><table>"
    Response.Write "          <tr><td>ͼƬ���ͣ�</td><td><input name='UpFileType' type='text' id='UpFileType' value='gif|jpg|jpeg|jpe|bmp|png' size='50' maxlength='200'></td></tr>"
    Response.Write "          <tr><td>Flash�ļ���</td><td><input name='UpFileType' type='text' id='UpFileType' value='swf' size='50' maxlength='50'></td></tr>"
    Response.Write "          <tr><td>Windowsý�壺</td><td><input name='UpFileType' type='text' id='UpFileType' value='mid|mp3|wmv|asf|avi|mpg' size='50' maxlength='200'></td></tr>"
    Response.Write "          <tr><td>Realý�壺</td><td><input name='UpFileType' type='text' id='UpFileType' value='ram|rm|ra' size='20' maxlength='200'></td></tr>"
    Response.Write "          <tr><td>�����ļ���</td><td><input name='UpFileType' type='text' id='UpFileType' value='rar|exe|doc|zip' size='50' maxlength='200'></td></tr>"
    Response.Write "      </table></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf



    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>����HTML��ʽ��</font></strong><br>������֧��FSO�������á�����HTML������<br>�����ѡ���Ժ���ÿһ�θ������ɷ�ʽǰ���������ɾ��������ǰ���ɵ��ļ���Ȼ���ڱ���Ƶ�����������������������ļ���</td>"
    Response.Write "      <td>" & GetUseCreateHTML(0, 0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>��Ŀ��ר���б����ҳ����</font></strong><br>������ݺ��Զ����µ���Ŀ��ר���б�ҳ����</td>"
    Response.Write "      <td><input name='UpdatePages' type='text' id='UpdatePages' value='3' size='5' maxlength='5'> ҳ <font color='#FF0000'>*</font>&nbsp;&nbsp;<font color='blue'>�磺����ҳ����Ϊ3����ÿ���Զ�����ǰ��ҳ����4ҳ�Ժ�ķ�ҳΪ�̶����ɵ�ҳ�棬����������������һҳ����������һ���̶�ҳ�棬���ܼ�¼������ÿҳ��¼����������ʱ������ҳ����3��4ҳ�����в��ּ�¼�ظ���</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr><td colspan='2'><font color='red'><b>���²�������������HTML��ʽ����Ϊ������ʱ����Ч��<br>�����ѡ���Ժ���ÿһ�θ������²���ǰ���������ɾ��������ǰ���ɵ��ļ���Ȼ���ڱ���������ú����������������ļ���</b></font></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>�Զ�����HTMLʱ�����ɷ�ʽ��</font></strong><br>���/�޸���Ϣʱ��ϵͳ�����Զ������й�ҳ���ļ�����������ѡ���Զ�����ʱ�ķ�ʽ��</td>"
    Response.Write "      <td>" & GetAutoCreateType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>��Ŀ�б��ļ��Ĵ��λ�ã�</font></strong></td>"
    Response.Write "      <td>" & GetListFileType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>Ŀ¼�ṹ��ʽ��</font></strong></td>"
    Response.Write "      <td>" & GetStructureType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>����ҳ�ļ���������ʽ��</font></strong></td>"
    Response.Write "      <td>" & GetFileNameType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>Ƶ����ҳ����չ����</font></strong></td>"
    Response.Write "      <td>" & arrFileExt_Index(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>��Ŀҳ��ר��ҳ����չ����</font></strong></td>"
    Response.Write "      <td>" & arrFileExt_List(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>����ҳ����չ����</font></strong></td>"
    Response.Write "      <td>" & arrFileExt_Item(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    If IsCustom_Content = True Then
        Call EditCustom_Content("Add", "", "Channel")
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "      <input  type='submit' name='Submit' value=' �� �� '> &nbsp; <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Channel.asp'"" style='cursor:hand;'></td>"
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "</form>"

    Call ShowChekcFormJS
End Sub

Sub ShowChekcFormJS()
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(document.myform.ChannelName.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('������Ƶ�����ƣ�');" & vbCrLf
    Response.Write "    document.myform.ChannelName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.ChannelType[1].checked==true){" & vbCrLf
    Response.Write "    if(document.myform.ChannelDir.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('������Ƶ��Ŀ¼��');" & vbCrLf
    Response.Write "      document.myform.ChannelDir.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(document.myform.ChannelShortName.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('��������Ŀ���ƣ�');" & vbCrLf
    Response.Write "      document.myform.ChannelShortName.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    if(document.myform.LinkUrl.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('������Ƶ�������ӵ�ַ��');" & vbCrLf
    Response.Write "      document.myform.LinkUrl.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

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
End Sub

Sub Modify()
    Dim iChannelID, rsChannel
    iChannelID = Trim(Request("iChannelID"))
    If iChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸�Ƶ��ID</li>"
        Exit Sub
    Else
        iChannelID = PE_CLng(iChannelID)
    End If
    Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & iChannelID)
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����Ƶ����</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Sub
    End If
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
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

    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Channel.asp'>Ƶ������</a>&nbsp;&gt;&gt;&nbsp;�޸�Ƶ�����ã�<font color='red'>" & rsChannel("ChannelName") & "</font></td></tr></table>"
    Response.Write "<form method='post' action='Admin_Channel.asp' name='myform' onSubmit='return CheckForm();'>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    If rsChannel("ChannelType") <> 2 Then
        If rsChannel("ModuleType") <> 4 Then
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>Ƶ������</td>" & vbCrLf
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>ǰ̨��ʽ</td>" & vbCrLf
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>�ϴ�ѡ��</td>" & vbCrLf			
        End If
        If rsChannel("ModuleType") = 4 Then		
             Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>�ϴ�ѡ��</td>" & vbCrLf
        End If					
        '�����Σ�����Ƶ������
        If rsChannel("ModuleType") <> 6 And rsChannel("ModuleType") <> 4 And rsChannel("ModuleType") <> 7 And rsChannel("ModuleType") <> 8 Then
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>����ѡ��</td>" & vbCrLf
            If IsCustom_Content = True Then
                Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>��������</td>" & vbCrLf
            End If
        End If
        '2005-12-23
    End If
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong> Ƶ�����ƣ�</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelName' type='text' id='ChannelName' size='49' maxlength='30' value='" & rsChannel("ChannelName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��ͼƬ��</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelPicUrl' type='text' id='ChannelPicUrl' size='49' maxlength='200' value='" & rsChannel("ChannelPicUrl") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��˵����</strong><br>�������Ƶ��������ʱ����ʾ�趨��˵�����֣���֧��HTML��</td>" & vbCrLf
    Response.Write "      <td valign='middle'><textarea name='ReadMe' cols='40' rows='3' id='ReadMe'>" & rsChannel("ReadMe") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ�����ͣ�</strong><br><font color=red>������ѡ��Ƶ��һ����Ӻ�Ͳ����ٸ���Ƶ�����͡�</font></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelType' type='radio' value='2'"
    If rsChannel("ChannelType") > 1 Then
        Response.Write " checked "
    Else
        Response.Write " disabled"
    End If
    Response.Write " onClick=""ChannelSetting.style.display='none'""><font color=blue><b>�ⲿƵ��</b></font></legend>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�ⲿƵ��ָ���ӵ���ϵͳ����ĵ�ַ�С�����Ƶ��׼�����ӵ���վ�е�����ϵͳʱ����ʹ�����ַ�ʽ��<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�ⲿƵ�������ӵ�ַ��"
    Response.Write "      <input name='LinkUrl' type='text' id='LinkUrl' value='" & rsChannel("LinkUrl") & "' size='40' maxlength='200'"
    If rsChannel("ChannelType") <= 1 Then
        Response.Write " disabled"
    End If
    Response.Write "><br><br>"
    Response.Write "<input name='ChannelType' type='radio' value='1'"
    If rsChannel("ChannelType") <= 1 Then
        Response.Write " checked"
    Else
        Response.Write " disabled"
    End If
    Response.Write "><font color=blue><b>ϵͳ�ڲ�Ƶ��</b></font></legend>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ�ڲ�Ƶ��ָ�����ڱ�ϵͳ���й���ģ�飨���š����¡�ͼƬ�ȣ�����������µ�Ƶ������Ƶ���߱�����ʹ�ù���ģ����ȫ��ͬ�Ĺ��ܡ����磬���һ����Ϊ������ѧԺ������Ƶ������Ƶ��ʹ�á����¡�ģ��Ĺ��ܣ�������ӵġ�����ѧԺ��Ƶ������ԭ����Ƶ�������й��ܡ�<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�˹�����Ҫ������֧��FSO�ſ��á�<br>"
    Response.Write "     <table id='ChannelSetting' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'"
    If rsChannel("ChannelType") > 1 Then Response.Write " style='display:none'"
    Response.Write ">"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td colspan='2'><strong>�ڲ�Ƶ����������</strong></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��ʹ�õĹ���ģ�飺</strong></td>"
    Response.Write "      <td><select name='ModuleType' id='ModuleType' disabled>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 1) & ">����</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 2) & ">����</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 3) & ">ͼƬ</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 4) & ">���԰�</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 5) & ">�̳�</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 6) & ">����</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 7) & ">����</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 8) & ">�˲�</option>"
    Response.Write "      </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��Ŀ¼��</strong>��Ƶ��Ӣ������<br>"
    Response.Write "        <font color='#FF0000'>ֻ����Ӣ�ģ����ܴ��ո��\������/���ȷ��š�</font><br><font color='#0000FF'>������</font>News��Article��Soft</td>"
    Response.Write "      <td><input name='ChannelDir' type='text' id='ChannelDir' value='" & rsChannel("ChannelDir") & "' size='20' maxlength='50' disabled>"
    If rsChannel("ChannelType") <= 1 Then Response.Write "<input name='ChannelDir' type='hidden' id='ChannelDir' value='" & rsChannel("ChannelDir") & "'>"
    Response.Write "<font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    If rsChannel("ModuleType") <> 4 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>Ƶ�����ӵ�ַ������������</font></strong><br>�����Ҫ��ǰ̨����Ƶ����Ϊ��վ��һ��<font color='red'>������վ��</font>�����ʣ���������������ַ���磺http://news.powereasy.net���������뱣��Ϊ�ա�</td>"
        Response.Write "      <td><input name='ChannelUrl' type='input' size='30' maxlength='100'"
        If SiteUrlType = 0 Then
            Response.Write " disabled"
        Else
            Response.Write " value='" & rsChannel("LinkUrl") & "'"
        End If
        Response.Write "> <font color='red'>* ���ܴ�Ŀ¼</font><br>���Ҫ���ô˹��ܣ������ڡ���վѡ��н������ӵ�ַ��ʽ����Ϊ������·����</td>"
        Response.Write "    </tr>" & vbCrLf
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>��Ŀ���ƣ�</strong><br>���磺Ƶ������Ϊ������ѧԺ��������Ŀ����Ϊ�����¡��򡰽̡̳�</td>"
    Response.Write "      <td><input name='ChannelShortName' type='text' id='ChannelShortName' size='20' maxlength='30' value='" & rsChannel("ChannelShortName") & "'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>��Ŀ��λ��</strong><br>���磺��ƪ������������������</td>"
    Response.Write "      <td><input name='ChannelItemUnit' type='text' id='ChannelItemUnit' size='10' maxlength='30' value='" & rsChannel("ChannelItemUnit") & "'></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "       </table>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>�򿪷�ʽ��</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='OpenType' " & RadioValue(rsChannel("OpenType"), 0) & ">��ԭ���ڴ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='OpenType' type='radio' " & RadioValue(rsChannel("OpenType"), 1) & ">���´��ڴ�</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='200' class='tdbg5'><strong>���ñ�Ƶ����</strong></td>"
    Response.Write "      <td><input name='Disabled' type='radio' " & RadioValue(rsChannel("Disabled"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='Disabled' type='radio' " & RadioValue(rsChannel("Disabled"), False) & ">��</td>"
    Response.Write "    </tr>" & vbCrLf
    If rsChannel("ModuleType") = 4 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ���Ƶ������ʾƵ�����ƣ�</strong></td>"
        Response.Write "      <td><input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), False) & ">��</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��ڵ�������ʾƵ�����ƣ�</strong></td>"
        Response.Write "      <td><input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), False) & ">��</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��������Ե���˹��ܣ�</strong></td>"
        Response.Write "      <td>"
        Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 1) & ">��&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 0) & ">��&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "        <br><font color='#0000FF'>�趨Ϊ��Ҫ���ʱ�����ĳ���Ա�С�������Ϣ������ˡ�����Ȩ����˻�Ա�鲻�ܴ��ޡ�</font>"
        Response.Write "      </td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ����Ĭ�Ϸ��</strong></td>"
        Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(rsChannel("DefaultSkinID")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
    End If
    Response.Write "  </tbody>" & vbCrLf

    '������ �����빩���˲��޹ص�Ƶ��
    If rsChannel("ModuleType") = 6 Or rsChannel("ModuleType") = 7 Or rsChannel("ModuleType") = 8 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        If rsChannel("ModuleType") = 6 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ���Ƽ����շ����ã�</strong><br>��Ϣ�ڴ�Ƶ����Ϊ�Ƽ���ÿ��Ҫ�۳��Ļ�Ա����.</td>"
            Response.Write "      <td>"
            Response.Write "        �Ƽ���Ϣ�۳�&nbsp;<INPUT TYPE='text' NAME='CommandChannelPoint' MaxLength='5' Size='5' Value='" & PE_CLng(rsChannel("CommandChannelPoint")) & "'>&nbsp;����/��"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        If rsChannel("ModuleType") <> 8 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ������˼���</strong><br>�趨Ϊ��Ҫ���ʱ�����ĳ���Ա�С�������Ϣ������ˡ�����Ȩ����˻�Ա�鲻�ܴ��ޡ�</td>"
            Response.Write "      <td>"
            Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 0) & ">�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>�ڱ�Ƶ��������Ϣ����Ҫ����Ա���</font><br>"
            Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 1) & ">һ�����&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>��Ҫ��Ŀ���Ա������ˣ�ע���˼���Ϊ��С����</font><br>"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��META�ؼ��ʣ�</strong><br>��������������õĹؼ���<br>����ؼ�������,�ŷָ�</td>" & vbCrLf
        Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsChannel("Meta_Keywords") & "</textarea></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��META��ҳ������</strong><br>��������������õ���ҳ����<br>�����������,�ŷָ�</td>" & vbCrLf
        Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsChannel("Meta_Description") & "</textarea></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        If rsChannel("ModuleType") <> 4 Then
            Response.Write "  <tbody id='Tabs' style='display:none'>"
            If rsChannel("ModuleType") <> 5 Then
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��Ȩ�ޣ�</strong><br><font color='red'>Ƶ��Ȩ��Ϊ�̳й�ϵ����Ƶ����Ϊ����֤Ƶ����ʱ�����µ���Ŀ��Ϊ��������Ŀ��Ҳ��Ч���෴�����Ƶ����Ϊ������Ƶ���������µ���Ŀ������������Ȩ�ޡ�</font></td>"
                Response.Write "      <td>"
                Response.Write "        <table>"
                Response.Write "     <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' " & RadioValue(rsChannel("ChannelPurview"), 0) & ">����Ƶ��</td><td>�κ��ˣ������οͣ����������Ƶ���µ���Ϣ����������Ŀ��������ָ���������ĿȨ�ޡ�</td></tr>"
                Response.Write "     <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' " & RadioValue(rsChannel("ChannelPurview"), 1) & ">��֤Ƶ��</td><td>�οͲ����������������ָ����������Ļ�Ա�顣���Ƶ������Ϊ��֤Ƶ�������Ƶ���ġ�����HTML��ѡ��ֻ����Ϊ��������HTML����</td></tr>"
                Response.Write "        </table>"
                Response.Write "      </td>"
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>���������Ƶ���Ļ�Ա�飺</strong><br>���Ƶ��Ȩ������Ϊ����֤Ƶ���������ڴ��������������Ƶ���Ļ�Ա��</td>"
                Response.Write "      <td>" & GetUserGroup(rsChannel("arrGroupID") & "", "") & "</td>"
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ������˼���</strong><br>�趨Ϊ��Ҫ���ʱ�����ĳ���Ա�С�������Ϣ������ˡ�����Ȩ����˻�Ա�鲻�ܴ��ޡ�</td>"
                Response.Write "      <td>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 0) & ">�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>�ڱ�Ƶ��������Ϣ����Ҫ����Ա���</font><br>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 1) & ">һ�����&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>��Ҫ��Ŀ���Ա������ˣ�ע���˼���Ϊ��С������ͬ��</font><br>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 2) & ">�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>������Ҫ��Ŀ���Ա��ˣ�������Ҫ��Ŀ�ܱ�������</font><br>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 3) & ">�������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>������Ҫ��Ŀ���Ա��ˣ�������Ҫ��Ŀ�ܱ���ˣ�������ҪƵ������Ա���</font>"
                Response.Write "      </td>"
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='300' class='tdbg5'><strong>�������κ���Ŀ������Ȩ�ޣ�</strong></td>"
                Response.Write "      <td>"
                Response.Write "        <input name='EnableComment' type='checkbox' value='True' "
                If PE_CBool(rsChannel("EnableComment")) = True Then Response.write "checked"
                Response.Write ">����������<br>"
                Response.Write "        <input name='CheckComment' type='checkbox' value='True' "
                If PE_CBool(rsChannel("CheckComment")) = True Then Response.write "checked"		
                Response.Write ">������Ҫ���"
                Response.Write "      </td>"
                Response.Write "    </tr>" & vbCrLf				
                Response.Write "    <tr class='tdbg'>" & vbCrLf
                Response.Write "      <td width='200' class='tdbg5'><strong>�˸�ʱվ�ڶ���/Email֪ͨ���ݣ�</strong><br>��֧��HTML����</td>" & vbCrLf
                Response.Write "      <td><textarea name='EmailOfReject' cols='60' rows='4'>" & rsChannel("EmailOfReject") & "</textarea></td>" & vbCrLf
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>" & vbCrLf
                Response.Write "      <td width='200' class='tdbg5'><strong>���������ʱվ�ڶ���/Email֪ͨ���ݣ�</strong><br>��֧��HTML����</td>" & vbCrLf
                Response.Write "      <td><textarea name='EmailOfPassed' cols='60' rows='4'>" & rsChannel("EmailOfPassed") & "</textarea></td>" & vbCrLf
                Response.Write "    </tr>" & vbCrLf
            End If
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��META�ؼ��ʣ�</strong><br>��������������õĹؼ���<br>����ؼ�������,�ŷָ�</td>" & vbCrLf
            Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsChannel("Meta_Keywords") & "</textarea></td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ��META��ҳ������</strong><br>��������������õ���ҳ����<br>�����������,�ŷָ�</td>" & vbCrLf
            Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsChannel("Meta_Description") & "</textarea></td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf

            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>���/�޸���Ϣʱ�Ľ������ã�</strong></td>"
            Response.Write "      <td>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Charge'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Charge", ",") = True Then Response.Write " checked"
            Response.Write ">��ʾ���շ�ѡ���ǩ<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Vote'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Vote", ",") = True Then Response.Write " checked"
            Response.Write ">��ʾ���������á���ǩ<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='SoftParameter'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "SoftParameter", ",") = True Then Response.Write " checked"
            Response.Write ">��ʾ�������������ǩ����������ģ����Ч��<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Recieve'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Recieve", ",") = True Then Response.Write " checked"
            Response.Write ">��ʾ��ǩ�����á���ǩ����������ģ����Ч��<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Copyfee'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Copyfee", ",") = True Then Response.Write " checked"
            Response.Write ">��ʾ��������á���ǩ����������ģ����Ч��<br>"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
            If rsChannel("ModuleType") = 1 Then
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>Ĭ�ϸ�ѱ�׼��</strong>(��λ��Ԫ/ǧ��)</td>"
                Response.Write "      <td><input name='MoneyPerKw' type='text'id='MoneyPerKw' size='10' maxlength='10' value='" & rsChannel("MoneyPerKw") & "'> <font color=red>Ԫ/ǧ��</font></Td>"
                Response.Write "    </tr>" & vbCrLf
            End If
            Response.Write "  </tbody>" & vbCrLf
        End If
    End If

    '2005-12-23 ������ϢƵ��
    '������ �����빩��ͷ������˲��޹ص�Ƶ��
    If rsChannel("ModuleType") = 6 Or rsChannel("ModuleType") = 7 Or rsChannel("ModuleType") = 8 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ������ҳģ��</strong></td>"
        Response.Write "      <td><select name='Template_Index' id='Template_Index'>" & GetTemplate_Option(iChannelID, 1, rsChannel("Template_Index")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ���Ƶ������ʾƵ�����ƣ�</strong></td>"
        Response.Write "      <td><input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), False) & ">��</td>"
        Response.Write "    </tr>" & vbCrLf
        If rsChannel("ModuleType") <> 8 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ���ȵ�ĵ������Сֵ��</strong></td>"
            Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='" & rsChannel("HitsOfHot") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        If rsChannel("ModuleType") = 6 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ����ҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Index' type='text' id='MaxPerPage_Index' value='" & rsChannel("MaxPerPage_Index") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�������ҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='" & rsChannel("MaxPerPage_SearchResult") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>������Ϣҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_New' type='text' id='MaxPerPage_New' value='" & rsChannel("MaxPerPage_New") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>������Ϣҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Hot' type='text' id='MaxPerPage_Hot' value='" & rsChannel("MaxPerPage_Hot") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�Ƽ���Ϣҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Elite' type='text' id='MaxPerPage_Elite' value='" & rsChannel("MaxPerPage_Elite") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>ר���б�ҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SpecialList' type='text' id='MaxPerPage_SpecialList' value='" & rsChannel("MaxPerPage_SpecialList") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��ڵ�������ʾƵ�����ƣ�</strong></td>"
        Response.Write "      <td><input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), False) & ">��</td>"
        Response.Write "    </tr>" & vbCrLf
        If rsChannel("ModuleType") = 6 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�������ڸ��µ���ϢΪ����Ϣ��</strong></td>"
            Response.Write "      <td><input name='DaysOfNew' type='text' id='DaysOfNew' value='" & rsChannel("DaysOfNew") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>����������ÿ����ʾ����Ŀ����</strong></td>"
            Response.Write "      <td><input name='MaxPerLine' type='text' id='MaxPerLine' value='" & rsChannel("MaxPerLine") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>������Ŀ�˵�����ʾ��ʽ��</strong><br>���Ĵ˲�������Ҫˢ����ĿJS����Ч��</td>"
            Response.Write "      <td><select name='TopMenuType' id='TopMenuType'>" & GetMenuType_Option(rsChannel("TopMenuType")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ����Ĭ�Ϸ��</strong></td>"
        Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(rsChannel("DefaultSkinID")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        If rsChannel("ModuleType") <> 4 Then
            Response.Write "  <tbody id='Tabs' style='display:none'>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ������ҳģ��</strong></td>"
            Response.Write "      <td><select name='Template_Index' id='Template_Index'>" & GetTemplate_Option(iChannelID, 1, rsChannel("Template_Index")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ���Ƶ������ʾƵ�����ƣ�</strong></td>"
            Response.Write "      <td><input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), False) & ">��</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��ڵ�������ʾƵ�����ƣ�</strong></td>"
            Response.Write "      <td><input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), False) & ">��</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ��ڱ�Ƶ����ʾ��״�����˵���</strong></td>"
            Response.Write "      <td><input name='ShowClassTreeGuide' type='radio' " & RadioValue(rsChannel("ShowClassTreeGuide"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowClassTreeGuide' type='radio' " & RadioValue(rsChannel("ShowClassTreeGuide"), False) & ">��</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>ǰ̨�Ƿ���ʾ������ʡ�Ժţ�</strong><br>��ģ����ָ�����ⳤ��С�ڱ���ʵ�ʳ���ʱ�����Ծ����Ƿ��ڱ��������ʾʡ�Ժ�</td>"
            Response.Write "      <td><input name='ShowSuspensionPoints' type='radio' " & RadioValue(rsChannel("ShowSuspensionPoints"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowSuspensionPoints' type='radio' " & RadioValue(rsChannel("ShowSuspensionPoints"), False) & ">��</td>"
            Response.Write "    </tr>" & vbCrLf
            If rsChannel("ModuleType") <> 5 Then
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ���ȵ�ĵ������Сֵ��</strong></td>"
                Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='" & rsChannel("HitsOfHot") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
                Response.Write "    </tr>" & vbCrLf
            End If
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�������ڸ��µ���ϢΪ����Ϣ��</strong></td>"
            Response.Write "      <td><input name='DaysOfNew' type='text' id='DaysOfNew' value='" & rsChannel("DaysOfNew") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>����������ÿ����ʾ����Ŀ����</strong></td>"
            Response.Write "      <td><input name='MaxPerLine' type='text' id='MaxPerLine' value='" & rsChannel("MaxPerLine") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>��Ϣ�б����������Ƶĳ��ȣ�</strong></td>"
            Response.Write "      <td><input name='AuthorInfoLen' type='text' id='AuthorInfoLen' value='" & rsChannel("AuthorInfoLen") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ����ҳר���б��������</strong></td>"
            Response.Write "      <td><input name='JS_SpecialNum' type='text' id='JS_SpecialNum' value='" & rsChannel("JS_SpecialNum") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>Ƶ����ҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Index' type='text' id='MaxPerPage_Index' value='" & rsChannel("MaxPerPage_Index") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�������ҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='" & rsChannel("MaxPerPage_SearchResult") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>������Ϣҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_New' type='text' id='MaxPerPage_New' value='" & rsChannel("MaxPerPage_New") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>������Ϣҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Hot' type='text' id='MaxPerPage_Hot' value='" & rsChannel("MaxPerPage_Hot") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�Ƽ���Ϣҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Elite' type='text' id='MaxPerPage_Elite' value='" & rsChannel("MaxPerPage_Elite") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>ר���б�ҳ��ÿҳ��Ϣ����</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SpecialList' type='text' id='MaxPerPage_SpecialList' value='" & rsChannel("MaxPerPage_SpecialList") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>��Ƶ����Ĭ�Ϸ��</strong></td>"
            Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(rsChannel("DefaultSkinID")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>������Ŀ�˵�����ʾ��ʽ��</strong><br>���Ĵ˲�������Ҫˢ����ĿJS����Ч��</td>"
            Response.Write "      <td><select name='TopMenuType' id='TopMenuType'>" & GetMenuType_Option(rsChannel("TopMenuType")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>�ײ���Ŀ��������ʾ��ʽ��</strong><br>���Ĵ˲�������Ҫˢ����ĿJS����Ч��</td>"
            Response.Write "      <td><select name='ClassGuideType' id='ClassGuideType'>" & GetGuideType_Option(rsChannel("ClassGuideType")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "  </tbody>" & vbCrLf
        End If
    End If

  '  If rsChannel("ModuleType") <> 4 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�Ƿ������ڱ�Ƶ���ϴ��ļ���</strong></td>"
        Response.Write "      <td><input name='EnableUploadFile' type='radio' " & RadioValue(rsChannel("EnableUploadFile"), True) & ">�� &nbsp;&nbsp;&nbsp;&nbsp;<input name='EnableUploadFile' type='radio' " & RadioValue(rsChannel("EnableUploadFile"), False) & ">��</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�ϴ��ļ��ı���Ŀ¼��</strong><br><font color='red'>����Զ��ڻ򲻶��ڵĸ����ϴ�Ŀ¼���Է�������վ����</font></td>"
        Response.Write "      <td><input name='UploadDir' type='text' id='UploadDir' value='" & rsChannel("UploadDir") & "' size='20' maxlength='20'>&nbsp;&nbsp;<font color='red'>ֻ����Ӣ�ĺ����֣����ܴ��ո��\������/���ȷ��š�</font></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�����ϴ�������ļ���С��</strong></td>"
        Response.Write "      <td><input name='MaxFileSize' type='text' id='MaxFileSize' value='" & rsChannel("MaxFileSize") & "' size='10' maxlength='10'> KB&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>��ʾ��1 KB = 1024 Byte��1 MB = 1024 KB</font></td>"
        Response.Write "    </tr>" & vbCrLf
        Dim arrFileType
        If rsChannel("UpFileType") & "" = "" Then
            arrFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")
        Else
            arrFileType = Split(rsChannel("UpFileType"), "$")
            If UBound(arrFileType) < 4 Then
                arrFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")
            End If
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>�����ϴ����ļ����ͣ�</strong><br>�����ļ�����֮���ԡ�|���ָ�</td>"
        Response.Write "      <td><table>"
        Response.Write "          <tr><td>ͼƬ���ͣ�</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(0)) & "' size='50' maxlength='200'></td></tr>"
        Response.Write "          <tr><td>Flash�ļ���</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(1)) & "' size='50' maxlength='50'></td></tr>"
        Response.Write "          <tr><td>Windowsý�壺</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(2)) & "' size='50' maxlength='200'></td></tr>"
        Response.Write "          <tr><td>Realý�壺</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(3)) & "' size='20' maxlength='200'></td></tr>"
        Response.Write "          <tr><td>�����ļ���</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(4)) & "' size='50' maxlength='200'></td></tr>"
        Response.Write "      </table></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
  '  End If

    '������ ���ι����в��õ���Ϣ
    If rsChannel("ModuleType") <> 6 Or rsChannel("ModuleType") <> 7 Or rsChannel("ModuleType") <> 4 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>����HTML��ʽ��</font></strong><br>������֧��FSO�������á�����HTML������<br>ÿһ�θ������ɷ�ʽ������Ҫ��ɾ��������ǰ���ļ������������������ļ���</td>"
        Response.Write "      <td>" & GetUseCreateHTML(rsChannel("UseCreateHTML"), rsChannel("ModuleType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>��Ŀ��ר���б����ҳ����</font></strong><br>������ݺ��Զ����µ���Ŀ��ר���б�ҳ����</td>"
        Response.Write "      <td><input name='UpdatePages' type='text' id='UpdatePages' value='" & rsChannel("UpdatePages") & "' size='5' maxlength='5'> ҳ <font color='#FF0000'>*</font>&nbsp;&nbsp;<font color='blue'>�磺����ҳ����Ϊ3����ÿ���Զ�����ǰ��ҳ����4ҳ�Ժ�ķ�ҳΪ�̶����ɵ�ҳ�棬����������������һҳ����������һ���̶�ҳ�棬���ܼ�¼������ÿҳ��¼����������ʱ������ҳ����3��4ҳ�����в��ּ�¼�ظ���</font></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td colspan='2'><font color='red'><b>���²�������������HTML��ʽ����Ϊ������ʱ����Ч��<br>�����ѡ����ÿһ�θ������²���ǰ���������ɾ��������ǰ���ɵ��ļ���Ȼ���ڱ���������ú����������������ļ���</b></font></td></tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>�Զ�����HTMLʱ�����ɷ�ʽ��</font></strong><br>���/�޸���Ϣʱ��ϵͳ�����Զ������й�ҳ���ļ�����������ѡ���Զ�����ʱ�ķ�ʽ��</td>"
        Response.Write "      <td>" & GetAutoCreateType(rsChannel("AutoCreateType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>��Ŀ�б��ļ��Ĵ��λ�ã�</font></strong></td>"
        Response.Write "      <td>" & GetListFileType(rsChannel("ListFileType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>Ŀ¼�ṹ��ʽ��</font></strong><br>ÿһ�θ���Ŀ¼�ṹ������Ҫ��ɾ��������ǰ���ļ������������������ļ���<br>��Ѱ治֧��Ŀ¼�ṹ�޸ġ�</td>"
        Response.Write "      <td>" & GetStructureType(rsChannel("StructureType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>����ҳ�ļ���������ʽ��</font></strong></td>"
        Response.Write "      <td>" & GetFileNameType(rsChannel("FileNameType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>Ƶ����ҳ����չ����</font></strong></td>"
        Response.Write "      <td>" & arrFileExt_Index(rsChannel("FileExt_Index")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>��Ŀҳ��ר��ҳ����չ����</font></strong></td>"
        Response.Write "      <td>" & arrFileExt_List(rsChannel("FileExt_List")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>����ҳ����չ����</font></strong></td>"
        Response.Write "      <td>" & arrFileExt_Item(rsChannel("FileExt_Item")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    End If
    '2005-12-23
    If IsCustom_Content = True And rsChannel("ModuleType") <> 4 And rsChannel("ModuleType") <> 6 And rsChannel("ModuleType") <> 7 And rsChannel("ModuleType") <> 8 Then
        Call EditCustom_Content("Modify", rsChannel("Custom_Content"), "Channel")
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center'><tr><td colspan='2' align='center'>"
    If rsChannel("ModuleType") <> 4 Then
        Response.Write "     <br><font color='red'>�ڸ���Ƶ���йز���ǰ���������ɾ��������ǰ���ɵ��ļ������Ĳ��������������������ļ���</font><br><br>" & vbCrLf
    End If
    Response.Write "     <input name='iChannelID' type='hidden' id='iChannelID' value='" & rsChannel("ChannelID") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input name='ModuleType' type='hidden' id='hidden' value='" & rsChannel("ModuleType") & "'>"
    Response.Write "        <input name='Submit'  type='submit' id='Submit' value='�����޸Ľ��'> &nbsp; <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Channel.asp'"" style='cursor:hand;'></td>"
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "</form>"

    rsChannel.Close
    Set rsChannel = Nothing
    Call ShowChekcFormJS
End Sub

Sub SaveAdd()
    Dim rsChannel
    Dim OrderID, ChannelType, LinkUrl

    ChannelName = Trim(Request("ChannelName"))
    ChannelShortName = Trim(Request("ChannelShortName"))
    ChannelItemUnit = Trim(Request("ChannelItemUnit"))
    LinkUrl = Trim(Request("LinkUrl"))
    ChannelType = PE_CLng(Trim(Request("ChannelType")))
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    ChannelDir = Trim(Request("ChannelDir"))
    UploadDir = ReplaceBadChar(Trim(Request("UploadDir")))
    If ChannelName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Ƶ�����Ʋ���Ϊ�գ�</li>"
    Else
        ChannelName = ReplaceBadChar(ChannelName)
    End If
    If ChannelType = 1 Then
        If ChannelDir = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>Ƶ��Ŀ¼����Ϊ�գ�</li>"
        ElseIf LCase(ChannelDir) = "others" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��Others���ѱ�ϵͳʹ�ã������Ƶ��Ŀ¼����</li>"
        Else
            If IsValidStr(ChannelDir) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>Ƶ��Ŀ¼��ֻ��ΪӢ����ĸ�����ֵ���ϣ��ҵ�һ���ַ�����ΪӢ����ĸ��</li>"
            Else
                If fso.FolderExists(Server.MapPath(InstallDir & ChannelDir)) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>Ƶ��Ŀ¼�Ѿ����ڣ�������ָ��һ��Ŀ¼��</li>"
                End If
            End If
        End If
        If ChannelShortName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����Ŀ���ƣ�</li>"
        End If
        If UploadDir = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ���ϴ�Ŀ¼</li>"
        End If
    Else
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӵ�ַ����Ϊ�գ�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If ChannelItemUnit = "" Then ChannelItemUnit = "��"
    
    ChannelID = GetNewID("PE_Channel", "ChannelID")
    OrderID = GetNewID("PE_Channel", "OrderID")
    
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    rsChannel.Open "Select * from PE_Channel Where ChannelName='" & ChannelName & "'", Conn, 1, 3
    If Not (rsChannel.BOF And rsChannel.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Ƶ�������Ѿ����ڣ�</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Sub
    End If
    rsChannel.addnew
    rsChannel("ChannelID") = ChannelID
    rsChannel("OrderID") = OrderID
    rsChannel("ChannelName") = ChannelName
    rsChannel("ChannelType") = ChannelType
    If ChannelType = 1 Then
        If SiteUrlType = 0 Then
            rsChannel("LinkUrl") = ""
        Else
            rsChannel("LinkUrl") = Trim(Request("ChannelUrl"))
        End If
    Else
        rsChannel("LinkUrl") = LinkUrl
    End If
    rsChannel("ModuleType") = ModuleType
    rsChannel("ChannelDir") = ChannelDir
    rsChannel("ChannelShortName") = ChannelShortName
    rsChannel("ChannelItemUnit") = ChannelItemUnit
    rsChannel("UploadDir") = UploadDir

    Call SaveChannel(rsChannel)

    rsChannel("ItemCount") = 0
    rsChannel("ItemChecked") = 0
    rsChannel("CommentCount") = 0
    rsChannel("SpecialCount") = 0
    rsChannel("HitsCount") = 0

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
    rsChannel("Custom_Content") = Custom_Content
    If ModuleType = 1 then
        Dim rsCheckChannel
        Set rsCheckChannel = Conn.Execute("Select * from PE_MailChannel where ChannelID = "&ChannelID)
        If rsCheckChannel.bof and rsCheckChannel.eof Then
            Conn.Execute ("insert into PE_MailChannel(ChannelID,UserID,arrClass,SendNum,IsUse) values("&ChannelID&",'','',10," & PE_False & ")")
        End If
        rsCheckChannel.Close
        set rsCheckChannel = nothing	
    End If
    rsChannel.Update
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteEntry(2, AdminName, "���Ƶ���ɹ���" & ChannelName)

    If ChannelType = 1 Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table PE_Admin add AdminPurview_" & ChannelDir & " Int null")
        Else
            Conn.Execute ("alter table PE_Admin add COLUMN AdminPurview_" & ChannelDir & " INTEGER")	
        End If
        Call CreateChannelDir(ChannelID, ChannelDir, UploadDir, ModuleType)
        Call AddTemplate(ModuleType, ChannelID)
        Call AddJsFile(ModuleType, ChannelID)
        Call ReloadLeft("Admin_Channel.asp")
    Else
        Call CloseConn
        Response.Redirect "Admin_Channel.asp"
    End If
End Sub

Sub SaveModify()
    Dim ChannelType, LinkUrl
    Dim rsChannel, sqlChannel
    Dim CommandChannelPoint '������ϢҪ�۳��ĵ������趨Ƶ���Ƽ�ÿ��Ҫ�۳��ĵ���
    ChannelID = Trim(Request("iChannelID"))
    ChannelType = PE_CLng(Trim(Request("ChannelType")))
    ChannelName = Trim(Request("ChannelName"))
    ChannelShortName = Trim(Request("ChannelShortName"))
    ChannelItemUnit = Trim(Request("ChannelItemUnit"))
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    LinkUrl = Trim(Request("LinkUrl"))
    UploadDir = Trim(Request("UploadDir"))
    CommandChannelPoint = PE_CLng(Trim(Request("CommandChannelPoint")))
 
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�Ƶ��ID</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If ChannelName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Ƶ�����Ʋ���Ϊ�գ�</li>"
    Else
        ChannelName = ReplaceBadChar(ChannelName)
    End If
    If ChannelType = 1 Then
        If ModuleType <> 4 Then
            If ChannelShortName = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��ָ����Ŀ���ƣ�</li>"
            End If
            If UploadDir = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��ָ���ϴ�Ŀ¼</li>"
            End If
        End If
    Else
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӵ�ַ����Ϊ�գ�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    sqlChannel = "Select * from PE_Channel Where ChannelID=" & ChannelID
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    rsChannel.Open sqlChannel, Conn, 1, 3
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����Ƶ����</li>"
        rsChannel.Close
        Set rsChannel = Nothing
    Else
        If ChannelType = 1 Then
            If RenameDir(InstallDir & rsChannel("ChannelDir") & "/" & rsChannel("UploadDir"), InstallDir & rsChannel("ChannelDir") & "/" & UploadDir) = True Then
                rsChannel("UploadDir") = UploadDir
            End If
        End If
        rsChannel("ChannelName") = ChannelName
        rsChannel("ChannelShortName") = ChannelShortName
        If ChannelItemUnit <> "" Then
            rsChannel("ChannelItemUnit") = ChannelItemUnit
        End If
        If ChannelType = 1 Then
            If SiteUrlType = 0 Then
                rsChannel("LinkUrl") = ""
            Else
                rsChannel("LinkUrl") = Trim(Request("ChannelUrl"))
            End If
        Else
            rsChannel("LinkUrl") = LinkUrl
        End If
        rsChannel("CommandChannelPoint") = CommandChannelPoint
        Call SaveChannel(rsChannel)

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
        rsChannel("Custom_Content") = Custom_Content

        rsChannel.Update
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    'Call ClearSiteCache(0)
    Call WriteEntry(2, AdminName, "�޸�Ƶ���ɹ���" & ChannelName)

    If ChannelType = 1 Then
        Call ReloadLeft("Admin_Channel.asp")
    Else
        Call CloseConn
        Response.Redirect "Admin_Channel.asp"
    End If
End Sub

Sub SaveChannel(rsChannel)
    Dim ChannelPurview, UseCreateHTML
    ChannelPurview = PE_CLng(Trim(Request("ChannelPurview")))
    UseCreateHTML = PE_CLng(Trim(Request("UseCreateHTML")))
    If ChannelPurview = 1 Then
        UseCreateHTML = 0
    End If
    rsChannel("ChannelPicUrl") = Trim(Request("ChannelPicUrl"))
    rsChannel("ReadMe") = Trim(Request("ReadMe"))
    rsChannel("OpenType") = PE_CLng(Trim(Request("OpenType")))
    rsChannel("ChannelPurview") = ChannelPurview
    rsChannel("arrGroupID") = Trim(Request("GroupID"))
    rsChannel("CheckLevel") = PE_CLng(Trim(Request("CheckLevel")))
    rsChannel("EmailOfReject") = Trim(Request("EmailOfReject"))
    rsChannel("EmailOfPassed") = Trim(Request("EmailOfPassed"))
    rsChannel("Meta_Keywords") = Trim(Request("Meta_Keywords"))
    rsChannel("Meta_Description") = Trim(Request("Meta_Description"))
    rsChannel("arrEnabledTabs") = ReplaceBadChar(Trim(Request("arrEnabledTabs")))
    rsChannel("MoneyPerKw") = PE_CDbl(Trim(Request("MoneyPerKw")))

    rsChannel("Disabled") = PE_CBool(Trim(Request("Disabled")))
    rsChannel("ShowName") = PE_CBool(Trim(Request("ShowName")))
    rsChannel("ShowNameOnPath") = PE_CBool(Trim(Request("ShowNameOnPath")))
    rsChannel("ShowClassTreeGuide") = PE_CBool(Trim(Request("ShowClassTreeGuide")))
    rsChannel("ShowSuspensionPoints") = PE_CBool(Trim(Request("ShowSuspensionPoints")))
    rsChannel("EnableUploadFile") = PE_CBool(Trim(Request("EnableUploadFile")))

    rsChannel("MaxFileSize") = PE_CLng(Trim(Request("MaxFileSize")))
    rsChannel("UpFileType") = Replace(Trim(Request("UpFileType")), ",", "$")
    rsChannel("HitsOfHot") = PE_CLng(Trim(Request("HitsOfHot")))
    rsChannel("DaysOfNew") = PE_CLng(Trim(Request("DaysOfNew")))
    rsChannel("MaxPerLine") = PE_CLng(Trim(Request("MaxPerLine")))
    rsChannel("AuthorInfoLen") = PE_CLng(Trim(Request("AuthorInfoLen")))
    rsChannel("JS_SpecialNum") = PE_CLng(Trim(Request("JS_SpecialNum")))
    rsChannel("MaxPerPage_Index") = PE_CLng(Trim(Request("MaxPerPage_Index")))
    rsChannel("MaxPerPage_SearchResult") = PE_CLng(Trim(Request("MaxPerPage_SearchResult")))
    rsChannel("MaxPerPage_New") = PE_CLng(Trim(Request("MaxPerPage_New")))
    rsChannel("MaxPerPage_Hot") = PE_CLng(Trim(Request("MaxPerPage_Hot")))
    rsChannel("MaxPerPage_Elite") = PE_CLng(Trim(Request("MaxPerPage_Elite")))
    rsChannel("MaxPerPage_SpecialList") = PE_CLng(Trim(Request("MaxPerPage_SpecialList")))
    rsChannel("Template_Index") = PE_CLng(Trim(Request("Template_Index")))
    rsChannel("DefaultSkinID") = PE_CLng(Trim(Request("DefaultSkinID")))
    rsChannel("TopMenuType") = PE_CLng(Trim(Request("TopMenuType")))
    rsChannel("ClassGuideType") = PE_CLng(Trim(Request("ClassGuideType")))
    rsChannel("UseCreateHTML") = UseCreateHTML
    rsChannel("StructureType") = PE_CLng(Trim(Request("StructureType")))
    rsChannel("ListFileType") = PE_CLng(Trim(Request("ListFileType")))
    rsChannel("FileNameType") = PE_CLng(Trim(Request("FileNameType")))
    rsChannel("AutoCreateType") = PE_CLng(Trim(Request("AutoCreateType")))
    rsChannel("FileExt_Index") = PE_CLng(Trim(Request("FileExt_Index")))
    rsChannel("FileExt_List") = PE_CLng(Trim(Request("FileExt_List")))
    rsChannel("FileExt_Item") = PE_CLng(Trim(Request("FileExt_Item")))
    rsChannel("UpdatePages") = PE_CLng1(Trim(Request("UpdatePages")))
    rsChannel("EnableComment") = PE_CBooL(Trim(Request("EnableComment")))
    rsChannel("CheckComment") = PE_CBooL(Trim(Request("CheckComment")))		
End Sub

Sub CreateChannelDir(iChannelID, DirName, sUploadDir, iModuleType)
    On Error Resume Next
    Dim fsfl, fl, fsfm, fm, strDir
    If Not fso.FolderExists(Server.MapPath(InstallDir & DirName)) Then
        fso.CreateFolder Server.MapPath(InstallDir & DirName)
    End If
    Select Case iModuleType
    Case 1
        strDir = "Article"
    Case 2
        strDir = "Soft"
    Case 3
        strDir = "Photo"
    Case 5
        strDir = "Shop"
    End Select
    Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & strDir))
    For Each fl In fsfl.Files
        If LCase(Left(fl.name, 7)) <> LCase(strDir) And Not IsNumeric(Left(fl.name, InStr(fl.name, ".") - 1)) And GetFileExt(fl.name) = "asp" Then
            fl.Copy Server.MapPath(InstallDir & DirName & "/" & fl.name), True
        End If
    Next
    Set fsfl = Nothing
    
    Set fl = fso.CreateTextFile(Server.MapPath(InstallDir & DirName & "/Channel_Config.asp"), True)
    fl.WriteLine ("<" & "%")
    fl.WriteLine ("ChannelID = " & iChannelID)
    fl.WriteLine ("%" & ">")
    fl.Close
    Set fl = Nothing
    'Set fl = fso.CreateTextFile(Server.MapPath(InstallDir & DirName & "/Index.asp"), True)
    'fl.WriteLine ("<!" & "--#include file=""CommonCode.asp""" & "-->")
    'fl.WriteLine ("<" & "%")
    'fl.WriteLine ("Call PE_" & strDir & ".ShowIndex")
    'fl.WriteLine ("Set PE_" & strDir & " = Nothing")
    'fl.WriteLine ("%" & ">")
    'fl.Close
    'Set fl = Nothing

    If Trim(sUploadDir) <> "" Then
        If Not fso.FolderExists(Server.MapPath(InstallDir & DirName & "/" & Trim(sUploadDir))) Then
            fso.CreateFolder Server.MapPath(InstallDir & DirName & "/" & Trim(sUploadDir))
        End If
    End If
    If Not fso.FolderExists(Server.MapPath(InstallDir & DirName & "/Images")) Then
        fso.CreateFolder Server.MapPath(InstallDir & DirName & "/Images")
    End If
    Set fsfm = fso.GetFolder(Server.MapPath(InstallDir & strDir & "/Images"))
    For Each fm In fsfm.Files
        fm.Copy Server.MapPath(InstallDir & DirName & "/Images/" & fm.name), True
    Next
    fm.Close
    Set fm = Nothing
    Set fsfm = Nothing
End Sub

Function RenameDir(strFolderName, strTargetName)
    RenameDir = False
    On Error Resume Next
    If LCase(strFolderName) = LCase(strTargetName) Then Exit Function
    If Not fso.FolderExists(Server.MapPath(strFolderName)) Then Exit Function
    If fso.FolderExists(Server.MapPath(strTargetName)) Then Exit Function
    fso.MoveFolder Server.MapPath(strFolderName), Server.MapPath(strTargetName)
    If Err Then
        Err.Clear
    Else
        RenameDir = True
    End If
End Function

Sub DelChannelDir(DirName)
    On Error Resume Next
    If IsNull(DirName) Or Trim(DirName) = "" Then Exit Sub
    If fso.FolderExists(Server.MapPath(InstallDir & DirName)) Then
        fso.DeleteFolder Server.MapPath(InstallDir & DirName)
    End If
End Sub

Sub DelChannel()
    On Error Resume Next
    Dim ChannelID, rsChannel, sqlChannel
    ChannelID = Trim(Request("iChannelID"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ����Ƶ��ID</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If

    sqlChannel = "Select * from PE_Channel Where ChannelID=" & ChannelID
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    rsChannel.Open sqlChannel, Conn, 1, 3
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����Ƶ����</li>"
    Else
        If rsChannel("ChannelType") = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ɾ��ϵͳƵ��������㲻���ô�Ƶ�������Խ��ô�Ƶ����</li>"
        Else
            If rsChannel("ChannelType") = 1 Then
                Call DelChannelDir(rsChannel("ChannelDir"))

                'ɾ������Ƶ������
                Dim Infotable
                If rsChannel("ModuleType") = 1 Then
                    Infotable = "Article"
                ElseIf rsChannel("ModuleType") = 2 Then
                    Infotable = "Photo"
                ElseIf rsChannel("ModuleType") = 3 Then
                    Infotable = "Soft"
                ElseIf rsChannel("ModuleType") = 5 Then
                    Infotable = "Product"
                End If
                
                Dim rsComment
                Set rsComment = Conn.Execute("Select I." & Infotable & "ID,I.ChannelID,C.InfoID from PE_" & Infotable & " I inner join PE_Comment C on I." & Infotable & "ID=C.InfoID where  I.ChannelID=" & ChannelID & "")
               
                Do While Not rsComment.EOF
                    Conn.Execute "delete from PE_Comment where InfoID=" & rsComment("InfoID")
                    rsComment.MoveNext
                Loop
         
                Set rsComment = Nothing

                Dim rs
                Set rs = Conn.Execute("Select FieldName From PE_Field Where ChannelID=" & ChannelID)
                Do While Not rs.EOF
                    Conn.Execute ("alter table PE_" & Infotable & " drop COLUMN " & rs("FieldName") & "")
                    rs.MoveNext
                Loop
                Set rs = Nothing
                Conn.Execute ("alter table PE_Admin drop COLUMN AdminPurview_" & rsChannel("ChannelDir") & "")
                Conn.Execute ("delete from PE_Class where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Special where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Article where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Soft where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Photo where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_JsFile where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Template where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Announce where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Vote where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Author where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_CopyFrom where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Favorite where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Field where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_KeyLink where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_NewKeys where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Item where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_HistrolyNews where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_MailChannel where ChannelID=" & ChannelID)
                
            End If
            rsChannel.Delete
            rsChannel.Update
            Call ReloadLeft("Admin_Channel.asp")
        End If
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteEntry(2, AdminName, "ɾ��Ƶ���ɹ���ChannelID��" & ChannelID)
    'Call ClearSiteCache(0)
End Sub

Sub DisabledChannel(ActionType)
    Dim ChannelID
    ChannelID = Trim(Request("iChannelID"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ƶ��ID</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If ActionType = 0 Then
        Conn.Execute ("update PE_Channel set Disabled=" & PE_True & " where ChannelID=" & ChannelID)
    Else
        Conn.Execute ("update PE_Channel set Disabled=" & PE_False & " where ChannelID=" & ChannelID)
    End If
    Call ReloadLeft("Admin_Channel.asp")
    'Call ClearSiteCache(0)
End Sub

Function getMoveNum(ByVal ChannelID, ByVal MoveNum)
    Dim sqlChannelList, rsOrder, i
    sqlChannelList = "Select OrderID,ModuleType From PE_Channel Where 1 <> 1 "
    If Not (FoundInArr(AllModules, "Supply", ",")) Then
        sqlChannelList = sqlChannelList & " Or ModuleType = 6"
    End If
    If Not (FoundInArr(AllModules, "Job", ",")) Then
        sqlChannelList = sqlChannelList & " Or  ModuleType = 8"
    End If
    If Not (FoundInArr(AllModules, "House", ",")) Then
        sqlChannelList = sqlChannelList & " Or  ModuleType = 7"
    End If
    Set rsOrder = Server.CreateObject("Adodb.Recordset")
    Dim CurrentOrderID
    CurrentOrderID = PE_CLng(Conn.Execute("Select OrderID From PE_Channel Where ChannelID = " & ChannelID & "")(0))
    rsOrder.Open sqlChannelList, Conn, 1, 1
     i = 0
    Do While Not rsOrder.EOF
        If CurrentOrderID > rsOrder("OrderID") Then
            i = i + 1
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    getMoveNum = MoveNum + i
End Function

Sub UpOrder()
    Dim ChannelID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsChannel
    ChannelID = Trim(Request("iChannelID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        cOrderID = PE_CLng(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ��Ҫ���������֣�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    MoveNum = getMoveNum(ChannelID, MoveNum)
    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_Channel")
    MaxOrderID = mrs(0) + 1
    '�Ƚ���ǰ��Ŀ������󣬰�������Ŀ
    Conn.Execute ("update PE_Channel set OrderID=" & MaxOrderID & " where ChannelID=" & ChannelID)
    
    'Ȼ��λ�ڵ�ǰ��Ŀ���ϵ���Ŀ��OrderID���μ�һ����ΧΪҪ����������
    sqlOrder = "select * from PE_Channel where OrderID<" & cOrderID & "  order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '�����ǰ��Ŀ�Ѿ��������棬�������ƶ�
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '�õ�Ҫ����λ�õ�OrderID����������Ŀ
        Conn.Execute ("update PE_Channel set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    'Ȼ���ٽ���ǰ��Ŀ������Ƶ���Ӧλ�ã���������Ŀ
    Conn.Execute ("update PE_Channel set OrderID=" & tOrderID & " where ChannelID=" & ChannelID)

    Call ReloadLeft("Admin_Channel.asp?Action=Order")
    'Call ClearSiteCache(0)
End Sub

Sub DownOrder()
    Dim ChannelID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsChannel, PrevID, NextID
    ChannelID = Trim(Request("iChannelID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        cOrderID = PE_CLng(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������</li>"
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ��Ҫ���������֣�</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_Channel")
    MaxOrderID = mrs(0) + 1
    '�Ƚ���ǰ��Ŀ������󣬰�������Ŀ
    Conn.Execute ("update PE_Channel set OrderID=" & MaxOrderID & " where ChannelID=" & ChannelID)
    
    'Ȼ��λ�ڵ�ǰ��Ŀ���µ���Ŀ��OrderID���μ�һ����ΧΪҪ�½�������
    sqlOrder = "select * from PE_Channel where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '�����ǰ��Ŀ�Ѿ��������棬�������ƶ�
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '�õ�Ҫ����λ�õ�OrderID����������Ŀ
        Conn.Execute ("update PE_Channel set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    'Ȼ���ٽ���ǰ��Ŀ������Ƶ���Ӧλ�ã���������Ŀ
    Conn.Execute ("update PE_Channel set OrderID=" & tOrderID & " where ChannelID=" & ChannelID)
    
    Call ReloadLeft("Admin_Channel.asp?Action=Order")
    'Call ClearSiteCache(0)
End Sub

Sub UpdateData()
    Call UpdateChannelData(PE_CLng(Trim(Request("iChannelID"))))

    Call WriteSuccessMsg("����Ƶ�����ݳɹ���", ComeUrl)
    'Call ClearSiteCache(0)
End Sub

Sub UpdateChannelFiles()
    Dim iChannelID, rsChannel, sqlChannel, DirName, ModuleType, UploadDir
    Dim fsfl, fl, fsfm, fm, strDir
    
    iChannelID = PE_CLng(Trim(Request("iChannelID")))
    If iChannelID > 0 Then
        sqlChannel = "select ChannelID,ChannelDir,ModuleType,UploadDir from PE_Channel where ChannelType = 1 and ChannelID=" & iChannelID & ""
    Else
        sqlChannel = "select ChannelID,ChannelDir,ModuleType,UploadDir from PE_Channel where ChannelType = 1 And ModuleType<>4 And ModuleType<>6 And ModuleType<>7 And ModuleType<>8"
    End If
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        Call CreateChannelDir(rsChannel("ChannelID"), rsChannel("ChannelDir"), rsChannel("UploadDir"), rsChannel("ModuleType"))
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteSuccessMsg("����Ƶ���ļ��ɹ���", ComeUrl)
End Sub

Function Admin_GetSkin_Option(iSkinID)
    Dim sqlSkin, rsSkin, strSkin
    If IsNull(iSkinID) Then iSkinID = 0
    strSkin = ""
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Conn.Execute(sqlSkin)
    If rsSkin.BOF And rsSkin.EOF Then
        strSkin = strSkin & "<option value=''>������ӷ��</option>"
    Else
        If iSkinID = 0 Then
            strSkin = strSkin & "<option value='0' selected>ʹ��ϵͳ��Ĭ�Ϸ��</option>"
            Do While Not rsSkin.EOF
                If rsSkin("IsDefault") = True Then
                    strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'>" & rsSkin("SkinName") & "��Ĭ�ϣ�</option>"
                Else
                    strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'>" & rsSkin("SkinName") & "</option>"
                End If
                rsSkin.MoveNext
            Loop
        Else
            strSkin = strSkin & "<option value='0'>ʹ��ϵͳ��Ĭ�Ϸ��</option>"
            Do While Not rsSkin.EOF
                strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'"
                If rsSkin("SkinID") = iSkinID Then
                    strSkin = strSkin & " selected"
                End If
                strSkin = strSkin & ">" & rsSkin("SkinName")
                If rsSkin("IsDefault") = True Then
                    strSkin = strSkin & "��Ĭ�ϣ�"
                End If
                strSkin = strSkin & "</option>"
                rsSkin.MoveNext
            Loop
        End If
    End If
    rsSkin.Close
    Set rsSkin = Nothing
    Admin_GetSkin_Option = strSkin
End Function

Sub AddTemplate(ChannelID_Source, ChannelID_Target)
    
    Dim sqlTemplate, rsTemplate, trs
    '���´���ֻ���Ƶ�ǰ�����µ�ģ�壬�����ڷ����л��ᶪʧģ�壬��ʱ���õ�
    'Dim rsProjectName, ProjectName
    'Set rsProjectName = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    'If rsProjectName.BOF And rsProjectName.EOF Then
    '    Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
    '    Exit Sub
    'Else
    '    ProjectName = rsProjectName("TemplateProjectName")
    'End If
    'Set rsProjectName = Nothing
    'Set trs = Conn.Execute("select * from PE_Template where ChannelID=" & ChannelID_Source & " and ProjectName='" & ProjectName & "'")
    
    sqlTemplate = "select top 1 * from PE_Template"
    Set rsTemplate = Server.CreateObject("adodb.recordset")
    rsTemplate.Open sqlTemplate, Conn, 1, 3

    Set trs = Conn.Execute("select * from PE_Template where ChannelID=" & ChannelID_Source)
    Do While Not trs.EOF
        rsTemplate.addnew
        rsTemplate("ChannelID") = ChannelID_Target
        rsTemplate("TemplateName") = trs("TemplateName")
        rsTemplate("TemplateType") = trs("TemplateType")
        rsTemplate("TemplateContent") = trs("TemplateContent")
        rsTemplate("IsDefault") = trs("IsDefault")
        rsTemplate("ProjectName") = trs("ProjectName")
        rsTemplate("IsDefaultInProject") = trs("IsDefaultInProject")
        rsTemplate("Deleted") = trs("Deleted")
        rsTemplate.Update
        trs.MoveNext
    Loop
    rsTemplate.Close
    Set rsTemplate = Nothing
    Set trs = Nothing
End Sub

Sub AddJsFile(ChannelID_Source, ChannelID_Target)
    Dim sqlJsFile, rsJsFile, trs
    sqlJsFile = "select top 1 * from PE_JsFile"
    Set rsJsFile = Server.CreateObject("adodb.recordset")
    rsJsFile.Open sqlJsFile, Conn, 1, 3
    Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID_Source)
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = "<li>�Ҳ���ָ����JsFile</li>"
        Exit Sub
    End If
    Do While Not trs.EOF
        rsJsFile.addnew
        rsJsFile("ChannelID") = ChannelID_Target
        rsJsFile("JsName") = trs("JsName")
        rsJsFile("JsReadme") = trs("JsReadme")
        rsJsFile("JsFileName") = trs("JsFileName")
        rsJsFile("JsType") = trs("JsType")
        rsJsFile("Config") = trs("Config")
        rsJsFile.Update
        trs.MoveNext
    Loop
    rsJsFile.Close
    Set rsJsFile = Nothing
    Set trs = Nothing
End Sub

Sub ReloadLeft(strUrl)
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "  parent.left.location.reload();" & vbCrLf
    Response.Write "  window.location.href='" & strUrl & "';"
    Response.Write "</script>" & vbCrLf
End Sub

Function GetMenuType_Option(MenuType)
    Dim strMenuType
    strMenuType = strMenuType & "<option " & OptionValue(MenuType, 1) & ">�޼�����˵�</option>"
    strMenuType = strMenuType & "<option " & OptionValue(MenuType, 2) & ">��ͨ�����˵�</option>"
    strMenuType = strMenuType & "<option " & OptionValue(MenuType, 3) & ">�޲˵�</option>"
    GetMenuType_Option = strMenuType
End Function

Function GetGuideType_Option(GuideType)
    Dim strGuideType
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 1) & ">ƽ��ʽ��ÿ��2������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 2) & ">ƽ��ʽ��ÿ��3������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 3) & ">ƽ��ʽ��ÿ��4������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 4) & ">ƽ��ʽ��ÿ��5������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 5) & ">ƽ��ʽ��ÿ��6������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 6) & ">ƽ��ʽ��ÿ��7������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 7) & ">ƽ��ʽ��ÿ��8������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 8) & ">����ʽ��һ�У�ÿ������ʾ2������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 9) & ">����ʽ��һ�У�ÿ������ʾ3������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 10) & ">����ʽ��һ�У�ÿ������ʾ4������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 11) & ">����ʽ��һ�У�ÿ������ʾ5������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 12) & ">����ʽ��һ�У�ÿ������ʾ6������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 13) & ">����ʽ��һ�У�ÿ������ʾ7������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 14) & ">����ʽ��һ�У�ÿ������ʾ8������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 15) & ">����ʽ�����У�ÿ������ʾ2������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 16) & ">����ʽ�����У�ÿ������ʾ3������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 17) & ">����ʽ�����У�ÿ������ʾ4������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 18) & ">����ʽ�����У�ÿ������ʾ5������Ŀ��</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 19) & ">����ʽ�����У�ÿ������ʾ6������Ŀ��</option>"
    GetGuideType_Option = strGuideType
End Function


Function GetOrderType_Option(OrderType)
    Dim strOrderType
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 1) & ">" & ChannelShortName & "ID������</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 2) & ">" & ChannelShortName & "ID������</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 3) & ">����ʱ�䣨����</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 4) & ">����ʱ�䣨����</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 5) & ">�������������</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 6) & ">�������������</option>"
    GetOrderType_Option = strOrderType
End Function

Function GetUseCreateHTML(UseCreateHTML, ModuleType)
    Dim strUseCreateHTML
    strUseCreateHTML = strUseCreateHTML & "<input name='UseCreateHTML' type='radio' value='0'"
    If UseCreateHTML = 0 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " checked"
    strUseCreateHTML = strUseCreateHTML & ">������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>��Ƶ���е���Ϣ���Ƚ��٣���1000��ʱ������ѡ�ô��ַ�ʽ���˷�ʽ��ķ�ϵͳ��Դ��</font><br>"
    strUseCreateHTML = strUseCreateHTML & "<input type='radio' name='UseCreateHTML' value='1'"
    If UseCreateHTML = 1 And ObjInstalled_FSO = True Then strUseCreateHTML = strUseCreateHTML & " checked"
    If ModuleType = 4 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " disabled "
    strUseCreateHTML = strUseCreateHTML & ">ȫ������&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>�˷�ʽ�����ɺ����ʡϵͳ��Դ��������Ϣ����ʱ�����ɹ��̽��Ƚϳ���</font><br>"
    strUseCreateHTML = strUseCreateHTML & "<input type='radio' name='UseCreateHTML' value='2'"
    If UseCreateHTML = 2 And ObjInstalled_FSO = True Then strUseCreateHTML = strUseCreateHTML & " checked"
    If ModuleType = 4 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " disabled "
    strUseCreateHTML = strUseCreateHTML & ">��ҳ������ҳΪHTML����Ŀ��ר��ҳΪASP<br>"
    strUseCreateHTML = strUseCreateHTML & "<input type='radio' name='UseCreateHTML' value='3'"
    If UseCreateHTML = 3 And ObjInstalled_FSO = True Then strUseCreateHTML = strUseCreateHTML & " checked"
    If ModuleType = 4 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " disabled "
    strUseCreateHTML = strUseCreateHTML & ">��ҳ������ҳ����Ŀ��ר�����ҳΪHTML������ҳΪASP <font color='red'><b>���Ƽ���</b></font>"
    GetUseCreateHTML = strUseCreateHTML
End Function

Function GetAutoCreateType(AutoCreateType)
    Dim strAutoCreateType
    strAutoCreateType = strAutoCreateType & "<input name='AutoCreateType' type='radio' " & RadioValue(AutoCreateType, 0) & ">���Զ����ɣ��ɹ���Ա�ֹ��������ҳ��<br>"
    strAutoCreateType = strAutoCreateType & "<input name='AutoCreateType' type='radio' " & RadioValue(AutoCreateType, 1) & ">�Զ�����ȫ������ҳ��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>��������HTML��ʽ������Ϊ��ȫ�����ɡ�ʱ������������ҳ�棻��������HTML��ʽ������Ϊ������ʱ����������õ�ѡ�������й�ҳ�档</font><br>"
    strAutoCreateType = strAutoCreateType & "<input name='AutoCreateType' type='radio' " & RadioValue(AutoCreateType, 2) & ">�Զ����ɲ�������ҳ��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����������HTML��ʽ������Ϊ��ȫ�����ɡ�ʱ����Ч���˷�ʽֻ������ҳ������ҳ����Ŀ��ר�����ҳ������ҳ���ɹ���Ա�ֹ����ɡ�</font><br>"
    GetAutoCreateType = strAutoCreateType
End Function

Function GetListFileType(ListFileType)
    Dim strListFileType
    strListFileType = strListFileType & "<input name='ListFileType' type='radio' " & RadioValue(ListFileType, 0) & ">�б��ļ���Ŀ¼������������Ŀ���ļ�����<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/ASP/JiChu/index.html����Ŀ��ҳ��<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Article/ASP/JiChu/List_2.html���ڶ�ҳ��</font><br>"
    strListFileType = strListFileType & "<input name='ListFileType' type='radio' " & RadioValue(ListFileType, 1) & ">�б��ļ�ͳһ������ָ���ġ�List���ļ�����<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/List/List_236.html����Ŀ��ҳ��<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Article/List/List_236_2.html���ڶ�ҳ��</font><br>"
    strListFileType = strListFileType & "<input name='ListFileType' type='radio' " & RadioValue(ListFileType, 2) & ">�б��ļ�ͳһ������Ƶ���ļ�����<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/List_236.html����Ŀ��ҳ��<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Article/List_236_2.html���ڶ�ҳ��</font><br>"
    GetListFileType = strListFileType
End Function

Function GetStructureType(StructureType)
    Dim strStructureType
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 0) & ">Ƶ��/����/С��/�·�/�ļ�����Ŀ�ּ����ٰ��·ݱ��棩<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/ASP/JiChu/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 1) & ">Ƶ��/����/С��/����/�ļ�����Ŀ�ּ����ٰ����ڷ֣�ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/ASP/JiChu/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 2) & ">Ƶ��/����/С��/�ļ�����Ŀ�ּ������ٰ��·ݣ�<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/ASP/JiChu/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 3) & ">Ƶ��/��Ŀ/�·�/�ļ�����Ŀƽ�����ٰ��·ݱ��棩<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/JiChu/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 4) & ">Ƶ��/��Ŀ/����/�ļ�����Ŀƽ�����ٰ����ڷ֣�ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/JiChu/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 5) & ">Ƶ��/��Ŀ/�ļ�����Ŀƽ�������ٰ��·ݣ�<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/JiChu/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 6) & ">Ƶ��/�ļ���ֱ�ӷ���Ƶ��Ŀ¼�У�<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 7) & ">Ƶ��/HTML/�ļ���ֱ�ӷ���ָ���ġ�HTML���ļ����У�<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/HTML/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 8) & ">Ƶ��/���/�ļ���ֱ�Ӱ���ݱ��棬ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/2004/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 9) & ">Ƶ��/�·�/�ļ���ֱ�Ӱ��·ݱ��棬ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 10) & ">Ƶ��/����/�ļ���ֱ�Ӱ����ڱ��棬ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 11) & ">Ƶ��/���/�·�/�ļ����Ȱ���ݣ��ٰ��·ݱ��棬ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/2004/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 12) & ">Ƶ��/���/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/2004/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 13) & ">Ƶ��/�·�/����/�ļ����Ȱ��·ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/200408/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 14) & ">Ƶ��/���/�·�/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article/2004/200408/2004-08-25/1368.html</font>"
    GetStructureType = strStructureType
End Function

Function GetFileNameType(FileNameType)
    Dim strFileNameType
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 0) & ">����ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����1358.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 1) & ">����ʱ��.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����20040828112308.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 2) & ">Ƶ��Ӣ����_����ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article_1358.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 3) & ">Ƶ��Ӣ����_����ʱ��.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article_20040828112308.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 4) & ">����ʱ��_ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����20040828112308_1358.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 5) & ">Ƶ��Ӣ����_����ʱ��_ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>����Article_20040828112308_1358.html</font>"
    GetFileNameType = strFileNameType
End Function

Function arrFileExt_Index(FileExt_Index)
    Dim strFileExt_Index
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 0) & ">.html&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 1) & ">.htm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 2) & ">.shtml&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 3) & ">.shtm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 4) & ">.asp"
    arrFileExt_Index = strFileExt_Index
End Function

Function arrFileExt_List(FileExt_List)
    Dim strFileExt_List
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 0) & ">.html&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 1) & ">.htm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 2) & ">.shtml&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 3) & ">.shtm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 4) & ">.asp"
    arrFileExt_List = strFileExt_List
End Function

Function arrFileExt_Item(FileExt_Item)
    Dim strFileExt_Item
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 0) & ">.html&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 1) & ">.htm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 2) & ">.shtml&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 3) & ">.shtm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 4) & ">.asp"
    arrFileExt_Item = strFileExt_Item
End Function


Function GetModuleTypeName(ModuleType)
    Dim strModuleType
    Select Case ModuleType
    Case 1
        strModuleType = "����"
    Case 2
        strModuleType = "����"
    Case 3
        strModuleType = "ͼƬ"
    Case 4
        strModuleType = "���԰�"
    Case 5
        strModuleType = "�̳�"
    Case Else
        strModuleType = "����"
    End Select
    GetModuleTypeName = strModuleType
End Function
%>
