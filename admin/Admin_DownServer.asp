<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 1   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "DeliverType"   '����Ȩ��

Response.Write "<html><head><title>����������</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle(ChannelName & "����----�������������", 10123)

Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>��������</strong></td><td colspan='5'>"
Response.Write "    <a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'>������������ҳ</a>&nbsp;|&nbsp;<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&Action=Add'>����·�����</a>&nbsp;|&nbsp;<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&Action=Order'>����������</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveDownServer
Case "Del"
    Call Del
Case "Order"
    Call Order
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "SaveAllShowType"
    Call SaveAllShowType
Case Else
    Call Main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub Main()
    Dim rsDownServer, sqlDownServer
    sqlDownServer = "select * from PE_DownServer where ChannelID=" & ChannelID & " order by OrderID"
    Set rsDownServer = Conn.Execute(sqlDownServer)

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' align='center'>"
    Response.Write "    <td width='40'><strong>ID</strong></td>"
    Response.Write "    <td width='100'><strong>��������</strong></td>"
    Response.Write "    <td width='120'><strong>������LOGO</strong></td>"
    Response.Write "    <td width='60'><strong>��ʾ��ʽ</strong></td>"
    Response.Write "    <td><strong>��������ַ</strong></td>"
    Response.Write "    <td width='60'><strong>����</strong></td>"
    Response.Write "  </tr>"
    If rsDownServer.BOF And rsDownServer.EOF Then
        Response.Write "  <tr class='tdbg' align='center' height='50'><td colspan='10'>û���κ����ط�����</td></tr>"
    Else
        Do While Not rsDownServer.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='40' align='center'>" & rsDownServer("ServerID") & "</td>"
            Response.Write "    <td width='100' align='center'><a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&action=Modify&ServerID=" & rsDownServer("ServerID") & "'>" & rsDownServer("ServerName") & "</a></td>"
            Response.Write "    <td width='120' align='center'>"
            If rsDownServer("ServerLogo") <> "" Then
                Response.Write "<img src='" & rsDownServer("ServerLogo") & "'>"
            End If
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>"
            If rsDownServer("ShowType") = 1 Then
                Response.Write "��ʾLOGO"
            Else
                Response.Write "��ʾ����"
            End If
            Response.Write "</td>"
            Response.Write "    <td>" & rsDownServer("ServerUrl") & "</td>"
            Response.Write "    <td width='60' align='center'>"
            Response.Write "<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&action=Modify&ServerID=" & rsDownServer("ServerID") & "'>�޸�</a> "
            Response.Write "<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&Action=Del&ServerID=" & rsDownServer("ServerID") & "' onClick=""return confirm('ȷ��Ҫɾ���˷�������Ϣ��');"">ɾ��</a>"
            Response.Write "</td></tr>"
            rsDownServer.MoveNext
        Loop
    End If
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"
    Response.Write "  <tr>"
    Response.Write "    <td height='22' align='center'>"
    Response.Write "    <form action='Admin_DownServer.asp?Action=SaveAllShowType' method='post'>����������ʾ��ʽ"
    Response.Write "<select name='ShowType'><option value='0'>��ʾ����</option><option value='1'>��ʾLOGO</option></select></select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "<input type='submit' name='Submit' value='�� ��'>"
    Response.Write "</form>"
    Response.Write "</td></tr>"
    Response.Write "</table>"
    Response.Write "<br><b><font color=red>ע�⣺</font></b><br>&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>ɾ��ĳ�������������Ϣʱ,��֮��ص����ش�����ϢҲ��һ��ɾ������</font><br><br>"
    rsDownServer.Close
    Set rsDownServer = Nothing
End Sub

Sub Add()
    Call ShowJS_Soft
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'>�������������</a>&nbsp;&gt;&gt;&nbsp;��Ӿ��������</td></tr></table>"
    Response.Write "<form method='post' action='Admin_DownServer.asp' name='form1'>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table id='Tabs' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>���������ƣ�</strong><br>�ڴ�������ǰ̨��ʾ�ľ��������������㶫���ء��Ϻ����صȡ�</td>"
    Response.Write "          <td class='tdbg'><input name='ServerName' type='text' id='ServerName' size='50' maxlength='30'>&nbsp;</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>������LOGO��</strong><br>���������LOGO�ľ��Ե�ַ����http://www.powereasy.net/Soft/Images/ServerLogo.gif</td>"
    Response.Write "          <td class='tdbg'><input name='ServerLogo' type='text' id='ServerLogo' size='50' maxlength='200'>&nbsp;</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>��������ַ��</strong><br>������������ȷ�ķ�������ַ��<br>��http://www.powereasy.net/�����ĵ�ַ</td>"
    Response.Write "          <td class='tdbg'><input name='ServerUrl' type='text' id='ServerUrl' size='50' maxlength='200'>&nbsp;</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "          <td class='tdbg'><select name='ShowType' size=1><option value='0'>��ʾ����</option><option value='1'>��ʾLOGO</option></select></td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "      <table id='Tabs' style='display:none' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�Ķ�Ȩ�ޣ�</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0' checked>�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'>���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'>ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup("", "")
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�Ķ�������</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & Session("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'> "
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0������Ȩ�޵Ļ�Ա�Ķ���" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷��鿴��" & ChannelShortName & "</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ظ��շѣ�</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0' checked>���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'>�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='24' size='8' maxlength='8' style='text-align:center'> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'>��Ա�ظ��鿴������ <input name='ReadTimes' type='text' value='10' size='8' maxlength='8' style='text-align:center'> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'>�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'>����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'>ÿ�Ķ�һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ֳɱ�����</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='0' size='5' maxlength='4' style='text-align:center'> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰����������Ķ�����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "      </table>"
    Response.Write "      <br><br><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "      <input  type='submit' name='Submit' value=' �� �� '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "      <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'""  style='cursor:hand;'>" & vbCrLf
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub ShowJS_Soft()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "    if(ID==0){" & vbCrLf
    Response.Write "      editor.yToolbarsCss();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub SaveDownServer()
    Dim ServerID, ServerName, ServerUrl, ServerLogo, OrderID
    Dim rsDownServer, sqlDownServer
    Dim ShowType
    ServerID = PE_CLng(Trim(Request("ServerID")))
    ServerName = Trim(Request.Form("ServerName"))
    ServerUrl = Trim(Request.Form("ServerUrl"))
    ServerLogo = Trim(Request.Form("ServerLogo"))
    ShowType = PE_CLng(Request.Form("ShowType"))

    If ChannelID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Ƶ��ID��ʧ��</li>"
    End If
    If ServerName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������������Ϊ�գ�</li>"
    End If
    If ServerUrl = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������ַ����Ϊ�գ�</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    If Action = "SaveAdd" Then
        ServerID = GetNewID("PE_DownServer", "ServerId")
        OrderID = GetNewID("PE_DownServer", "OrderID")
        
        Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
        rsDownServer.Open "Select top 1 * from PE_DownServer", Conn, 1, 3
        rsDownServer.addnew
        'rsDownServer("ServerID") = ServerID
        'rsDownServer("OrderID") = OrderID
    Else
        If ServerID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵķ�����ID��</li>"
            Exit Sub
        End If
        sqlDownServer = "Select * from PE_DownServer Where ServerID=" & ServerID
        Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
        rsDownServer.Open sqlDownServer, Conn, 1, 3
        If rsDownServer.BOF And rsDownServer.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ���ķ������������Ѿ���ɾ����</li>"
            rsDownServer.Close
            Set rsDownServer = Nothing
            Exit Sub
        End If
    End If
    rsDownServer("ChannelID") = ChannelID
    rsDownServer("ServerName") = ServerName
    rsDownServer("ServerUrl") = ServerUrl
    rsDownServer("ServerLogo") = ServerLogo
    rsDownServer("ShowType") = ShowType

    rsDownServer("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
    rsDownServer("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
    rsDownServer("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
    rsDownServer("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
    rsDownServer("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
    rsDownServer("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
    rsDownServer("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))

    rsDownServer.Update
    rsDownServer.Close
    Set rsDownServer = Nothing
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?ChannelID=" & ChannelID & ""
End Sub


Sub Modify()
    Dim ServerID, rsDownServer, sqlDownServer
    ServerID = PE_CLng(Trim(Request("ServerID")))
    If ServerID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵķ�����ID��</li>"
        Exit Sub
    End If
    sqlDownServer = "Select * from PE_DownServer Where ServerID=" & ServerID
    Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
    rsDownServer.Open sqlDownServer, Conn, 1, 3
    If rsDownServer.BOF And rsDownServer.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���ķ������������Ѿ���ɾ����</li>"
        Exit Sub
    End If

    Call ShowJS_Soft
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'>�������������</a>&nbsp;&gt;&gt;&nbsp;�޸ľ������������</td></tr></table>"
    Response.Write "<form method='post' action='Admin_DownServer.asp' name='form1'>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table id='Tabs' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>���������ƣ�</strong><br>�ڴ�������ǰ̨��ʾ�ľ���������������ɽ���ء��������صȡ�</td>"
    Response.Write "      <td class='tdbg'><input name='ServerName' type='text' id='ServerName' size='50' maxlength='30' value='" & rsDownServer("ServerName") & "'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>������LOGO��</strong><br>���������LOGO�ľ��Ե�ַ����http://www.powereasy.net/Soft/Images/ServerLogo.gif</td>"
    Response.Write "      <td class='tdbg'><input name='ServerLogo' type='text' id='ServerLogo' size='50' maxlength='200' value='" & rsDownServer("ServerLogo") & "'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>��������ַ��</strong><br>������������ȷ�ķ�������ַ��<br>��http://www.powereasy.net/�����ĵ�ַ</td>"
    Response.Write "      <td class='tdbg'><input name='ServerUrl' type='text' id='ServerUrl' size='50' maxlength='200' value='" & rsDownServer("ServerUrl") & "'>&nbsp;</td>"
    Response.Write "    </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td class='tdbg'><select name='ShowType'><option value='0'"
    If rsDownServer("ShowType") = 0 Then Response.Write " selected"
    Response.Write ">��ʾ����</option>"
    Response.Write "<option value='1'"
    If rsDownServer("ShowType") = 1 Then Response.Write " selected"
    Response.Write ">��ʾLOGO</option>"
    Response.Write "</select>"
    Response.Write "</td></tr>"
    Response.Write "      </table>"
    Response.Write "      <table id='Tabs' style='display:none' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�Ķ�Ȩ�ޣ�</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0'"
    If rsDownServer("InfoPurview") = 0 Then Response.Write " checked"
    Response.Write ">�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'"
    If rsDownServer("InfoPurview") = 1 Then Response.Write " checked"
    Response.Write ">���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'"
    If rsDownServer("InfoPurview") = 2 Then Response.Write " checked"
    Response.Write ">ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup(rsDownServer("arrGroupID"), "")
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "�Ķ�������</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & rsDownServer("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'" & ">&nbsp;&nbsp;&nbsp;&nbsp; <font color='#0000FF'>�������0�����Ա�Ķ���" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷��鿴��" & ChannelShortName & "��</font></td>"
    Response.Write "          </tr>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ظ��շѣ�</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0'"
    If rsDownServer("ChargeType") = 0 Then Response.Write " checked"
    Response.Write ">���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'"
    If rsDownServer("ChargeType") = 1 Then Response.Write " checked"
    Response.Write ">�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='" & rsDownServer("PitchTime") & "' size='8' maxlength='8' style='text-align:center'" & "> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'"
    If rsDownServer("ChargeType") = 2 Then Response.Write " checked"
    Response.Write ">��Ա�ظ��鿴������ <input name='ReadTimes' type='text' value='" & rsDownServer("ReadTimes") & "' size='8' maxlength='8' style='text-align:center'" & "> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'"
    If rsDownServer("ChargeType") = 3 Then Response.Write " checked"
    Response.Write ">�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'"
    If rsDownServer("ChargeType") = 4 Then Response.Write " checked"
    Response.Write ">����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'"
    If rsDownServer("ChargeType") = 5 Then Response.Write " checked"
    Response.Write ">ÿ�Ķ�һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ֳɱ�����</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='" & rsDownServer("DividePercent") & "' size='5' maxlength='4' style='text-align:center'" & "> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰����������Ķ�����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "      </table>"
    Response.Write "      <input name='ServerID' type='hidden' id='ServerID' value='" & rsDownServer("ServerID") & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "      <input  type='submit' name='Submit' value=' �����޸Ľ�� '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "      <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'""  style='cursor:hand;'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    rsDownServer.Close
    Set rsDownServer = Nothing
End Sub

Sub Del()
    Dim ServerID, iOrderID
    Dim rs, sql
    ServerID = Trim(Request("ServerID"))
    If ServerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ���ķ�����ID��</li>"
        Exit Sub
    Else
        ServerID = PE_CLng(ServerID)
    End If
    If FoundErr = True Then
    Exit Sub
    End If
    sql = "select OrderID from PE_DownServer where ServerID=" & ServerID
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Open sql, Conn, 1, 3
    If rs.BOF Or rs.EOF Then
    FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ���Ĳ����ļ�¼�����ڻ��ѱ�ɾ����</li>"
        Exit Sub
    Else
        iOrderID = rs("OrderID")
    End If
    'ɾ�����ش�����Ϣ��PE_DownError�����ڸþ���������ı�����Ϣ
    Dim rsDownError, sqlDownError
    Dim UrlID
    Set rsDownError = Server.CreateObject("ADODB.Recordset")
    sqlDownError = "select D.ErrorID,S.DownloadUrl from PE_DownError D left join PE_Soft S on D.InfoID=S.SoftID where D.UrlID=" & ServerID
    rsDownError.Open sqlDownError, Conn, 1, 3
    Do While Not rsDownError.EOF
        If InStr(rsDownError("DownloadUrl"), "@@@") > 0 Then
            Conn.Execute ("delete from PE_DownError where ErrorID =" & rsDownError("ErrorID"))
        End If
        rsDownError.MoveNext
    Loop
    rsDownError.Close
    Set rsDownError = Nothing
    Conn.Execute ("update PE_DownServer set OrderID=OrderID-1 where OrderID>" & iOrderID)
    Conn.Execute ("delete from PE_DownServer where ServerID=" & ServerID)
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?ChannelID=" & ChannelID & ""
End Sub

Sub Order()
    Dim iCount, i, j
    Dim rs, sql
    Set rs = Server.CreateObject("Adodb.RecordSet")
    sql = "select * from PE_DownServer where ChannelID=" & ChannelID & " Order by OrderID"
    rs.Open sql, Conn, 1, 1
    iCount = rs.RecordCount
    j = 1
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4' align='center'><strong>����������</strong></td>"
    Response.Write "  </tr>"
    Do While Not rs.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""> "
        Response.Write "    <td align='center'>" & rs("ServerName") & "</td>"
        Response.Write "    <form action='Admin_DownServer.asp?Action=UpOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If j > 1 Then
            Response.Write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=ServerID value=" & rs("ServerID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rs("OrderID") & ">&nbsp;<input type=submit name=Submit value='�޸�'>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form>"
        Response.Write "    <form action='Admin_DownServer.asp?Action=DownOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If iCount > j Then
            Response.Write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=ServerID value=" & rs("ServerID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rs("OrderID") & ">&nbsp;<input type=submit name=Submit value='�޸�'>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "      <td width='200' align='center'>&nbsp;</td>"
        Response.Write "    </form>"
        Response.Write "  </tr>"
        j = j + 1
        rs.MoveNext
    Loop
    Response.Write "</table> "
    rs.Close
    Set rs = Nothing
End Sub


Sub UpOrder()
    Dim ServerID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rs
    ServerID = Trim(Request("ServerID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ServerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    Else
        ServerID = PE_CLng(ServerID)
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
    Set mrs = Conn.Execute("select max(OrderID) from PE_DownServer")
    MaxOrderID = mrs(0) + 1
    '�Ƚ���ǰ�������������
    Conn.Execute ("update PE_DownServer set OrderID=" & MaxOrderID & " where ServerID=" & ServerID)
    
    'Ȼ��λ�ڵ�ǰ���������ϵķ�������OrderID���μ�һ����ΧΪҪ����������
    sqlOrder = "select * from PE_DownServer where OrderID<" & cOrderID & " order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '�����ǰ�������Ѿ��������棬�������ƶ�
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '�õ�Ҫ����λ�õ�OrderID
        Conn.Execute ("update PE_DownServer set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    'Ȼ���ٽ���ǰ������������Ƶ���Ӧλ��
    Conn.Execute ("update PE_DownServer set OrderID=" & tOrderID & " where ServerID=" & ServerID)
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?Action=Order&ChannelID=" & ChannelID
End Sub

Sub DownOrder()
    Dim ServerID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rs, PrevID, NextID
    ServerID = Trim(Request("ServerID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ServerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    Else
        ServerID = PE_CLng(ServerID)
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
    Set mrs = Conn.Execute("select max(OrderID) from PE_DownServer")
    MaxOrderID = mrs(0) + 1
    '�Ƚ���ǰ�������������
    Conn.Execute ("update PE_DownServer set OrderID=" & MaxOrderID & " where ServerID=" & ServerID)
    
    'Ȼ��λ�ڵ�ǰ���������µ�ǰ��������OrderID���μ�һ����ΧΪҪ�½�������
    sqlOrder = "select * from PE_DownServer where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '�����ǰ�������Ѿ��������棬�������ƶ�
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '�õ�Ҫ����λ�õ�OrderID
        Conn.Execute ("update PE_DownServer set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    'Ȼ���ٽ���ǰ������������Ƶ���Ӧλ��
    Conn.Execute ("update PE_DownServer set OrderID=" & tOrderID & " where ServerID=" & ServerID)
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?Action=Order&ChannelID=" & ChannelID
End Sub

'���������������ص�ַ��ʾ��ʽ(logo��������)
Sub SaveAllShowType()
    Dim rsDownServer, sqlDownServer
    Dim ShowType, ChannelID
    ShowType = Trim(Request("ShowType"))
    ChannelID = PE_CLng(Request("ChannelID"))

    If ShowType = "" Then
       ShowType = "False"
    End If

    sqlDownServer = "Select * from PE_DownServer where ChannelID=" & ChannelID
    Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
    rsDownServer.Open sqlDownServer, Conn, 1, 3
    If rsDownServer.BOF And rsDownServer.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ�����صķ�������Ϣ�������Ѿ���ɾ����</li>"
        rsDownServer.Close
        Set rsDownServer = Nothing
    Else
        Do While Not rsDownServer.EOF
            rsDownServer("ShowType") = ShowType
            rsDownServer.Update
            rsDownServer.MoveNext
        Loop
        rsDownServer.Close
        Set rsDownServer = Nothing
        Call CloseConn
        Response.Redirect "Admin_DownServer.asp?ChannelID=" & ChannelID & ""
    End If
End Sub
%>
