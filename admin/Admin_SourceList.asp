<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.SourceList.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = False   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim TypeSelect, Group, strTypeName, AllKeyList, AllUserList

MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
If MaxPerPage <= 0 Then MaxPerPage = 40
TypeSelect = ReplaceBadChar(Trim(Request("TypeSelect")))
Group = ReplaceBadChar(Trim(Request("Group")))
FileName = "Admin_SourceList.asp"
strFileName = "Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=" & Group & "&KeyWord=" & Keyword
XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>ѡ��Ի���</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<form method='post' name='myform' action=''>" & vbCrLf
Select Case TypeSelect
Case "KeyList"
    strTypeName = "�ؼ���"
    Call Key
Case "UserList"
    strTypeName = "��Ա"
    Call UserList
Case "AdminList"
    strTypeName = "����Ա"
    Call AdminList
Case "AgentList"
    strTypeName = "������"
    Call AgentList
Case "ClientList"
    strTypeName = "�ͻ�"
    Call ClientList
Case "CompanyList"
    strTypeName = "��ҵ"
    Call CompanyList
Case "ContacterList"
    strTypeName = "��ϵ��"
    Call ContacterList
Case "AuthorList"
    strTypeName = "����"
    Call Author
Case "CopyFromList"
    strTypeName = "��Դ"
    Call CopyFrom
Case "ProducerList"
    strTypeName = "����"
    Call Producer
Case "TrademarkList"
    strTypeName = "Ʒ��"
    Call Trademark
Case Else
    Response.Write "������ʧ"
End Select
Response.Write "</form>"
Response.Write "</body></html>"
Call CloseConn


Sub AdminList()
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title' height='22'>" & vbCrLf
    Response.Write "    <td valign='top'><b>�Ѿ�ѡ���Ĺ���Ա��</b></td>" & vbCrLf
    Response.Write "    <td align='right'><a href='javascript:window.returnValue=myform.UserList.value;window.close();'>����&gt;&gt;</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td><input type='text' name='UserList' size='60' maxlength='200' readonly='readonly'></td>" & vbCrLf
    Response.Write "    <td align='center'><input type='button' name='del1' onclick='del(1)' value='ɾ�����'> <input type='button' name='del2' onclick='del(0)' value='ɾ��ȫ��'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>�б�</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='����'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan=2>"
    Dim i, rsAdmin, sql
    sql = "select AdminName from PE_Admin Where 1=1"
    If Keyword <> "" Then
        sql = sql & " and AdminName like '%" & Keyword & "%'"
    End If
    sql = sql & " order by ID"
    
    Set rsAdmin = Server.CreateObject("adodb.recordset")
    rsAdmin.Open sql, Conn, 1, 1
    If rsAdmin.BOF And rsAdmin.EOF Then
        totalPut = 0
        Response.Write "<li>û���κι���Ա</li>"
    Else
        totalPut = rsAdmin.RecordCount
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
            rsAdmin.Move (CurrentPage - 1) * MaxPerPage
        Else
                    CurrentPage = 1
                End If
        End If
        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'><tr>"
        Do While Not rsAdmin.EOF
            If AllUserList = "" Then
                AllUserList = rsAdmin("AdminName")
            Else
                AllUserList = AllUserList & "," & rsAdmin("AdminName")
            End If
            Response.Write "<td align='center'><a href='#' onclick='add(""" & rsAdmin("AdminName") & """)'>" & rsAdmin("AdminName") & "</a></td>"
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            If (i Mod 8) = 0 And i > 1 Then Response.Write "</tr><tr>"
            rsAdmin.MoveNext
        Loop
        Response.Write "</tr></table>"
    End If
    rsAdmin.Close
    Set rsAdmin = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td align='center' colspan=2><a href='#' onclick='add(""" & AllUserList & """)'>�����������й���Ա</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������Ա", True)
    Call ShowJS("����Ա")
End Sub


Sub ClientList()
    Dim arrGroupID, arrClientType, ClientType
    ClientType = Trim(Request("ClientType"))
    arrGroupID = GetArrFromDictionary("PE_Client", "GroupID")
    arrClientType = Array("��ҵ�ͻ�", "���˿ͻ�")

    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>�б�</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & "><input type='hidden' name='ClientType' value='" & ClientType & "'>&nbsp;&nbsp;<input type='submit' value='����'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' height='22' align='center'>"
    Response.Write "          <td width='30'>ID</td>"
    Response.Write "          <td width='100'>��������</td>"
    Response.Write "          <td>�ͻ�����</td>"
    Response.Write "          <td width='60'>�ͻ����</td>"
    Response.Write "          <td width='80'>����</td>"
    Response.Write "        </tr>"


    Dim i, rsClient, sql
    sql = "select * from PE_Client Where 1=1"
    If PE_CLng(Group) > 0 Then
        sql = sql & " and GroupID=" & PE_CLng(Group)
    End If
    If ClientType = "E" Then
        sql = sql & " and ClientType=0"
    ElseIf ClientType = "P" Then
        sql = sql & " and ClientType=1"
    End If
    If Keyword <> "" Then
        sql = sql & " and ClientName like '%" & Keyword & "%'"
    End If
    sql = sql & " order by ClientID desc"
    
    Set rsClient = Server.CreateObject("adodb.recordset")
    rsClient.Open sql, Conn, 1, 1
    If rsClient.BOF And rsClient.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='10'>û���κοͻ�</td></tr>"
    Else
        totalPut = rsClient.RecordCount
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
            rsClient.Move (CurrentPage - 1) * MaxPerPage
        Else
                    CurrentPage = 1
                End If
        End If
        Do While Not rsClient.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'>" & rsClient("ClientID") & "</td>"
            Response.Write "        <td width='100'><a href='#' onclick=""window.returnValue='" & rsClient("ClientName") & "$$$" & rsClient("ClientID") & "';window.close();"">" & rsClient("ShortedForm") & "</a></td>"
            Response.Write "        <td><a href='#' onclick=""window.returnValue='" & rsClient("ClientName") & "$$$" & rsClient("ClientID") & "';window.close();"">" & rsClient("ClientName") & "</a></td>"
            Response.Write "        <td width='60' align='center'>" & GetArrItem(arrClientType, rsClient("ClientType")) & "</td>"
            Response.Write "        <td width='80' align='center'>" & GetArrItem(arrGroupID, rsClient("GroupID")) & "</td>"
            Response.Write "      </tr>"

            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsClient.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsClient.Close
    Set rsClient = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    strFileName = strFileName & "&ClientType=" & ClientType
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "���ͻ�", True)
End Sub

Sub CompanyList()
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>�б�</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='����'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' height='22' align='center'>"
    Response.Write "          <td width='30'>ID</td>"
    Response.Write "          <td>��λ����</td>"
    Response.Write "          <td>��ϵ��ַ</td>"
    Response.Write "          <td width='60'>��������</td>"
    Response.Write "        </tr>"

    Dim i, rsCompany, sql
    sql = "select * from PE_Company Where 1=1"
    If Keyword <> "" Then
        sql = sql & " and CompanyName like '%" & Keyword & "%'"
    End If
    sql = sql & " order by CompanyID desc"
    
    Set rsCompany = Server.CreateObject("adodb.recordset")
    rsCompany.Open sql, Conn, 1, 1
    If rsCompany.BOF And rsCompany.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='10'>û���κ���ҵ</td></tr>"
    Else
        totalPut = rsCompany.RecordCount
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
            rsCompany.Move (CurrentPage - 1) * MaxPerPage
        Else
                    CurrentPage = 1
                End If
        End If
        Do While Not rsCompany.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'>" & rsCompany("CompanyID") & "</td>"
            Response.Write "        <td><a href='#' onclick=""window.returnValue='" & rsCompany("CompanyName") & "$$$" & rsCompany("CompanyID") & "';window.close();"">" & rsCompany("CompanyName") & "</a></td>"
            Response.Write "        <td><a href='#' onclick=""window.returnValue='" & rsCompany("CompanyName") & "$$$" & rsCompany("CompanyID") & "';window.close();"">" & rsCompany("Address") & "</a></td>"
            Response.Write "        <td width='60' align='center'>" & rsCompany("ZipCode") & "</td>"
            Response.Write "      </tr>"

            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsCompany.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsCompany.Close
    Set rsCompany = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����ҵ", True)
End Sub

Sub ContacterList()
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "function add(str1,id1)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  opener.myform.ContacterName.value=str1;" & vbCrLf
    Response.Write "  opener.myform.ContacterID.value=id1;" & vbCrLf
    Response.Write "  window.close();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>�б�</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='����'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' height='22' align='center'>"
    Response.Write "          <td width='100'><strong>�� ��</strong></td>"
    Response.Write "          <td width='60'><strong>��ν</strong></td>"
    Response.Write "          <td width='100'><strong>�����绰</strong></td>"
    Response.Write "          <td width='100'><strong>�ֻ�</strong></td>"
    Response.Write "          <td width='90'><strong>����</strong></td>"
    Response.Write "          <td width='100'><strong>��Ӧ�ͻ�</strong></td>"
    Response.Write "        </tr>"


    Dim i, rsContacterList, sql
    sql = "select C.ContacterID,C.TrueName,C.Title,C.OfficePhone,C.Mobile,C.UserType,C.ClientID,Cl.ShortedForm from PE_Contacter C left join PE_Client Cl on C.ClientID=Cl.ClientID where UserType>0"
    If Keyword <> "" Then
        sql = sql & " and TrueName like '%" & Keyword & "%'"
    End If
    sql = sql & " order by ContacterID desc"
    
    Set rsContacterList = Server.CreateObject("adodb.recordset")
    rsContacterList.Open sql, Conn, 1, 1
    If rsContacterList.BOF And rsContacterList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='10'>û���κ���ϵ��</td></tr>"
    Else
        totalPut = rsContacterList.RecordCount
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
            rsContacterList.Move (CurrentPage - 1) * MaxPerPage
        Else
                    CurrentPage = 1
                End If
        End If
        Do While Not rsContacterList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td align='center'><a href='#' onclick=""window.returnValue='" & rsContacterList("TrueName") & "$$$" & rsContacterList("ContacterID") & "';window.close();"">" & rsContacterList("TrueName") & "</a></td>"
            Response.Write "        <td width='60' align='center'><a href='#' onclick=""window.returnValue='" & rsContacterList("TrueName") & "$$$" & rsContacterList("ContacterID") & "';window.close();"">" & rsContacterList("Title") & "</a></td>"
            Response.Write "        <td width='100' align='center'>" & rsContacterList("OfficePhone") & "</td>"
            Response.Write "        <td width='100' align='center'>" & rsContacterList("Mobile") & "</td>"
            Response.Write "        <td width='90' align='center'>" & GetArrItem(Array("���˿ͻ�", "����ϵ��", "������ϵ��"), rsContacterList("UserType")) & "</td>"
            Response.Write "        <td width='100' align='center'>" & rsContacterList("ShortedForm") & "</td>"
            Response.Write "      </tr>"

            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsContacterList.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsContacterList.Close
    Set rsContacterList = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����ϵ��", True)
End Sub

Sub Producer()
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    opener.myform.ProducerName.value=obj;" & vbCrLf
    Response.Write "    window.close();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' class='title'><td>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Time'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Time" Then Response.Write " color='red'>"
    Response.Write "�������</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=All'><FONT style='font-size:12px'" & vbCrLf
    If Group = "All" Then Response.Write " color='red'"
    Response.Write ">ȫ������</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=MLand'><FONT style='font-size:12px'" & vbCrLf
    If Group = "MLand" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType1", "��½����") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=GT'><FONT style='font-size:12px'" & vbCrLf
    If Group = "GT" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType2", "��̨����") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=JK'><FONT style='font-size:12px'" & vbCrLf
    If Group = "JK" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType3", "�պ�����") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=OutSea'><FONT style='font-size:12px'" & vbCrLf
    If Group = "OutSea" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType4", "ŷ������") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Other'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Other" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowProducer/ProducerType5", "��������") & "</FONT></a>" & vbCrLf
    Response.Write "         | </td></tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>�б�</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='����'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    
    Dim i, rsProducer, sql
    Select Case Group
    Case "Time"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",LastUseTime Desc"
    Case "All"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = (sql & " and ProducerName like '%" & Keyword & "%'")
        sql = (sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc")
    Case "MLand"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0) and ProducerType=1"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc"
    Case "GT"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0) and ProducerType=2"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc"
    Case "JK"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0) and ProducerType=3"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc"
    Case "OutSea"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0) and ProducerType=4"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc"
    Case "Other"
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0) and ProducerType=0"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc"
    Case Else
        sql = "select * from PE_Producer Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and ProducerName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ProducerID Desc"
    End Select
    Set rsProducer = Server.CreateObject("adodb.recordset")
    rsProducer.Open sql, Conn, 1, 1
    If rsProducer.BOF And rsProducer.EOF Then
        totalPut = 0
        Response.Write "<li>û�г���</li>"
    Else
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

        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>"
        Response.Write "<tr align='center'><td width='100' >����</td><td width='100'>��д</td><td>���</td></tr>"
        Do While Not rsProducer.EOF
            If AllKeyList = "" Then
                AllKeyList = rsProducer("ProducerName")
            Else
                AllKeyList = AllKeyList & "|" & rsProducer("ProducerName")
            End If
            Response.Write "<tr><td align='center'><a href='#' onclick='add(""" & rsProducer("ProducerName") & """)'>" & rsProducer("ProducerName") & "</a></td><td>" & rsProducer("ProducerShortName") & "</td>"
            If IsNull(rsProducer("ProducerIntro")) Then
                Response.Write "<td>��</td></tr>"
            Else
                Response.Write "<td>" & Left(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), 50) & "</td></tr>"
            End If
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsProducer.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsProducer.Close
    Set rsProducer = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��������", True)

End Sub

Sub Trademark()
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    opener.myform.TrademarkName.value=obj;" & vbCrLf
    Response.Write "    window.close();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' class='title'><td>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Time'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Time" Then Response.Write " color='red'"
    Response.Write ">�������</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=All'><FONT style='font-size:12px'" & vbCrLf
    If Group = "All" Then Response.Write " color='red'"
    Response.Write ">ȫ��Ʒ��</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=MLand'><FONT style='font-size:12px'" & vbCrLf
    If Group = "MLand" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType1", "��½Ʒ��") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=GT'><FONT style='font-size:12px'" & vbCrLf
    If Group = "GT" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType2", "��̨Ʒ��") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=JK'><FONT style='font-size:12px'" & vbCrLf
    If Group = "JK" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType3", "�պ�Ʒ��") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=OutSea'><FONT style='font-size:12px'" & vbCrLf
    If Group = "OutSea" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType4", "ŷ��Ʒ��") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='Admin_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Other'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Other" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowTrademark/TrademarkType5", "����Ʒ��") & "</FONT></a>" & vbCrLf
    Response.Write "         | </td></tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>�б�</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='����'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    
    Dim i, rsTrademark, sql
    Select Case Group
    Case "Time"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case "All"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case "MLand"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0) and TrademarkType=1"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case "GT"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0) and TrademarkType=2"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case "JK"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0) and TrademarkType=3"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case "OutSea"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0) and TrademarkType=4"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case "Other"
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0) and TrademarkType=0"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    Case Else
        sql = "select * from PE_Trademark Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and TrademarkName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",TrademarkID Desc"
    End Select
    Set rsTrademark = Server.CreateObject("adodb.recordset")
    rsTrademark.Open sql, Conn, 1, 1
    If rsTrademark.BOF And rsTrademark.EOF Then
        totalPut = 0
        Response.Write "<li>û��Ʒ��</li>"
    Else
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

        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>"
        Response.Write "<tr align='center'><td width='100' >����</td><td width='100'>ͼƬ</td><td>���</td></tr>"
        Do While Not rsTrademark.EOF
            If AllKeyList = "" Then
                AllKeyList = rsTrademark("TrademarkName")
            Else
                AllKeyList = AllKeyList & "|" & rsTrademark("TrademarkName")
            End If
            Response.Write "<tr><td align='center'><a href='#' onclick='add(""" & rsTrademark("TrademarkName") & """)'>" & rsTrademark("TrademarkName") & "</a></td>"
            If Not IsNull(rsTrademark("TrademarkPhoto")) Then
                Response.Write "<td align=center><img src='" & rsTrademark("TrademarkPhoto") & "' width='60' height='23'></td>"
            Else
                Response.Write "<td>&nbsp;</td>"
            End If
            If IsNull(rsTrademark("TrademarkIntro")) Then
                Response.Write "<td>��</td></tr>"
            Else
                Response.Write "<td>" & Left(nohtml(PE_HtmlDecode(rsTrademark("TrademarkIntro"))), 50) & "</td></tr>"
            End If
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsTrademark.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsTrademark.Close
    Set rsTrademark = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "���̱�", True)

End Sub
%>
