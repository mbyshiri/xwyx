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
Const PurviewLevel_Others = "Bank"   '����Ȩ��
    
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>�����ʻ�����</title>" & vbCrLf
Response.Write "<link href='Admin_STYLE.CSS' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� �� �� �� �� ��", 10212)
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td width='70' height='30' class='tdbg'>��������</td>" & vbCrLf
Response.Write "    <td class='tdbg'><a href='Admin_Bank.asp'>�����ʻ�������ҳ</a> | <a href='Admin_Bank.asp?Action=Add'>��������ʻ�</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveBank
Case "Disable", "Enable"
    Call DisableBank
Case "SetDefault"
    Call SetDefault
Case "Del"
    Call Del
Case "Order"
    Call Order
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td width='30'>ID</td>" & vbCrLf
    Response.Write "    <td width='60'>�ʻ�����</td>" & vbCrLf
    Response.Write "    <td width='80'>������</td>" & vbCrLf
    Response.Write "    <td width='100'>����</td>" & vbCrLf
    Response.Write "    <td>�ʺ�/����</td>" & vbCrLf
    Response.Write "    <td width='50'>�Ƿ�Ĭ��</td>" & vbCrLf
    Response.Write "    <td width='40'>������</td>" & vbCrLf
    Response.Write "    <td width='150'>�������</td>" & vbCrLf
    Response.Write "    <td width='80'>�������</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Dim rsBank
    Set rsBank = Conn.Execute("select * from PE_Bank order by OrderID asc")
    If rsBank.BOF And rsBank.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>û���κ������ʻ�</td></tr>"
    Else
        Do While Not rsBank.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbg2'"">" & vbCrLf
            Response.Write "    <td width='30' align='center'>" & rsBank("BankID") & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'>" & rsBank("BankShortName") & "</td>" & vbCrLf
            Response.Write "    <td width='80' align='left'>" & rsBank("BankName") & "</td>" & vbCrLf
            Response.Write "    <td width='100' align='center'>" & rsBank("HolderName") & "</td>" & vbCrLf
            Response.Write "    <td align='left'>�ʺţ�" & rsBank("Accounts") & "<br>���ţ�" & rsBank("CardNum") & "</td>" & vbCrLf
            Response.Write "    <td width='50' align='center'>" & vbCrLf
            If rsBank("IsDefault") = True Then
                Response.Write "��"
            End If
            Response.Write "      </td>" & vbCrLf
            Response.Write "    <td width='40' align='center'>" & vbCrLf
            If rsBank("IsDisabled") = False Then
                Response.Write "��"
            Else
                Response.Write "<font color='red'>��</font>"
            End If
            Response.Write "</td>" & vbCrLf
            Response.Write "    <td width='150' align='center'>" & vbCrLf
            If rsBank("IsDefault") = True Then
                Response.Write "<font color='gray'>��ΪĬ�� ����</font> "
            Else
                Response.Write "<a href='Admin_Bank.asp?Action=SetDefault&BankID=" & rsBank("BankID") & "'>��ΪĬ��</a> "
                If rsBank("IsDisabled") = True Then
                    Response.Write "<a href='Admin_Bank.asp?Action=Enable&BankID=" & rsBank("BankID") & "'>����</a> "
                Else
                    Response.Write "<a href='Admin_Bank.asp?Action=Disable&BankID=" & rsBank("BankID") & "'>����</a> "
                End If
            End If
            Response.Write "<a href='Admin_Bank.asp?Action=Modify&BankID=" & rsBank("BankID") & "'>�޸�</a> "
            If rsBank("IsDefault") = True Then
                Response.Write "<font color='gray'>ɾ��</font> "
            Else
                Response.Write "<a href='Admin_Bank.asp?Action=Del&BankID=" & rsBank("BankID") & "'>ɾ��</a>"
            End If
            Response.Write "</td><form name='orderform' method='post' action='Admin_Bank.asp'>"
            Response.Write "    <td width='80' align='center'><input name='OrderID' type='text' id='OrderID' value='" & rsBank("OrderID") & "' size='4' maxlength='4' style='text-align:center '><input type='submit' name='Submit' value='�޸�'><input name='BankID' type='hidden' id='BankID' value='" & rsBank("BankID") & "'><input name='Action' type='hidden' id='Action' value='Order'></td></form>"
            Response.Write "  </tr>"
            rsBank.MoveNext
        Loop
    End If
    Set rsBank = Nothing
    Response.Write "</table>"
    Response.Write "<br>" & vbCrLf
    Response.Write "˵���������á�ĳ�����ʻ�����������Ϣʱ��������ʾ�������ʻ��������ʽ���ϸ������Ի���ʾ��<br>" & vbCrLf
End Sub

Sub Add()
    Response.Write "<form name='myform' method='post' action='Admin_Bank.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><b>�� �� �� �� �� ��</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�ʻ����ƣ�</td>" & vbCrLf
    Response.Write "      <td><input name='BankShortName' type='text' id='BankShortName' size='20' maxlength='20'> <font color='#FF0000'>*</font> ��������д��һ��¼��Ͳ����޸ġ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�����У�</td>" & vbCrLf
    Response.Write "      <td><input name='BankName' type='text' id='BankName' size='20' maxlength='50'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>������</td>" & vbCrLf
    Response.Write "      <td><input name='HolderName' type='text' id='HolderName' size='20' maxlength='20'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�ʺţ�</td>" & vbCrLf
    Response.Write "      <td><input name='Accounts' type='text' id='Accounts' size='20' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>���ţ�</td>" & vbCrLf
    Response.Write "      <td><input name='CardNum' type='text' id='CardNum' size='20' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>����ͼ�꣺</td>" & vbCrLf
    Response.Write "      <td><input name='BankPic' type='text' id='BankPic' size='40' maxlength='200'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�ʻ�˵����</td>" & vbCrLf
    Response.Write "      <td><textarea name='BankIntro' cols='40' rows='3'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='right'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td><input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'> ��ΪĬ�������ʻ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='50' colspan='2'><input name='Action' type='hidden' id='Action' value='SaveAdd'><input type='submit' name='Submit' value='���������ʻ�'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Modify()
    Dim BankID, rsBank
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ�������ʻ�ID</li>"
        Exit Sub
    Else
        BankID = PE_CLng(BankID)
    End If
    Set rsBank = Conn.Execute("select * from PE_Bank where BankID=" & BankID & "")
    If rsBank.BOF And rsBank.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���������ʻ���</li>"
        Set rsBank = Nothing
        Exit Sub
    End If
    Response.Write "<form name='myform' method='post' action='Admin_Bank.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><b>�� �� �� �� �� ��</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�ʻ����ƣ�</td>" & vbCrLf
    Response.Write "      <td><input name='BankShortName' type='text' id='BankShortName' size='20' maxlength='20' value='" & rsBank("BankShortName") & "' disabled> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�����У�</td>" & vbCrLf
    Response.Write "      <td><input name='BankName' type='text' id='BankName' size='20' maxlength='50' value='" & rsBank("BankName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>������</td>" & vbCrLf
    Response.Write "      <td><input name='HolderName' type='text' id='HolderName' size='20' maxlength='20' value='" & rsBank("HolderName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�ʺţ�</td>" & vbCrLf
    Response.Write "      <td><input name='Accounts' type='text' id='Accounts' size='20' maxlength='30' value='" & rsBank("Accounts") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>���ţ�</td>" & vbCrLf
    Response.Write "      <td><input name='CardNum' type='text' id='CardNum' size='20' maxlength='30' value='" & rsBank("CardNum") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>����ͼ�꣺</td>" & vbCrLf
    Response.Write "      <td><input name='BankPic' type='text' id='BankPic' size='40' maxlength='200' value='" & rsBank("BankPic") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�ʻ�˵����</td>" & vbCrLf
    Response.Write "      <td><textarea name='BankIntro' cols='40' rows='3'>" & rsBank("BankIntro") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='right'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td><input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'"
    If rsBank("IsDefault") = True Then Response.Write " checked"
    Response.Write "> ��ΪĬ�������ʻ�</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='50' colspan='2'><input name='BankID' type='hidden' id='BankID' value='" & BankID & "'>" & vbCrLf
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>" & vbCrLf
    Response.Write "          <input type='submit' name='Submit' value='���������ʻ�'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "" & vbCrLf
    Set rsBank = Nothing
End Sub

Sub SaveBank()
    Dim BankID, BankShortName, BankName, Accounts, CardNum, HolderName, BankIntro, BankPic, IsDefault, OrderID
    Dim rsBank, sqlBank
    BankID = Trim(Request("BankID"))
    BankShortName = ReplaceBadChar(Trim(Request("BankShortName")))
    BankName = Trim(Request("BankName"))
    Accounts = Trim(Request("Accounts"))
    CardNum = Trim(Request("CardNum"))
    HolderName = Trim(Request("HolderName"))
    BankIntro = Trim(Request("BankIntro"))
    BankPic = Trim(Request("BankPic"))
    IsDefault = Trim(Request("IsDefault"))
    If Action = "SaveAdd" And BankShortName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ���ʻ�����</li>"
    End If
    If BankName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��������</li>"
    End If
    If HolderName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������</li>"
    End If
    If Accounts = "" And CardNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ҫ�����ʺźͿ��ŵ�����һ��</li>"
    End If
    If IsDefault = "Yes" Then
        IsDefault = True
        Conn.Execute ("update PE_Bank set IsDefault=" & PE_False & "")
    Else
        IsDefault = False
    End If
    
    If Action = "SaveAdd" Then
        sqlBank = "select top 1 * from PE_Bank"
        Dim trs, mrs
        Set trs = Conn.Execute("select BankID from PE_Bank where BankShortName='" & BankShortName & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ָ�����ʻ������Ѿ����ڣ�</li>"
        Else
            Set mrs = Conn.Execute("select max(BankID) from PE_Bank")
            If IsNull(mrs(0)) Then
                BankID = 1
            Else
                BankID = mrs(0) + 1
            End If
            Set mrs = Nothing
            Set mrs = Conn.Execute("select max(OrderID) from PE_Bank")
            If IsNull(mrs(0)) Then
                OrderID = 1
            Else
                OrderID = mrs(0) + 1
            End If
            Set mrs = Nothing
        End If
        Set trs = Nothing
    Else
        If BankID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��BankID��</li>"
        Else
            BankID = PE_CLng(BankID)
        End If
        sqlBank = "select * from PE_Bank where BankID=" & BankID
    End If
    
    If FoundErr = True Then Exit Sub
    
    Set rsBank = Server.CreateObject("adodb.recordset")
    rsBank.Open sqlBank, Conn, 1, 3
    If Action = "SaveAdd" Then
        rsBank.AddNew
        rsBank("BankID") = BankID
        rsBank("OrderID") = OrderID
        rsBank("IsDisabled") = False
        rsBank("BankShortName") = BankShortName
    End If
    
    rsBank("BankName") = BankName
    rsBank("Accounts") = Accounts
    rsBank("CardNum") = CardNum
    rsBank("HolderName") = HolderName
    rsBank("BankIntro") = BankIntro
    rsBank("BankPic") = BankPic
    rsBank("IsDefault") = IsDefault
    rsBank.Update
    rsBank.Close
    Set rsBank = Nothing
    Call WriteEntry(2, AdminName, "���������ʻ���Ϣ�ɹ���" & BankName)
    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub

Sub Del()
    Dim BankID
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ�������ʻ�ID</li>"
        Exit Sub
    Else
        BankID = PE_CLng(BankID)
    End If
    
    Dim rsBank, trs
    Set rsBank = Server.CreateObject("adodb.recordset")
    rsBank.Open "select * from PE_Bank where BankID=" & BankID & "", Conn, 1, 3
    If rsBank.BOF And rsBank.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���������ʻ�"
    Else
        If rsBank("IsDefault") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ɾ��Ĭ�ϵ������ʻ�</li>"
        End If
        Set trs = Conn.Execute("select top 1 ItemID from PE_BankrollItem where Bank='" & rsBank("BankShortName") & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�������ʻ��Ѿ����ʽ���ϸ��¼�����Բ���ɾ����������Խ��ô������ʻ����Դﵽǰ̨����ʾ�������ʻ���Ŀ�ġ�</li>"
        End If
        Set trs = Nothing
        If FoundErr = False Then
            rsBank.Delete
            rsBank.Update
        End If
    End If
    rsBank.Close
    Set rsBank = Nothing
    Call WriteEntry(2, AdminName, "ɾ�������ʻ��ɹ���ID��" & BankID)

    Call CloseConn
    If FoundErr = False Then
        Response.Redirect "Admin_Bank.asp"
    End If
End Sub

Sub DisableBank()
    Dim BankID
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ�������ʻ�ID</li>"
    Else
        BankID = PE_CLng(BankID)
    End If
    If FoundErr = True Then Exit Sub
    Dim trs
    Set trs = Conn.Execute("select IsDefault from PE_Bank where BankID=" & BankID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���������ʻ�</li>"
    Else
        If trs(0) = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ܽ���Ĭ�ϵ������ʻ�</li>"
        End If
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub
    
    Select Case Action
    Case "Disable"
        Conn.Execute ("update PE_Bank set IsDisabled=" & PE_True & " where BankID=" & BankID & "")
    Case "Enable"
        Conn.Execute ("update PE_Bank set IsDisabled=" & PE_False & " where BankID=" & BankID & "")
    End Select

    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub

Sub Order()
    Dim BankID, OrderID
    BankID = Trim(Request("BankID"))
    OrderID = Trim(Request("OrderID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ�������ʻ�ID</li>"
    Else
        BankID = PE_CLng(BankID)
    End If
    If OrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��OrderID</li>"
    Else
        OrderID = PE_CLng(OrderID)
    End If
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_Bank set OrderID=" & OrderID & " where BankID=" & BankID & "")
    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub

Sub SetDefault()
    Dim BankID
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��BankID</li>"
        Exit Sub
    Else
        BankID = PE_CLng(BankID)
    End If
    Conn.Execute ("update PE_Bank set IsDefault=" & PE_False & "")
    Conn.Execute ("update PE_Bank set IsDefault=" & PE_True & " where  BankID=" & BankID)
    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub
%>
