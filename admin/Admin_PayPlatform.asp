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

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>����֧��ƽ̨����</title>" & vbCrLf
Response.Write "<link href='Admin_STYLE.CSS' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� �� ֧ �� ƽ ̨ �� ��", 10212)
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td width='70' height='30' class='tdbg'>��������</td>" & vbCrLf
Response.Write "    <td class='tdbg'><a href='Admin_PayPlatform.asp'>����֧��ƽ̨����</a> | <a href='Admin_PayPlatform.asp?ManageType=Order'>����֧��ƽ̨����</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Select Case Action
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveBank
Case "Disable", "Enable"
    Call DisableBank
Case "SetDefault"
    Call SetDefault
Case "Order"
    Call Order
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "����֧��ƽ̨�������ʧ�ܣ�ʧ��ԭ��" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim ManageType
    ManageType = Trim(Request("ManageType"))
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td width='30'>ID</td>" & vbCrLf
    Response.Write "    <td width='80'>ƽ̨����</td>" & vbCrLf
    Response.Write "    <td width='120'>�̻�ID</td>" & vbCrLf
    Response.Write "    <td>ƽ̨˵��</td>" & vbCrLf
    Response.Write "    <td width='60'>��������</td>" & vbCrLf
    Response.Write "    <td width='50'>�Ƿ�Ĭ��</td>" & vbCrLf
    Response.Write "    <td width='40'>������</td>" & vbCrLf
    If ManageType <> "Order" Then
    Response.Write "    <td width='100'>�������</td>" & vbCrLf
    Else
    Response.Write "    <td width='100'>�������</td>" & vbCrLf
    End If
    Response.Write "  </tr>" & vbCrLf
    Dim rsPayPlatform, PayPlatformUrl
    Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform order by OrderID asc")
    If rsPayPlatform.BOF And rsPayPlatform.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>û���κ�����֧��ƽ̨</td></tr>"
    Else
        Do While Not rsPayPlatform.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbg2'"">" & vbCrLf
            Response.Write "    <td width='30' align='center'>" & rsPayPlatform("PlatformID") & "</td>" & vbCrLf
            Response.Write "    <td width='80' align='center'>" & rsPayPlatform("PlatformName") & "</td>" & vbCrLf
            Response.Write "    <td width='120' align='left'>" & rsPayPlatform("AccountsID") & "</td>" & vbCrLf
            Response.Write "    <td align='left' style=word-wrap:break-word;'>" & rsPayPlatform("Description") & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'>" & rsPayPlatform("Rate") & "%</td>" & vbCrLf
            Response.Write "    <td width='50' align='center'>" & vbCrLf
            If rsPayPlatform("IsDefault") = True Then
                Response.Write "��"
            End If
            Response.Write "      </td>" & vbCrLf
            Response.Write "    <td width='40' align='center'>" & vbCrLf
            If rsPayPlatform("IsDisabled") = False Then
                Response.Write "��"
            Else
                Response.Write "<font color='red'>��</font>"
            End If
            Response.Write "</td>" & vbCrLf
            If ManageType <> "Order" Then
                Response.Write "    <td width='100' align='center'>" & vbCrLf
                Select Case rsPayPlatform("PlatformID")
                Case 1
                    PayPlatformUrl = "http://merchant3.chinabank.com.cn/register.do"
                Case 2
                    PayPlatformUrl = "http://www.ipay.cn"
                Case 3
                    PayPlatformUrl = "https://www.ips.com.cn"
                Case 4
                    PayPlatformUrl = "#"
                Case 5
                    PayPlatformUrl = "http://www.yeepay.com/"
                Case 6
                    PayPlatformUrl = "http://new.xpay.cn/SignUp/Default.aspx"
                Case 7
                    PayPlatformUrl = "https://www.cncard.net"
                Case 8
                    PayPlatformUrl = "https://www.alipay.com/"
                Case 9
                    PayPlatformUrl = "http://www.99bill.com/"
                Case 10
                    PayPlatformUrl = "#"
                Case 11
                    PayPlatformUrl = "http://www.99bill.com/"
                Case 12
                    PayPlatformUrl = "https://www.alipay.com/"
                Case 13
                    PayPlatformUrl = "http://union.tenpay.com/mch/mch_register.shtml?posid=123&actid=84&opid=50&whoid=31&sp_suggestuser=1201648901"
                End Select
                Response.Write "<a href='Admin_PayPlatform.asp?Action=Modify&PlatformID=" & rsPayPlatform("PlatformID") & "'>�޸�</a> "
                Response.Write "<a href='" & PayPlatformUrl & "' target='_blank'>�����̻�</a><br>"
                If rsPayPlatform("IsDisabled") = True Then
                    Response.Write "<a href='Admin_PayPlatform.asp?Action=Enable&PlatformID=" & rsPayPlatform("PlatformID") & "'>����</a> "
                Else
                    If rsPayPlatform("IsDefault") = True Then
                        Response.Write "<font color='gray'>����</font> "
                    Else
                        Response.Write "<a href='Admin_PayPlatform.asp?Action=Disable&PlatformID=" & rsPayPlatform("PlatformID") & "'>����</a> "
                    End If
                End If
                If rsPayPlatform("IsDisabled") = True Or rsPayPlatform("IsDefault") = True Then
                    Response.Write "<font color='gray'>��ΪĬ��</font> <br>"
                Else
                    Response.Write "<a href='Admin_PayPlatform.asp?Action=SetDefault&PlatformID=" & rsPayPlatform("PlatformID") & "'>��ΪĬ��</a>"
                End If
                Response.Write "</td>"
            Else
                Response.Write "<form name='orderform' method='post' action='Admin_PayPlatform.asp'>"
                Response.Write "    <td width='100' align='center'><input name='OrderID' type='text' id='OrderID' value='" & rsPayPlatform("OrderID") & "' size='4' maxlength='4' style='text-align:center '><input type='submit' name='Submit' value='�޸�'><input name='PlatformID' type='hidden' id='PlatformID' value='" & rsPayPlatform("PlatformID") & "'><input name='Action' type='hidden' id='Action' value='Order'></td></form>"
            End If
            Response.Write "  </tr>"
            rsPayPlatform.MoveNext
        Loop
    End If
    Set rsPayPlatform = Nothing
    Response.Write "</table>"
    Response.Write "<br>" & vbCrLf
    Response.Write "˵���������á�ĳ����֧��ƽ̨������֧��ʱ��������ʾ������֧��ƽ̨����������֧����¼�������Ի���ʾ��<br>" & vbCrLf
End Sub

Sub Modify()
    Dim PlatformID, rsPayPlatform
    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������֧��ƽ̨ID</li>"
        Exit Sub
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform where PlatformID=" & PlatformID & "")
    If rsPayPlatform.BOF And rsPayPlatform.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ��������֧��ƽ̨��</li>"
        Set rsPayPlatform = Nothing
        Exit Sub
    End If
    Response.Write "<form name='myform' method='post' action='Admin_PayPlatform.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><b>�� �� �� �� ֧ �� ƽ ̨</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>ƽ̨���ƣ�</td>" & vbCrLf
    Response.Write "      <td><input name='PlatformName' type='text' id='PlatformName' size='50' maxlength='20' value='" & rsPayPlatform("PlatformName") & "' disabled></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>������ʾ�����ƣ�</td>" & vbCrLf
    Response.Write "      <td><input name='ShowName' type='text' id='ShowName' size='50' maxlength='30' value='" & rsPayPlatform("ShowName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>˵����</td>" & vbCrLf
    Response.Write "      <td><textarea name='Description' cols='42' rows='5'>" & rsPayPlatform("Description") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�̻�ID</td>" & vbCrLf
    Response.Write "      <td><input name='AccountsID' type='text' id='AccountsID' size='50' maxlength='50' value='" & rsPayPlatform("AccountsID") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>MD5��Կ��</td>" & vbCrLf
    Response.Write "      <td><input name='MD5Key' type='password' id='MD5Key' size='50' maxlength='255' value='" & rsPayPlatform("MD5Key") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>�������ʣ�</td>" & vbCrLf
    Response.Write "      <td><input name='Rate' type='text' id='Rate' size='5' maxlength='5' value='" & rsPayPlatform("Rate") & "'>% <font color='#FF0000'>*</font><br>" & vbCrLf
    Response.Write "        <input name='PlusPoundage' type='checkbox' value='1' " & IsRadioChecked(rsPayPlatform("PlusPoundage"), True) & "> �������ɸ����˶���֧��</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='50' colspan='2'><input name='PlatformID' type='hidden' id='PlatformID' value='" & PlatformID & "'>" & vbCrLf
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>" & vbCrLf
    Response.Write "          <input type='submit' name='Submit' value='��������֧��ƽ̨'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "" & vbCrLf
    Set rsPayPlatform = Nothing
End Sub

Sub SaveBank()
    Dim PlatformID, ShowName, AccountsID, Accounts, MD5Key, Rate
    Dim rsPayPlatform, sqlPlatform
    PlatformID = Trim(Request("PlatformID"))
    ShowName = Trim(Request("ShowName"))
    AccountsID = Trim(Request("AccountsID"))
    Accounts = Trim(Request("Accounts"))
    MD5Key = Trim(Request("MD5Key"))
    Rate = Trim(Request("Rate"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��PlatformID��</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    If ShowName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��ƽ̨������ʾ������</li>"
    End If
    If AccountsID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ���̻�ID</li>"
    End If
    If MD5Key = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��MD5��Կ</li>"
    End If
    If Rate = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ����������</li>"
    Else
        Rate = PE_CDbl(Rate)
    End If
    
    
    If FoundErr = True Then Exit Sub
    
    sqlPlatform = "select * from PE_PayPlatform where PlatformID=" & PlatformID
    Set rsPayPlatform = Server.CreateObject("adodb.recordset")
    rsPayPlatform.Open sqlPlatform, Conn, 1, 3
    
    rsPayPlatform("ShowName") = ShowName
    rsPayPlatform("Description") = Trim(Request("Description"))
    rsPayPlatform("AccountsID") = AccountsID
    rsPayPlatform("MD5Key") = MD5Key
    rsPayPlatform("Rate") = Rate
    rsPayPlatform("PlusPoundage") = PE_CBool(Trim(Request("PlusPoundage")))
    rsPayPlatform.Update
    rsPayPlatform.Close
    Set rsPayPlatform = Nothing
    Call WriteEntry(2, AdminName, "��������֧��ƽ̨��Ϣ�ɹ���" & AccountsID)
    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub

Sub DisableBank()
    Dim PlatformID
    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������֧��ƽ̨ID</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    If FoundErr = True Then Exit Sub
    Dim trs
    Set trs = Conn.Execute("select IsDefault from PE_PayPlatform where PlatformID=" & PlatformID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ��������֧��ƽ̨</li>"
    Else
        If trs(0) = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ܽ���Ĭ�ϵ�����֧��ƽ̨</li>"
        End If
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub
    
    Select Case Action
    Case "Disable"
        Conn.Execute ("update PE_PayPlatform set IsDisabled=" & PE_True & " where PlatformID=" & PlatformID & "")
    Case "Enable"
        Conn.Execute ("update PE_PayPlatform set IsDisabled=" & PE_False & " where PlatformID=" & PlatformID & "")
    End Select

    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub

Sub Order()
    Dim PlatformID, OrderID
    PlatformID = Trim(Request("PlatformID"))
    OrderID = Trim(Request("OrderID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������֧��ƽ̨ID</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    If OrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��OrderID</li>"
    Else
        OrderID = PE_CLng(OrderID)
    End If
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_PayPlatform set OrderID=" & OrderID & " where PlatformID=" & PlatformID & "")
    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub

Sub SetDefault()
    Dim PlatformID
    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��PlatformID</li>"
        Exit Sub
    Else
        PlatformID = PE_CLng(PlatformID)
    End If

    Conn.Execute ("update PE_PayPlatform set IsDefault=" & PE_False & "")
    Conn.Execute ("update PE_PayPlatform set IsDefault=" & PE_True & " where  PlatformID=" & PlatformID)
    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub
%>
