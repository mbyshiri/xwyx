<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Payment"   '����Ȩ��

strFileName = "Admin_Payment.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword


Response.Write "<html><head><title>����֧����¼����</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
Call ShowJS_Main("����֧����¼")
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
Call ShowPageTitle("�� �� ֧ �� �� ¼ �� ��", 10204)
Response.Write "    <tr class='tdbg' height='30'> "
Response.Write "  <form name='form1' action='Admin_Payment.asp' method='get'>"
Response.Write "      <td>���ٲ��ң�"
Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
Response.Write "          <option value='0'"
If SearchType = 0 Then Response.Write " selected"
Response.Write ">��������֧����¼</option>"
Response.Write "          <option value='1'"
If SearchType = 1 Then Response.Write " selected"
Response.Write ">���10���ڵ�������֧����¼</option>"
Response.Write "          <option value='2'"
If SearchType = 2 Then Response.Write " selected"
Response.Write ">���һ���ڵ�������֧����¼</option>"
Response.Write "          <option value='3'"
If SearchType = 3 Then Response.Write " selected"
Response.Write ">δ�ύ������֧����¼</option>"
Response.Write "          <option value='4'"
If SearchType = 4 Then Response.Write " selected"
Response.Write ">δ�ɹ�������֧����¼</option>"
Response.Write "          <option value='5'"
If SearchType = 5 Then Response.Write " selected"
Response.Write ">֧���ɹ�������֧����¼</option>"
Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_Payment.asp'>����֧����¼��ҳ</a></td>"
Response.Write "  </form>"
Response.Write "<form name='form2' method='post' action='Admin_Payment.asp'>"
Response.Write "    <td>�߼���ѯ��"
Response.Write "      <select name='Field' id='Field'>"
Response.Write "      <option value='PaymentNum'>����֧����¼���</option>"
Response.Write "      <option value='UserName'>�û���</option>"
Response.Write "      <option value='PayTime'>֧��ʱ��</option>"
Response.Write "      </select>"
Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='30'>"
Response.Write "      <input type='submit' name='Submit2' value=' �� ѯ '>"
Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'>"
Response.Write " </td>"
Response.Write "</form>"
Response.Write "</table>"
Response.Write "<br>"
If Action = "Cancel" Then
    Call DelPayment
ElseIf Action = "Success" Then
    Call PaySuccess
Else
    Call main
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsPaymentList, sqlPaymentList, Querysql
    Dim TotalMoneyPay, TotalMoneyTrue
    TotalMoneyPay = 0
    TotalMoneyTrue = 0
    
    sqlPaymentList = "select top " & MaxPerPage & " * from PE_Payment "
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>�����ڵ�λ�ã�<a href='Admin_Payment.asp'>����֧����¼����</a>&nbsp;&gt;&gt;&nbsp;"

    Querysql = Querysql & " where 1=1 "
    Select Case SearchType
        Case 0
            Response.Write "��������֧����¼"
        Case 1
            Querysql = Querysql & " and datediff(" & PE_DatePart_D & ",PayTime," & PE_Now & ")<10"
            Response.Write "���10���ڵ�������֧����¼"
        Case 2
            Querysql = Querysql & " and datediff(" & PE_DatePart_M & ",PayTime," & PE_Now & ")<1"
            Response.Write "���һ���ڵ�������֧����¼"
        Case 3
            Querysql = Querysql & " and Status=1"
            Response.Write "δ�ύ������֧����¼"
        Case 4
            Querysql = Querysql & " and Status=2"
            Response.Write "δ�ɹ�������֧����¼"
        Case 5
            Querysql = Querysql & " and Status=3"
            Response.Write "֧���ɹ�������֧����¼"
        Case 10
            If Keyword = "" Then
                Response.Write "��������֧����¼"
            Else
                Select Case strField
                Case "PaymentNum"
                    Querysql = Querysql & " and PaymentNum like '%" & Keyword & "%'"
                    Response.Write "����֧����¼����к��С� <font color=red> " & Keyword & " </font> ��������֧����¼"
                Case "UserName"
                    Querysql = Querysql & " and UserName like '%" & Keyword & "%'"
                    Response.Write "�û����к��С� <font color=red>" & Keyword & "</font> ��������֧����¼"
                Case "PayTime"
                    If IsDate(Keyword) = True Then
                        If SystemDatabaseType = "SQL" Then
                            Querysql = Querysql & " and PayTime='" & Keyword & "'"
                        Else
                            Querysql = Querysql & " and PayTime=#" & Keyword & "#"
                        End If
                        Response.Write "֧��ʱ��Ϊ <font color=red>" & Keyword & "</font> ������֧����¼"
                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>��ѯ��֧��ʱ���ʽ����ȷ��</li>"
                    End If
                End Select
            End If
        Case Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ĳ�����</li>"
    End Select

    totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Payment " & Querysql)(0))
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
        Querysql = Querysql & " and PaymentID < (select min(PaymentID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " PaymentID from PE_Payment " & Querysql & " order by PaymentID desc) as QueryPayment) "
    End If
    sqlPaymentList = sqlPaymentList & Querysql & " order by PaymentID desc"

    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Payment.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ��������֧����¼��');"">"
    Response.Write "     <td>"
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'>ѡ��</td>"
    Response.Write "    <td width='80'>֧�����</td>"
    Response.Write "    <td width='60'>�û���</td>"
    Response.Write "    <td width='70'>֧��ƽ̨</td>"
    Response.Write "    <td width='120'>����ʱ��</td>"
    Response.Write "    <td width='70'>�����</td>"
    Response.Write "    <td width='70'>ʵ��ת��<br>���</td>"
    Response.Write "    <td width='60'>����״̬</td>"
    Response.Write "    <td width='70'>������Ϣ</td>"
    Response.Write "    <td>��ע</td>"
    Response.Write "    <td>����</td>"
    Response.Write "  </tr>"
    
    Set rsPaymentList = Server.CreateObject("Adodb.RecordSet")
    rsPaymentList.Open sqlPaymentList, Conn, 1, 1
    If rsPaymentList.BOF And rsPaymentList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κη�������������֧������</td></tr>"
    Else
        Dim i
        i = 0
        Do While Not rsPaymentList.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='30' align='center'><input name='PaymentID' type='checkbox' onclick='unselectall()' id='PaymentID' value='" & rsPaymentList("PaymentID") & "'></td>"
            Response.Write "    <td width='80' align='center'>" & rsPaymentList("PaymentNum") & "</td>"
            Response.Write "    <td width='60' align='center'><a href='Admin_User.asp?Action=Show&UserName=" & rsPaymentList("UserName") & "'>" & rsPaymentList("UserName") & "</a></td>"
            Response.Write "    <td width='70' align='center'>" & GetPayOnlineProviderName(rsPaymentList("eBankID")) & "</td>"
            Response.Write "    <td width='120' align='center'>" & rsPaymentList("PayTime") & "</td>"
            Response.Write "    <td width='70' align='right'>" & FormatNumber(rsPaymentList("MoneyPay"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='70' align='right'>" & FormatNumber(rsPaymentList("MoneyTrue"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='60' align='center'>"
            If rsPaymentList("eBankID") <> 8 Then
                Select Case rsPaymentList("Status")
                Case 1
                    Response.Write "δ�ύ"
                Case 2
                    Response.Write "�Ѿ��ύ����δ�ɹ�"
                Case 3
                    Response.Write "֧���ɹ�"
                End Select
            Else
                Select Case rsPaymentList("Status")
                Case 1
                    Response.Write "�ȴ���Ҹ���"
                Case 2
                    Response.Write "����Ѹ���"
                Case 3
                    Response.Write "���׳ɹ�"
                Case 4
                    Response.Write "�����ѷ������ȴ����ȷ���ջ�"
                End Select
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='70' align='center'>" & rsPaymentList("eBankInfo") & "</td>"
            Response.Write "    <td>" & rsPaymentList("Remark") & "</td>"
            Response.Write "    <td align='center'>"
            If rsPaymentList("Status") = 1 Then
                Response.Write "<a href='Admin_Payment.asp?Action=Cancel&PaymentID=" & rsPaymentList("PaymentID") & "' onclick=""return confirm('ȷ��Ҫɾ����������֧����¼��');"">ȡ��</a> "
                Response.Write "<a href='Admin_Payment.asp?Action=Success&PaymentID=" & rsPaymentList("PaymentID") & "' onclick=""return confirm('ȷ����������֧����¼�Ѿ�֧���ɹ�����');"">�ɹ�</a>"
            End If
            Response.Write "</td>"
            Response.Write "  </tr>"
            TotalMoneyPay = TotalMoneyPay + rsPaymentList("MoneyPay")
            TotalMoneyTrue = TotalMoneyTrue + rsPaymentList("MoneyTrue")
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsPaymentList.MoveNext
        Loop
    End If
    rsPaymentList.Close
    Set rsPaymentList = Nothing
        
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='5' align='right'>�ϼƽ�</td>"
    Response.Write "    <td width='70' align='right'>" & FormatNumber(TotalMoneyPay, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td width='70' align='right'>" & FormatNumber(TotalMoneyTrue, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='4' align='center'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='220' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ����������֧����¼</td>"
    Response.Write "    <td width='560'> <input name='Action' type='hidden' id='Action' value='Cancel'> <input type='submit' name='Submit' value='ɾ��ѡ��������֧����¼'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������֧����¼", True)
End Sub


Sub DelPayment()
    Dim PaymentID
    Dim rsPayment, sqlPayment
    PaymentID = Trim(Request("PaymentID"))
    If PaymentID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��֧����ID��</li>"
        Exit Sub
    Else
        If IsValidID(PaymentID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����ȷ��֧����ID��</li>"
            Exit Sub
        End If
    End If
    
    sqlPayment = "select * from PE_Payment where PaymentID in (" & PaymentID & ")"
    Set rsPayment = Server.CreateObject("Adodb.RecordSet")
    rsPayment.Open sqlPayment, Conn, 1, 3
    Do While Not rsPayment.EOF
        If rsPayment("Status") = 1 Then
            rsPayment.Delete
            rsPayment.Update
        End If
        rsPayment.MoveNext
    Loop
    rsPayment.Close
    Set rsPayment = Nothing
    Call CloseConn
    Call WriteSuccessMsg("�ɹ�ɾ��ѡ��������֧����¼", "Admin_Payment.asp")
End Sub

Sub PaySuccess()
    Dim PaymentID, PaymentNum, UserName, OrderFormID, MoneyReceipt, eBankID, MoneyPayout, ClientID
    Dim rsPayment, sqlPayment, trs, rsUser
    PaymentID = Trim(Request("PaymentID"))
    ClientID = 0
    If PaymentID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��֧����ID��</li>"
        Exit Sub
    Else
        PaymentID = PE_CLng(PaymentID)
    End If
    
    sqlPayment = "select * from PE_Payment where PaymentID=" & PaymentID & ""
    Set rsPayment = Server.CreateObject("Adodb.RecordSet")
    rsPayment.Open sqlPayment, Conn, 1, 3
    If rsPayment.BOF And rsPayment.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ķ�����</li>"
        rsPayment.Close
        Set rsPayment = Nothing
        Exit Sub
    End If
    If rsPayment("Status") > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��֧�����Ѿ��ύ�����У�</li>"
    Else
        PaymentNum = rsPayment("PaymentNum")
        UserName = rsPayment("UserName")
        OrderFormID = rsPayment("OrderFormID")
        MoneyReceipt = rsPayment("MoneyPay")
        eBankID = rsPayment("eBankID")
        rsPayment("Status") = 3
        rsPayment("eBankInfo") = "֧�����"
        rsPayment("Remark") = "δ֪"
        rsPayment.Update
    End If
    rsPayment.Close
    Set rsPayment = Nothing

    Set rsUser = Conn.Execute("select ClientID from PE_User where UserName='" & UserName & "'")
    If Not (rsUser.EOF And rsUser.BOF) Then ClientID = rsUser(0)
      
    If FoundErr = True Then Exit Sub
    
    '����Ƿ��Ѿ��м�¼�����Ѿ��У�����д�����ݿ�Ĳ���
    Set trs = Conn.Execute("select * from PE_BankrollItem where PaymentID=" & PaymentID & "")
    If Not (trs.BOF And trs.EOF) Then
        ErrMsg = ErrMsg & "<li>�ʽ���ϸ���Ѿ�����ؼ�¼��</li>"
        FoundErr = True
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub
    
    '���ʽ��������ӽ��
    Conn.Execute ("update PE_User set Balance=Balance+" & MoneyReceipt & " where UserName='" & UserName & "'")
    
    ' ���ʽ���ϸ������������¼
    Call AddBankrollItem("", UserName, ClientID, MoneyReceipt, 3, "", eBankID, 1, 0, PaymentID, "����֧�����ţ�" & PaymentNum, Now())
        
    If OrderFormID > 0 Then
        Dim rsOrder
        Set rsOrder = Server.CreateObject("adodb.recordset")
        rsOrder.Open "select * from PE_OrderForm where OrderFormID=" & OrderFormID & "", Conn, 1, 3
        If Not (rsOrder.BOF And rsOrder.EOF) Then
            If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") Then
                If rsOrder("MoneyTotal") - rsOrder("MoneyReceipt") > MoneyReceipt Then
                    MoneyPayout = MoneyReceipt
                    rsOrder("MoneyReceipt") = rsOrder("MoneyReceipt") + MoneyReceipt
                Else
                    MoneyPayout = rsOrder("MoneyTotal") - rsOrder("MoneyReceipt")
                    rsOrder("MoneyReceipt") = rsOrder("MoneyTotal")
                End If
                rsOrder.Update
                '���ʽ���ϸ�������֧����¼
                Call AddBankrollItem("", UserName, ClientID, MoneyPayout, 4, "", 0, 2, OrderFormID, 0, "֧���������ã������ţ�" & rsOrder("OrderFormNum"), Now())
                
                '���ʽ�����п۳�֧������
                Conn.Execute ("update PE_User set Balance=Balance-" & MoneyPayout & " where UserName='" & UserName & "'")
            End If
        End If
        rsOrder.Close
        Set rsOrder = Nothing
    End If
    Call CloseConn
    Call WriteSuccessMsg("����֧���ɹ�", "Admin_Payment.asp")
    If ErrMsg <> "" Then
        FoundErr = True
    End If
End Sub
%>
