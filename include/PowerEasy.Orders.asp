<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim IsOfficial
IsOfficial = False
Dim strServerName
strServerName = LCase(Request.ServerVariables("SERVER_NAME"))
If strServerName = "www.powereasy.net" Or strServerName = "powereasy.net" Or strServerName = "www.powereasy.net.cn" Or strServerName = "powereasy.net.cn" Then
    IsOfficial = True
End If

Function GetOrderInfo(OrderFormID, UserName, ShowButton, OrderType)
    Dim rsOrder, sqlOrder, strOrderInfo
    If UserName = "" Then
        sqlOrder = "select * from PE_OrderForm where UserName='' and OrderFormID=" & OrderFormID & ""
    Else
        If OrderType = 1 Then
            sqlOrder = "select * from PE_OrderForm where AgentName='" & UserName & "' and OrderFormID=" & OrderFormID & ""
        Else
            sqlOrder = "select * from PE_OrderForm where UserName='" & UserName & "' and OrderFormID=" & OrderFormID & ""
        End If
    End If
    Set rsOrder = Conn.Execute(sqlOrder)
    If rsOrder.BOF And rsOrder.EOF Then
        FoundErr = True
        ErrMsg = "<li>�Ҳ���ָ���Ķ�����</li>"
        rsOrder.Close
        Set rsOrder = Nothing
        Exit Function
    End If

    strOrderInfo = strOrderInfo & "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>"
    strOrderInfo = strOrderInfo & "    <td height='22'><b>�� �� �� Ϣ</b>��������ţ�" & rsOrder("OrderFormNum") & "��</td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    strOrderInfo = strOrderInfo & "  <tr>"
    strOrderInfo = strOrderInfo & "    <td height='25'><table width='100%'  border='0' cellpadding='2' cellspacing='0'>"
    strOrderInfo = strOrderInfo & "      <tr class='tdbg'>"
    If rsOrder("UserName") = "" Then
        strOrderInfo = strOrderInfo & "        <td colspan='2'>�ͻ����ƣ�</td>"
    Else
        strOrderInfo = strOrderInfo & "        <td colspan='2'>�ͻ����ƣ�" & PE_HTMLEncode(GetClientName(rsOrder("ClientID"))) & "</td>"
    End If
    strOrderInfo = strOrderInfo & "        <td width='20%'>�� �� ����<a href='User_Order.asp'>" & rsOrder("UserName") & "</a></td>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>�� �� �̣�" & PE_HTMLEncode(rsOrder("AgentName")) & "</td>"
    strOrderInfo = strOrderInfo & "        <td width='26%'>�µ�ʱ�䣺<font color='red'>" & rsOrder("InputTime") & "</font></td>"
    strOrderInfo = strOrderInfo & "      <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "      <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>��Ҫ��Ʊ��"
    If rsOrder("NeedInvoice") = True Then
        strOrderInfo = strOrderInfo & "��"
    Else
        strOrderInfo = strOrderInfo & "<font color='red'>��</font>"
    End If
    strOrderInfo = strOrderInfo & "</td>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>�ѿ���Ʊ��"
    If rsOrder("Invoiced") = True Then
        strOrderInfo = strOrderInfo & "��"
    Else
        strOrderInfo = strOrderInfo & "<font color='red'>��</font>"
    End If
    strOrderInfo = strOrderInfo & "</td>"
    strOrderInfo = strOrderInfo & "        <td width='20%'>����״̬��<font color='red'>"
    Select Case rsOrder("Status")
    Case 0, 1
        strOrderInfo = strOrderInfo & "�ȴ�ȷ��"
    Case 2, 3
        strOrderInfo = strOrderInfo & "�Ѿ�ȷ��"
    Case 4
        strOrderInfo = strOrderInfo & "�ѽ���"
    End Select
    strOrderInfo = strOrderInfo & "</font></td>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>���������<font color='red'>"
    If rsOrder("MoneyTotal") > rsOrder("MoneyReceipt") Then
        If rsOrder("MoneyReceipt") > 0 Then
            strOrderInfo = strOrderInfo & "���ն���"
        Else
            strOrderInfo = strOrderInfo & "�ȴ����"
        End If
    Else
        strOrderInfo = strOrderInfo & "�Ѿ�����"
    End If
    strOrderInfo = strOrderInfo & "</font></td>"
    strOrderInfo = strOrderInfo & "        <td width='24%'>����״̬��<font color='red'>"
    Select Case rsOrder("DeliverStatus")
    Case 0, 1
        strOrderInfo = strOrderInfo & "������"
    Case 2
        strOrderInfo = strOrderInfo & "�ѷ���"
    Case 3
        strOrderInfo = strOrderInfo & "��ǩ��"
    End Select
    strOrderInfo = strOrderInfo & "</font></td>"
    strOrderInfo = strOrderInfo & "      </tr>"
    strOrderInfo = strOrderInfo & "    </table>      </td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    strOrderInfo = strOrderInfo & "  <tr align='center'>"
    strOrderInfo = strOrderInfo & "    <td height='25'><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>�ջ���������</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("ContacterName")) & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>��ϵ�绰��</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Phone")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg' valign='top'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>�ջ��˵�ַ��</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Address")) & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>�������룺</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & rsOrder("ZipCode") & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>�ջ������䣺</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & rsOrder("Email") & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>�ջ����ֻ���</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Mobile")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>���ʽ��</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & GetPaymentType(rsOrder("PaymentType")) & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>�ͻ���ʽ��</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & GetDeliverType(rsOrder("DeliverType")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg' valign='top'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>��Ʊ��Ϣ��</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>"
    If rsOrder("NeedInvoice") = True Then strOrderInfo = strOrderInfo & PE_HTMLEncode(rsOrder("InvoiceContent"))
    strOrderInfo = strOrderInfo & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>��ע/���ԣ�</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Remark")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "    </table></td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    strOrderInfo = strOrderInfo & "  <tr><td>"
    strOrderInfo = strOrderInfo & "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' bgcolor='#0099FF'>"
    strOrderInfo = strOrderInfo & "  <tr align='center' class='tdbg2' height='25'>"
    strOrderInfo = strOrderInfo & "    <td><b>�� Ʒ �� ��</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='45'><b>��λ</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='55'><b>����</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>ԭ��</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>ʵ��</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>ָ����</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='85'><b>�� ��</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>��������</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='45'><b>��ע</b></td>"
    strOrderInfo = strOrderInfo & "  </tr>"

    Dim dblPrice, dblAmount, dblSubtotal, dblTotal, TotalPresentExp, TotalPresentPoint, TotalPresentMoney, HaveSoft, HaveCard
    Dim rsOrderItem, rsCard
    dblSubtotal = 0
    dblTotal = 0
    TotalPresentExp = 0
    TotalPresentPoint = 0
    TotalPresentMoney = 0

    HaveSoft = False
    HaveCard = False
    Set rsOrderItem = Conn.Execute("select I.ItemID,P.ProductID,P.ProductName,P.ProductKind,I.SaleType,I.PresentExp,I.PresentMoney,I.PresentPoint,I.Price_Original,I.Price,I.TruePrice,I.Amount,P.Unit,I.BeginDate,I.ServiceTerm,I.Remark from PE_OrderFormItem I inner join PE_Product P on I.ProductID=P.ProductID where I.OrderFormID=" & rsOrder("OrderFormID") & " order by I.ItemID")
    Do While Not rsOrderItem.EOF
        dblPrice = rsOrderItem("TruePrice")
        dblAmount = rsOrderItem("Amount")
        dblSubtotal = dblPrice * dblAmount
        dblTotal = dblTotal + dblSubtotal
        TotalPresentExp = TotalPresentExp + rsOrderItem("PresentExp") * rsOrderItem("Amount")
        TotalPresentMoney = TotalPresentMoney + rsOrderItem("PresentMoney") * rsOrderItem("Amount")
        TotalPresentPoint = TotalPresentPoint + rsOrderItem("PresentPoint") * rsOrderItem("Amount")
        If rsOrderItem("ProductKind") = 2 Then
            HaveSoft = True
        ElseIf rsOrderItem("ProductKind") = 3 Then
            Set rsCard = Conn.Execute("select top 1 CardID from PE_Card where ProductID=" & rsOrderItem("ProductID") & " and OrderFormItemID=" & rsOrderItem("ItemID") & "")
            If rsCard.BOF And rsCard.EOF Then
                HaveCard = True
            End If
            Set rsCard = Nothing
        End If
        
        strOrderInfo = strOrderInfo & "  <tr valign='middle' class='tdbg' height='20'>"
        strOrderInfo = strOrderInfo & "    <td width='*'>" & rsOrderItem("ProductName")
        Select Case rsOrderItem("SaleType")
        Case 1 '��������
        
        Case 2 '����
            strOrderInfo = strOrderInfo & " <font color='red'>��������</font>"
        Case 3 '����
            strOrderInfo = strOrderInfo & " <font color='red'>�����ͣ�</font>"
        Case 4 '����
            strOrderInfo = strOrderInfo & " <font color='red'>��������</font>"
        End Select
            
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='45' align=center>" & rsOrderItem("Unit") & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='55' align='center'>" & dblAmount & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='65' align='right'>" & FormatNumber(rsOrderItem("Price_Original"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='65' align='right'>" & FormatNumber(rsOrderItem("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='65' align='right'>" & FormatNumber(dblPrice, 2, vbTrue, vbFalse, vbTrue) & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='85' align='right'>" & FormatNumber(dblSubtotal, 2, vbTrue, vbFalse, vbTrue) & "</td>"
        strOrderInfo = strOrderInfo & "    <td width='65' align=center>"
        If rsOrderItem("ServiceTerm") > 0 Then
            If DateAdd("yyyy", rsOrderItem("ServiceTerm"), rsOrderItem("BeginDate")) <= Now() Then
                strOrderInfo = strOrderInfo & "<font color='red'>"
            End If
        End If
        Select Case rsOrderItem("ServiceTerm")
        Case -1
            strOrderInfo = strOrderInfo & "������"
        Case 0
            strOrderInfo = strOrderInfo & "��"
        Case 1
            strOrderInfo = strOrderInfo & "һ��"
        Case 2
            strOrderInfo = strOrderInfo & "����"
        Case 3
            strOrderInfo = strOrderInfo & "����"
        Case 4
            strOrderInfo = strOrderInfo & "����"
        Case 5
            strOrderInfo = strOrderInfo & "����"
        Case Else
            strOrderInfo = strOrderInfo & "δ֪"
        End Select
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "    <td align=center width='40'>"
        If rsOrderItem("Remark") <> "" Then
            strOrderInfo = strOrderInfo & "<a href='#' title='" & rsOrderItem("Remark") & "'>�鿴</a>"
        End If
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "  </tr>"
        rsOrderItem.MoveNext
    Loop
    rsOrderItem.Close
    Set rsOrderItem = Nothing

    strOrderInfo = strOrderInfo & "  <tr class='tdbg' height='30' >"
    strOrderInfo = strOrderInfo & "    <td colspan='6' align='right'><b>�ϼƣ�</b></td>"
    strOrderInfo = strOrderInfo & "    <td align='right'><b>" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>"
    strOrderInfo = strOrderInfo & "    <td colspan='2'> </td>"
    strOrderInfo = strOrderInfo & "    </tr>"
    
    Dim Discount_Payment, Charge_Deliver, strTotalMoney
    Discount_Payment = rsOrder("Discount_Payment")
    Charge_Deliver = rsOrder("Charge_Deliver")
    
    strOrderInfo = strOrderInfo & "    <tr class='tdbg'>" & vbCrLf
    strOrderInfo = strOrderInfo & "      <td colspan='4'>���ʽ�ۿ��ʣ�" & Discount_Payment & "%"
    strTotalMoney = "ʵ�ʽ�(" & dblTotal & "��" & Discount_Payment & "%"
    If Discount_Payment > 0 And Discount_Payment < 100 Then
        dblTotal = dblTotal * Discount_Payment / 100
    End If
    strOrderInfo = strOrderInfo & "&nbsp;&nbsp;&nbsp;&nbsp;�˷ѣ�" & Charge_Deliver & " Ԫ"
    strTotalMoney = strTotalMoney & "��" & Charge_Deliver & ")"
    dblTotal = dblTotal + Charge_Deliver
    
    strOrderInfo = strOrderInfo & "&nbsp;&nbsp;&nbsp;&nbsp;˰�ʣ�" & TaxRate & "%&nbsp;&nbsp;&nbsp;&nbsp;�۸�˰��"
    If IncludeTax = True Then
        strOrderInfo = strOrderInfo & "��"
        If rsOrder("NeedInvoice") <> True Then
            strTotalMoney = strTotalMoney & "��(1-" & TaxRate & "%)"
            dblTotal = dblTotal * (100 - TaxRate) / 100
        Else
            strTotalMoney = strTotalMoney & "��100%"
        End If
    Else
        strOrderInfo = strOrderInfo & "��"
        If rsOrder("NeedInvoice") = True Then
            strTotalMoney = strTotalMoney & "��(1+" & TaxRate & "%)"
            dblTotal = dblTotal * (100 + TaxRate) / 100
        Else
            strTotalMoney = strTotalMoney & "��100%"
        End If
    End If
    strTotalMoney = strTotalMoney & "��" & dblTotal & " Ԫ"
    strOrderInfo = strOrderInfo & "<br>" & strTotalMoney

    strOrderInfo = strOrderInfo & "<br>���� <font color='red'>" & rsOrder("PresentMoney") + TotalPresentMoney & "</font> Ԫ�ֽ�ȯ������ <font color='red'>" & rsOrder("PresentExp") + TotalPresentExp & "</font> �����,���� <font color='red'>" & rsOrder("PresentPoint") + TotalPresentPoint & "</font> " & PointUnit & PointName
    
    strOrderInfo = strOrderInfo & "    </td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='2' align='right'><b>ʵ�ʽ�</b></td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td align=right><b> ��" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='2' align='left'><b>�Ѹ��</b>��"
    If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") Then
        strOrderInfo = strOrderInfo & "<font color='red'>" & FormatNumber(rsOrder("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue) & "</font><br>"
        strOrderInfo = strOrderInfo & "<font color='blue'><b>��Ƿ�</b>��" & FormatNumber(rsOrder("MoneyTotal") - rsOrder("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue) & "</font>"
    Else
        strOrderInfo = strOrderInfo & FormatNumber(rsOrder("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue)
    End If
    strOrderInfo = strOrderInfo & "</b></td></tr>"
    
    strOrderInfo = strOrderInfo & "</table></td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    If (UserName <> "" Or OrderType = 1) And ShowButton = True Then
        strOrderInfo = strOrderInfo & "  <tr align='right'>"
        strOrderInfo = strOrderInfo & "    <td height='30' align='center'>"
        If rsOrder("Status") = 1 And rsOrder("MoneyReceipt") = 0 Then
            strOrderInfo = strOrderInfo & "<input type='button' name='Submit' value='ɾ������' onClick=""javascript:if(confirm('ȷ��Ҫɾ���˶�����')){window.location.href='User_Order.asp?Action=DelOrder&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "';}"">"
        End If
        If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") Then
            If rsOrder("AgentName") = UserName And rsOrder("UserName") <> UserName Then
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='���������֧��' onClick=""javascript:if(confirm('ȷ��Ҫ֧���˶�����')){window.location.href='User_Order.asp?Action=AgentPayment&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "';}"">"
            Else
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='������пۿ�֧��' onClick=""window.location.href='User_Order.asp?Action=AddPayment&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "'"">"
            End If
            strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='����֧��' onClick=""window.location.href='../Shop/PayOnline.asp?OrderFormID=" & rsOrder("OrderFormID") & "'"">"
        Else
            If HaveSoft = True Then
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='���ز�Ʒ' onClick=""window.location.href='User_Down.asp'"">"
            End If
            If HaveCard = True Then
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='��ȡ�����ֵ��' onClick=""window.location.href='User_Exchange.asp?Action=GetCard'"">"
            End If
        End If
        If rsOrder("DeliverStatus") = 2 Then
            strOrderInfo = strOrderInfo & "&nbsp;<input type='button' name='Submit' value=' ǩ �� ' onClick=""javascript:if(confirm('ȷ���Ѿ��յ��˶����еĻ�������')){window.location.href='User_Order.asp?Action=Received2&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "';}"">"
        End If
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "  </tr>"
    End If
    strOrderInfo = strOrderInfo & "</table><br><b>ע��</b>��<font color='blue'>ԭ��</font>��ָ��Ʒ��ԭʼ���ۼۣ���<font color='green'>ʵ��</font>��ָϵͳ�Զ������������Ʒ���ռ۸񣬡�<font color='red'>ָ����</font>��ָ����Ա�ֶ�ָ�������ռ۸���Ʒ���������ۼ۸��ԡ�ָ���ۡ�Ϊ׼��<br>"

    strOrderInfo = strOrderInfo & "<script language='javascript'>" & vbCrLf
    strOrderInfo = strOrderInfo & "var tID=0;" & vbCrLf
    strOrderInfo = strOrderInfo & "function ShowTabs(ID){" & vbCrLf
    strOrderInfo = strOrderInfo & "  if(ID!=tID){" & vbCrLf
    strOrderInfo = strOrderInfo & "    TabTitle[tID].className='title5';" & vbCrLf
    strOrderInfo = strOrderInfo & "    TabTitle[ID].className='title6';" & vbCrLf
    strOrderInfo = strOrderInfo & "    Tabs[tID].style.display='none';" & vbCrLf
    strOrderInfo = strOrderInfo & "    Tabs[ID].style.display='';" & vbCrLf
    strOrderInfo = strOrderInfo & "    tID=ID;" & vbCrLf
    strOrderInfo = strOrderInfo & "  }" & vbCrLf
    strOrderInfo = strOrderInfo & "}" & vbCrLf
    strOrderInfo = strOrderInfo & "</script>" & vbCrLf

    strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��Ʊ��¼</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>���˻���¼</td>" & vbCrLf
    If IsOfficial Then
        strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>������к�</td>" & vbCrLf
    End If
    strOrderInfo = strOrderInfo & "  <td>&nbsp;</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "</tr></table>" & vbCrLf

    strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='Tabs' style='display:'>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>������</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='120'>����ʱ��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>���׷�ʽ</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='50'>����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>֧�����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td>��ע/˵��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

    Dim rsBankroll, sqlBankroll, TotalIncome, TotalPayout
    sqlBankroll = "select * from PE_BankrollItem where OrderFormID=" & OrderFormID & " and Money<0 order by ItemID desc"
    Set rsBankroll = Conn.Execute(sqlBankroll)
    If rsBankroll.BOF And rsBankroll.EOF Then
        strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>û����ظ����¼</td></tr>" & vbCrLf
    Else
        Do While Not rsBankroll.EOF
            If rsBankroll("Money") > 0 Then
                TotalIncome = TotalIncome + rsBankroll("Money")
            Else
                TotalPayout = TotalPayout + rsBankroll("Money")
            End If

            strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsBankroll("UserName") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='120' align='center'>" & rsBankroll("DateAndTime") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>"
            Select Case rsBankroll("MoneyType")
            Case 1
                strOrderInfo = strOrderInfo & "�ֽ�"
            Case 2
                strOrderInfo = strOrderInfo & "���л��"
            Case 3
                strOrderInfo = strOrderInfo & "����֧��"
            Case 4
                strOrderInfo = strOrderInfo & "�������"
            End Select
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='50' align='center'>"
            Select Case rsBankroll("CurrencyType")
            Case 1
                strOrderInfo = strOrderInfo & "�����"
            Case 2
                strOrderInfo = strOrderInfo & "��Ԫ"
            Case 3
                strOrderInfo = strOrderInfo & "����"
            End Select
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='right'>" & FormatNumber(Abs(rsBankroll("Money")), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td>" & rsBankroll("Remark") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

            rsBankroll.MoveNext
        Loop
    End If
    rsBankroll.Close
    Set rsBankroll = Nothing

    strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='4' align='right'>�ϼƽ�</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td align='right'>" & FormatNumber(TotalPayout, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='2' align='center'> </td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf
    strOrderInfo = strOrderInfo & "</table>" & vbCrLf
    
    strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='Tabs' style='display:none'>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>��Ʊ����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>��Ʊ����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td>��Ʊ̧ͷ</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>��Ʊ���</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>��Ʊ��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='120'>��Ʊʱ��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

    Dim rsInvoice
    Set rsInvoice = Conn.Execute("select * from PE_InvoiceItem where OrderFormID=" & OrderFormID & " order by InvoiceID")
    If rsInvoice.BOF And rsInvoice.EOF Then
        strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>û����ط�Ʊ��¼</td></tr>" & vbCrLf
    Else
        Do While Not rsInvoice.EOF
            strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsInvoice("InvoiceDate") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>"
            Select Case rsInvoice("InvoiceType")
            Case 0
                strOrderInfo = strOrderInfo & "��˰��ͨ��Ʊ"
            Case 1
                strOrderInfo = strOrderInfo & "��˰��ͨ��Ʊ"
            Case 2
                strOrderInfo = strOrderInfo & "��ֵ˰��Ʊ"
            End Select
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsInvoice("InvoiceNum") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td>" & rsInvoice("InvoiceTitle") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='right'>" & rsInvoice("TotalMoney") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>" & rsInvoice("Drawer") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='120' align='center'>" & rsInvoice("InputTime") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

            rsInvoice.MoveNext
        Loop
    End If
    rsInvoice.Close
    Set rsInvoice = Nothing
    strOrderInfo = strOrderInfo & "</table>" & vbCrLf
    
    strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='Tabs' style='display:none'>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>����/�ͻ��˻�</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>��ݹ�˾��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>��ݵ���</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>������</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>¼��Ա</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>�ͻ���ǩ��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td>��ע/�˻�ԭ��</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>����</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

    Dim rsDeliver
    Set rsDeliver = Conn.Execute("select * from PE_DeliverItem where OrderFormID=" & OrderFormID & " order by DeliverID")
    If rsDeliver.BOF And rsDeliver.EOF Then
        strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>û����ط��˻���¼</td></tr>" & vbCrLf
    Else
        Do While Not rsDeliver.EOF
            strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsDeliver("DeliverDate") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>"
            If rsDeliver("DeliverDirection") = 1 Then
                strOrderInfo = strOrderInfo & "����"
            Else
                strOrderInfo = strOrderInfo & "<font color='red'>�˻�</font>"
            End If
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsDeliver("ExpressCompany") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsDeliver("ExpressNumber") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>" & rsDeliver("HandlerName") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>" & rsDeliver("Inputer") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>"
            If rsDeliver("Received") = True Then
                strOrderInfo = strOrderInfo & "<font color='red'><b>��</b></font>"
            End If
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td>" & rsDeliver("Remark") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>"
            If rsDeliver("DeliverDirection") = 1 And rsDeliver("Received") <> True Then
                strOrderInfo = strOrderInfo & "<a href='User_Order.asp?Action=Received&OrderType=" & OrderType & "&DeliverID=" & rsDeliver("DeliverID") & "' onclick=""return confirm('ȷ���Ѿ��յ��˶����еĻ�������');"">ǩ��</a>"
            End If
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

            rsDeliver.MoveNext
        Loop
    End If
    rsDeliver.Close
    Set rsDeliver = Nothing
    strOrderInfo = strOrderInfo & "</table>" & vbCrLf

    If IsOfficial Then
        strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='Tabs' style='display:none'>" & vbCrLf
        strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='150'>��Ӧ�Ĳ�Ʒ</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='150'>�󶨵�����</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='120'>����ʱ��</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='80'>��ֹ����</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td>���к�����</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

        Dim rsSiteKey, arrEdition, arrTemp
        arrEdition = Array("CMS 2006��׼��", "CMS 2006רҵ��", "CMS 2006��ҵ��", "eShop 2006��׼��", "eShop 2006רҵ��", "eShop 2006��ҵ��", "CRM 2006��׼��", "CRM 2006רҵ��", "CRM 2006��ҵ��", "�ʾ����ϵͳ")
        Set rsSiteKey = Conn.Execute("select * from PE_SiteKey where OrderFormID=" & OrderFormID & " order by CreateTime Desc")
        If rsSiteKey.BOF And rsSiteKey.EOF Then
            strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>û��������к�</td></tr>" & vbCrLf
        Else
            Do While Not rsSiteKey.EOF
                strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
                strOrderInfo = strOrderInfo & "    <td width='150' align='left'>"
                Select Case rsSiteKey("SystemVersion")
                Case 403
                    strOrderInfo = strOrderInfo & "����4.03"
                    Select Case rsSiteKey("SystemEdition")
                    Case "1"
                        strOrderInfo = strOrderInfo & "����"
                    Case "3"
                        strOrderInfo = strOrderInfo & "��ҵ"
                    End Select
                    Select Case rsSiteKey("DatabaseType")
                    Case "1"
                        strOrderInfo = strOrderInfo & "Access��"
                    Case "2"
                        strOrderInfo = strOrderInfo & "SQL��"
                    End Select
                Case 2005
                    strOrderInfo = strOrderInfo & "����2005"
                    Select Case rsSiteKey("SystemEdition")
                    Case "1"
                        strOrderInfo = strOrderInfo & "����"
                    Case "2"
                        strOrderInfo = strOrderInfo & "��׼"
                    Case "3"
                        strOrderInfo = strOrderInfo & "��ҵ"
                    Case "4"
                        strOrderInfo = strOrderInfo & "ѧУ"
                    Case "5"
                        strOrderInfo = strOrderInfo & "����"
                    Case "9"
                        strOrderInfo = strOrderInfo & "ȫ��"
                    End Select
                    Select Case rsSiteKey("DatabaseType")
                    Case 1
                        strOrderInfo = strOrderInfo & "Access��"
                    Case 2
                        strOrderInfo = strOrderInfo & "SQL��"
                    End Select
                Case 2006
                    strOrderInfo = strOrderInfo & "����"
                    arrTemp = Split(rsSiteKey("SystemEdition"), "|")
                    strOrderInfo = strOrderInfo & arrEdition(CLng(arrTemp(0)) - 1)
                    If arrTemp(1) = "1" Then
                        strOrderInfo = strOrderInfo & "+SDMS"
                    End If
                    If arrTemp(2) = "1" Then
                        strOrderInfo = strOrderInfo & "+Equipment"
                    End If
                    If arrTemp(3) = "1" Then
                        strOrderInfo = strOrderInfo & "+HR"
                    End If
                    If arrTemp(4) = "1" Then
                        strOrderInfo = strOrderInfo & "+SD"
                    End If
                    If arrTemp(5) = "1" Then
                        strOrderInfo = strOrderInfo & "+House"
                    End If
                End Select
                strOrderInfo = strOrderInfo & "</td>" & vbCrLf
                strOrderInfo = strOrderInfo & "    <td width='150' align='center'>" & rsSiteKey("SiteUrl") & "</td>" & vbCrLf
                strOrderInfo = strOrderInfo & "    <td width='120' align='center'>" & rsSiteKey("CreateTime") & "</td>" & vbCrLf
                strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsSiteKey("EndDate") & "</td>" & vbCrLf
                strOrderInfo = strOrderInfo & "    <td><textarea cols='30' rows='4'>" & rsSiteKey("SiteKey") & "</textarea></td>" & vbCrLf
                strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

                rsSiteKey.MoveNext
            Loop
            strOrderInfo = strOrderInfo & "<tr class='tdbg5'><td colspan='20' height='50' align='left'><b>���к�ʹ�÷�����</b><br>�������Ӧ���������к�����ȫ�����Ƶ�����ϵͳ�ĺ�̨����ϵͳ���ã�����վ��Ϣ���ã������кſ��У��������ü��ɡ�</td></tr>" & vbCrLf
        End If
        rsSiteKey.Close
        Set rsSiteKey = Nothing
        strOrderInfo = strOrderInfo & "</table>" & vbCrLf
    End If

    GetOrderInfo = strOrderInfo
End Function


Function GetPaymentType(PaymentType)
    Dim rsPaymentType
    Set rsPaymentType = Conn.Execute("select TypeName from PE_PaymentType where TypeID=" & PaymentType & "")
    If rsPaymentType.BOF And rsPaymentType.EOF Then
        GetPaymentType = ""
    Else
        GetPaymentType = rsPaymentType("TypeName")
    End If
    Set rsPaymentType = Nothing
End Function

Function GetDeliverType(DeliverType)
    Dim rsDeliverType
    Set rsDeliverType = Conn.Execute("select TypeName from PE_DeliverType where TypeID=" & DeliverType & "")
    If rsDeliverType.BOF And rsDeliverType.EOF Then
        GetDeliverType = ""
    Else
        GetDeliverType = rsDeliverType("TypeName")
    End If
    Set rsDeliverType = Nothing
End Function

Function GetClientName(ClientID)
    Dim rsClient
    Set rsClient = Conn.Execute("select ClientName from PE_Client where ClientID=" & ClientID & "")
    If rsClient.BOF And rsClient.EOF Then
        GetClientName = ""
    Else
        GetClientName = rsClient("ClientName")
    End If
    Set rsClient = Nothing
End Function
%>
