<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
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
        ErrMsg = "<li>找不到指定的订单！</li>"
        rsOrder.Close
        Set rsOrder = Nothing
        Exit Function
    End If

    strOrderInfo = strOrderInfo & "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>"
    strOrderInfo = strOrderInfo & "    <td height='22'><b>订 单 信 息</b>（订单编号：" & rsOrder("OrderFormNum") & "）</td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    strOrderInfo = strOrderInfo & "  <tr>"
    strOrderInfo = strOrderInfo & "    <td height='25'><table width='100%'  border='0' cellpadding='2' cellspacing='0'>"
    strOrderInfo = strOrderInfo & "      <tr class='tdbg'>"
    If rsOrder("UserName") = "" Then
        strOrderInfo = strOrderInfo & "        <td colspan='2'>客户名称：</td>"
    Else
        strOrderInfo = strOrderInfo & "        <td colspan='2'>客户名称：" & PE_HTMLEncode(GetClientName(rsOrder("ClientID"))) & "</td>"
    End If
    strOrderInfo = strOrderInfo & "        <td width='20%'>用 户 名：<a href='User_Order.asp'>" & rsOrder("UserName") & "</a></td>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>代 理 商：" & PE_HTMLEncode(rsOrder("AgentName")) & "</td>"
    strOrderInfo = strOrderInfo & "        <td width='26%'>下单时间：<font color='red'>" & rsOrder("InputTime") & "</font></td>"
    strOrderInfo = strOrderInfo & "      <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "      <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>需要发票："
    If rsOrder("NeedInvoice") = True Then
        strOrderInfo = strOrderInfo & "√"
    Else
        strOrderInfo = strOrderInfo & "<font color='red'>×</font>"
    End If
    strOrderInfo = strOrderInfo & "</td>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>已开发票："
    If rsOrder("Invoiced") = True Then
        strOrderInfo = strOrderInfo & "√"
    Else
        strOrderInfo = strOrderInfo & "<font color='red'>×</font>"
    End If
    strOrderInfo = strOrderInfo & "</td>"
    strOrderInfo = strOrderInfo & "        <td width='20%'>订单状态：<font color='red'>"
    Select Case rsOrder("Status")
    Case 0, 1
        strOrderInfo = strOrderInfo & "等待确认"
    Case 2, 3
        strOrderInfo = strOrderInfo & "已经确认"
    Case 4
        strOrderInfo = strOrderInfo & "已结清"
    End Select
    strOrderInfo = strOrderInfo & "</font></td>"
    strOrderInfo = strOrderInfo & "        <td width='18%'>付款情况：<font color='red'>"
    If rsOrder("MoneyTotal") > rsOrder("MoneyReceipt") Then
        If rsOrder("MoneyReceipt") > 0 Then
            strOrderInfo = strOrderInfo & "已收定金"
        Else
            strOrderInfo = strOrderInfo & "等待汇款"
        End If
    Else
        strOrderInfo = strOrderInfo & "已经付清"
    End If
    strOrderInfo = strOrderInfo & "</font></td>"
    strOrderInfo = strOrderInfo & "        <td width='24%'>物流状态：<font color='red'>"
    Select Case rsOrder("DeliverStatus")
    Case 0, 1
        strOrderInfo = strOrderInfo & "配送中"
    Case 2
        strOrderInfo = strOrderInfo & "已发货"
    Case 3
        strOrderInfo = strOrderInfo & "已签收"
    End Select
    strOrderInfo = strOrderInfo & "</font></td>"
    strOrderInfo = strOrderInfo & "      </tr>"
    strOrderInfo = strOrderInfo & "    </table>      </td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    strOrderInfo = strOrderInfo & "  <tr align='center'>"
    strOrderInfo = strOrderInfo & "    <td height='25'><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>收货人姓名：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("ContacterName")) & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>联系电话：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Phone")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg' valign='top'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>收货人地址：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Address")) & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>邮政编码：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & rsOrder("ZipCode") & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>收货人邮箱：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & rsOrder("Email") & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>收货人手机：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Mobile")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>付款方式：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & GetPaymentType(rsOrder("PaymentType")) & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>送货方式：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & GetDeliverType(rsOrder("DeliverType")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "        <tr class='tdbg' valign='top'>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>发票信息：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>"
    If rsOrder("NeedInvoice") = True Then strOrderInfo = strOrderInfo & PE_HTMLEncode(rsOrder("InvoiceContent"))
    strOrderInfo = strOrderInfo & "</td>"
    strOrderInfo = strOrderInfo & "          <td width='12%' class='tdbg5' align='right'>备注/留言：</td>"
    strOrderInfo = strOrderInfo & "          <td width='38%'>" & PE_HTMLEncode(rsOrder("Remark")) & "</td>"
    strOrderInfo = strOrderInfo & "        </tr>"
    strOrderInfo = strOrderInfo & "    </table></td>"
    strOrderInfo = strOrderInfo & "  </tr>"
    strOrderInfo = strOrderInfo & "  <tr><td>"
    strOrderInfo = strOrderInfo & "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' bgcolor='#0099FF'>"
    strOrderInfo = strOrderInfo & "  <tr align='center' class='tdbg2' height='25'>"
    strOrderInfo = strOrderInfo & "    <td><b>商 品 名 称</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='45'><b>单位</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='55'><b>数量</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>原价</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>实价</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>指定价</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='85'><b>金 额</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='65'><b>服务期限</b></td>"
    strOrderInfo = strOrderInfo & "    <td width='45'><b>备注</b></td>"
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
        Case 1 '正常销售
        
        Case 2 '换购
            strOrderInfo = strOrderInfo & " <font color='red'>（换购）</font>"
        Case 3 '赠送
            strOrderInfo = strOrderInfo & " <font color='red'>（赠送）</font>"
        Case 4 '批发
            strOrderInfo = strOrderInfo & " <font color='red'>（批发）</font>"
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
            strOrderInfo = strOrderInfo & "无限期"
        Case 0
            strOrderInfo = strOrderInfo & "无"
        Case 1
            strOrderInfo = strOrderInfo & "一年"
        Case 2
            strOrderInfo = strOrderInfo & "两年"
        Case 3
            strOrderInfo = strOrderInfo & "三年"
        Case 4
            strOrderInfo = strOrderInfo & "四年"
        Case 5
            strOrderInfo = strOrderInfo & "五年"
        Case Else
            strOrderInfo = strOrderInfo & "未知"
        End Select
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "    <td align=center width='40'>"
        If rsOrderItem("Remark") <> "" Then
            strOrderInfo = strOrderInfo & "<a href='#' title='" & rsOrderItem("Remark") & "'>查看</a>"
        End If
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "  </tr>"
        rsOrderItem.MoveNext
    Loop
    rsOrderItem.Close
    Set rsOrderItem = Nothing

    strOrderInfo = strOrderInfo & "  <tr class='tdbg' height='30' >"
    strOrderInfo = strOrderInfo & "    <td colspan='6' align='right'><b>合计：</b></td>"
    strOrderInfo = strOrderInfo & "    <td align='right'><b>" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>"
    strOrderInfo = strOrderInfo & "    <td colspan='2'> </td>"
    strOrderInfo = strOrderInfo & "    </tr>"
    
    Dim Discount_Payment, Charge_Deliver, strTotalMoney
    Discount_Payment = rsOrder("Discount_Payment")
    Charge_Deliver = rsOrder("Charge_Deliver")
    
    strOrderInfo = strOrderInfo & "    <tr class='tdbg'>" & vbCrLf
    strOrderInfo = strOrderInfo & "      <td colspan='4'>付款方式折扣率：" & Discount_Payment & "%"
    strTotalMoney = "实际金额：(" & dblTotal & "×" & Discount_Payment & "%"
    If Discount_Payment > 0 And Discount_Payment < 100 Then
        dblTotal = dblTotal * Discount_Payment / 100
    End If
    strOrderInfo = strOrderInfo & "&nbsp;&nbsp;&nbsp;&nbsp;运费：" & Charge_Deliver & " 元"
    strTotalMoney = strTotalMoney & "＋" & Charge_Deliver & ")"
    dblTotal = dblTotal + Charge_Deliver
    
    strOrderInfo = strOrderInfo & "&nbsp;&nbsp;&nbsp;&nbsp;税率：" & TaxRate & "%&nbsp;&nbsp;&nbsp;&nbsp;价格含税："
    If IncludeTax = True Then
        strOrderInfo = strOrderInfo & "是"
        If rsOrder("NeedInvoice") <> True Then
            strTotalMoney = strTotalMoney & "×(1-" & TaxRate & "%)"
            dblTotal = dblTotal * (100 - TaxRate) / 100
        Else
            strTotalMoney = strTotalMoney & "×100%"
        End If
    Else
        strOrderInfo = strOrderInfo & "否"
        If rsOrder("NeedInvoice") = True Then
            strTotalMoney = strTotalMoney & "×(1+" & TaxRate & "%)"
            dblTotal = dblTotal * (100 + TaxRate) / 100
        Else
            strTotalMoney = strTotalMoney & "×100%"
        End If
    End If
    strTotalMoney = strTotalMoney & "＝" & dblTotal & " 元"
    strOrderInfo = strOrderInfo & "<br>" & strTotalMoney

    strOrderInfo = strOrderInfo & "<br>返还 <font color='red'>" & rsOrder("PresentMoney") + TotalPresentMoney & "</font> 元现金券，赠送 <font color='red'>" & rsOrder("PresentExp") + TotalPresentExp & "</font> 点积分,赠送 <font color='red'>" & rsOrder("PresentPoint") + TotalPresentPoint & "</font> " & PointUnit & PointName
    
    strOrderInfo = strOrderInfo & "    </td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='2' align='right'><b>实际金额：</b></td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td align=right><b> ￥" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='2' align='left'><b>已付款：</b>￥"
    If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") Then
        strOrderInfo = strOrderInfo & "<font color='red'>" & FormatNumber(rsOrder("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue) & "</font><br>"
        strOrderInfo = strOrderInfo & "<font color='blue'><b>尚欠款：</b>￥" & FormatNumber(rsOrder("MoneyTotal") - rsOrder("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue) & "</font>"
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
            strOrderInfo = strOrderInfo & "<input type='button' name='Submit' value='删除订单' onClick=""javascript:if(confirm('确定要删除此订单吗？')){window.location.href='User_Order.asp?Action=DelOrder&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "';}"">"
        End If
        If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") Then
            If rsOrder("AgentName") = UserName And rsOrder("UserName") <> UserName Then
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='代理商余额支付' onClick=""javascript:if(confirm('确定要支付此订单吗？')){window.location.href='User_Order.asp?Action=AgentPayment&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "';}"">"
            Else
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='从余额中扣款支付' onClick=""window.location.href='User_Order.asp?Action=AddPayment&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "'"">"
            End If
            strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='在线支付' onClick=""window.location.href='../Shop/PayOnline.asp?OrderFormID=" & rsOrder("OrderFormID") & "'"">"
        Else
            If HaveSoft = True Then
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='下载产品' onClick=""window.location.href='User_Down.asp'"">"
            End If
            If HaveCard = True Then
                strOrderInfo = strOrderInfo & "&nbsp;&nbsp;<input type='button' name='Submit' value='获取虚拟充值卡' onClick=""window.location.href='User_Exchange.asp?Action=GetCard'"">"
            End If
        End If
        If rsOrder("DeliverStatus") = 2 Then
            strOrderInfo = strOrderInfo & "&nbsp;<input type='button' name='Submit' value=' 签 收 ' onClick=""javascript:if(confirm('确定已经收到此订单中的货物了吗？')){window.location.href='User_Order.asp?Action=Received2&OrderType=" & OrderType & "&OrderFormID=" & rsOrder("OrderFormID") & "';}"">"
        End If
        strOrderInfo = strOrderInfo & "</td>"
        strOrderInfo = strOrderInfo & "  </tr>"
    End If
    strOrderInfo = strOrderInfo & "</table><br><b>注：</b>“<font color='blue'>原价</font>”指商品的原始零售价，“<font color='green'>实价</font>”指系统自动计算出来的商品最终价格，“<font color='red'>指定价</font>”指管理员手动指定的最终价格。商品的最终销售价格以“指定价”为准。<br>"

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
    strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>付款信息</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>发票记录</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>发退货记录</td>" & vbCrLf
    If IsOfficial Then
        strOrderInfo = strOrderInfo & "  <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>相关序列号</td>" & vbCrLf
    End If
    strOrderInfo = strOrderInfo & "  <td>&nbsp;</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "</tr></table>" & vbCrLf

    strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='Tabs' style='display:'>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>付款人</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='120'>交易时间</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>交易方式</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='50'>币种</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>支出金额</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td>备注/说明</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

    Dim rsBankroll, sqlBankroll, TotalIncome, TotalPayout
    sqlBankroll = "select * from PE_BankrollItem where OrderFormID=" & OrderFormID & " and Money<0 order by ItemID desc"
    Set rsBankroll = Conn.Execute(sqlBankroll)
    If rsBankroll.BOF And rsBankroll.EOF Then
        strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>没有相关付款记录</td></tr>" & vbCrLf
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
                strOrderInfo = strOrderInfo & "现金"
            Case 2
                strOrderInfo = strOrderInfo & "银行汇款"
            Case 3
                strOrderInfo = strOrderInfo & "在线支付"
            Case 4
                strOrderInfo = strOrderInfo & "虚拟货币"
            End Select
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='50' align='center'>"
            Select Case rsBankroll("CurrencyType")
            Case 1
                strOrderInfo = strOrderInfo & "人民币"
            Case 2
                strOrderInfo = strOrderInfo & "美元"
            Case 3
                strOrderInfo = strOrderInfo & "其他"
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
    strOrderInfo = strOrderInfo & "    <td colspan='4' align='right'>合计金额：</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td align='right'>" & FormatNumber(TotalPayout, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td colspan='2' align='center'> </td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf
    strOrderInfo = strOrderInfo & "</table>" & vbCrLf
    
    strOrderInfo = strOrderInfo & "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='Tabs' style='display:none'>" & vbCrLf
    strOrderInfo = strOrderInfo & "  <tr align='center' class='title'>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>日期</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>发票类型</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>发票号码</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td>发票抬头</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>发票金额</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>开票人</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='120'>开票时间</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

    Dim rsInvoice
    Set rsInvoice = Conn.Execute("select * from PE_InvoiceItem where OrderFormID=" & OrderFormID & " order by InvoiceID")
    If rsInvoice.BOF And rsInvoice.EOF Then
        strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>没有相关发票记录</td></tr>" & vbCrLf
    Else
        Do While Not rsInvoice.EOF
            strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsInvoice("InvoiceDate") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>"
            Select Case rsInvoice("InvoiceType")
            Case 0
                strOrderInfo = strOrderInfo & "地税普通发票"
            Case 1
                strOrderInfo = strOrderInfo & "国税普通发票"
            Case 2
                strOrderInfo = strOrderInfo & "增值税发票"
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
    strOrderInfo = strOrderInfo & "    <td width='80'>日期</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>发货/客户退货</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>快递公司名</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='80'>快递单号</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>经手人</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>录入员</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>客户已签收</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td>备注/退货原因</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "    <td width='60'>操作</td>" & vbCrLf
    strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

    Dim rsDeliver
    Set rsDeliver = Conn.Execute("select * from PE_DeliverItem where OrderFormID=" & OrderFormID & " order by DeliverID")
    If rsDeliver.BOF And rsDeliver.EOF Then
        strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>没有相关发退货记录</td></tr>" & vbCrLf
    Else
        Do While Not rsDeliver.EOF
            strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsDeliver("DeliverDate") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>"
            If rsDeliver("DeliverDirection") = 1 Then
                strOrderInfo = strOrderInfo & "发货"
            Else
                strOrderInfo = strOrderInfo & "<font color='red'>退货</font>"
            End If
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsDeliver("ExpressCompany") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='80' align='center'>" & rsDeliver("ExpressNumber") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>" & rsDeliver("HandlerName") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>" & rsDeliver("Inputer") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>"
            If rsDeliver("Received") = True Then
                strOrderInfo = strOrderInfo & "<font color='red'><b>√</b></font>"
            End If
            strOrderInfo = strOrderInfo & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td>" & rsDeliver("Remark") & "</td>" & vbCrLf
            strOrderInfo = strOrderInfo & "    <td width='60' align='center'>"
            If rsDeliver("DeliverDirection") = 1 And rsDeliver("Received") <> True Then
                strOrderInfo = strOrderInfo & "<a href='User_Order.asp?Action=Received&OrderType=" & OrderType & "&DeliverID=" & rsDeliver("DeliverID") & "' onclick=""return confirm('确定已经收到此订单中的货物了吗？');"">签收</a>"
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
        strOrderInfo = strOrderInfo & "    <td width='150'>对应的产品</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='150'>绑定的域名</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='120'>生成时间</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td width='80'>截止日期</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "    <td>序列号内容</td>" & vbCrLf
        strOrderInfo = strOrderInfo & "  </tr>" & vbCrLf

        Dim rsSiteKey, arrEdition, arrTemp
        arrEdition = Array("CMS 2006标准版", "CMS 2006专业版", "CMS 2006企业版", "eShop 2006标准版", "eShop 2006专业版", "eShop 2006企业版", "CRM 2006标准版", "CRM 2006专业版", "CRM 2006企业版", "问卷调查系统")
        Set rsSiteKey = Conn.Execute("select * from PE_SiteKey where OrderFormID=" & OrderFormID & " order by CreateTime Desc")
        If rsSiteKey.BOF And rsSiteKey.EOF Then
            strOrderInfo = strOrderInfo & "<tr class='tdbg'><td colspan='20' height='50' align='center'>没有相关序列号</td></tr>" & vbCrLf
        Else
            Do While Not rsSiteKey.EOF
                strOrderInfo = strOrderInfo & "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
                strOrderInfo = strOrderInfo & "    <td width='150' align='left'>"
                Select Case rsSiteKey("SystemVersion")
                Case 403
                    strOrderInfo = strOrderInfo & "动易4.03"
                    Select Case rsSiteKey("SystemEdition")
                    Case "1"
                        strOrderInfo = strOrderInfo & "个人"
                    Case "3"
                        strOrderInfo = strOrderInfo & "企业"
                    End Select
                    Select Case rsSiteKey("DatabaseType")
                    Case "1"
                        strOrderInfo = strOrderInfo & "Access版"
                    Case "2"
                        strOrderInfo = strOrderInfo & "SQL版"
                    End Select
                Case 2005
                    strOrderInfo = strOrderInfo & "动易2005"
                    Select Case rsSiteKey("SystemEdition")
                    Case "1"
                        strOrderInfo = strOrderInfo & "个人"
                    Case "2"
                        strOrderInfo = strOrderInfo & "标准"
                    Case "3"
                        strOrderInfo = strOrderInfo & "企业"
                    Case "4"
                        strOrderInfo = strOrderInfo & "学校"
                    Case "5"
                        strOrderInfo = strOrderInfo & "政府"
                    Case "9"
                        strOrderInfo = strOrderInfo & "全能"
                    End Select
                    Select Case rsSiteKey("DatabaseType")
                    Case 1
                        strOrderInfo = strOrderInfo & "Access版"
                    Case 2
                        strOrderInfo = strOrderInfo & "SQL版"
                    End Select
                Case 2006
                    strOrderInfo = strOrderInfo & "动易"
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
            strOrderInfo = strOrderInfo & "<tr class='tdbg5'><td colspan='20' height='50' align='left'><b>序列号使用方法：</b><br>将上面对应域名的序列号内容全部复制到动易系统的后台－－系统设置－－网站信息配置－－序列号框中，保存设置即可。</td></tr>" & vbCrLf
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
