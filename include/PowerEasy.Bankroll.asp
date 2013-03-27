<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


'**************************************************
'方法名：AddBankrollItem
'作  用：添加一条资金明细
'参  数：AdminName ----管理员名称
'        UserName ---- 用户名称
'        ClientID ----客户ID
'        Money ---- 金额
'        MoneyType ---- 类型  1--现金 2--汇款 3--在线支付  4--虚拟货币
'        Bank ---- 银行名称
'        eBankID ---- 网上支付平台ID
'        Income_PayOut ---- 类型  1--收入  2--支出  3--开户
'        OrderFormID ---- 支出时的订单ID
'        PaymentID ---- 在线支付的支付单ID
'        Remark ---- 备注
'        DateAndTime ---- 发生时间
'**************************************************
Sub AddBankrollItem(AdminName, UserName, ClientID, Money, MoneyType, Bank, eBankID, Income_PayOut, OrderFormID, PaymentID, Remark, DateAndTime)
    Dim rsBankroll, sqlBankroll
    sqlBankroll = "select top 1 * from PE_BankrollItem"
    Set rsBankroll = Server.CreateObject("adodb.recordset")
    rsBankroll.Open sqlBankroll, Conn, 1, 3
    rsBankroll.addnew
    rsBankroll("UserName") = UserName
    rsBankroll("ClientID") = ClientID
    rsBankroll("DateAndTime") = DateAndTime
    If Income_PayOut = 2 Then
        rsBankroll("Money") = 0 - Abs(Money)
    Else
        rsBankroll("Money") = Abs(Money)
    End If
    rsBankroll("MoneyType") = MoneyType
    rsBankroll("CurrencyType") = 1
    rsBankroll("Bank") = Bank
    rsBankroll("eBankID") = eBankID
    rsBankroll("Income_Payout") = Income_PayOut
    If OrderFormID = 0 Then
        rsBankroll("OrderFormID") = -GetRndNum(8)
    Else
        rsBankroll("OrderFormID") = OrderFormID
    End If
    rsBankroll("PaymentID") = PaymentID
    rsBankroll("Remark") = Remark
    rsBankroll("LogTime") = Now()
    rsBankroll("Inputer") = AdminName
    rsBankroll("IP") = UserTrueIP
    rsBankroll.Update
    rsBankroll.Close
    Set rsBankroll = Nothing
End Sub
%>
