<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************


'**************************************************
'��������AddBankrollItem
'��  �ã����һ���ʽ���ϸ
'��  ����AdminName ----����Ա����
'        UserName ---- �û�����
'        ClientID ----�ͻ�ID
'        Money ---- ���
'        MoneyType ---- ����  1--�ֽ� 2--��� 3--����֧��  4--�������
'        Bank ---- ��������
'        eBankID ---- ����֧��ƽ̨ID
'        Income_PayOut ---- ����  1--����  2--֧��  3--����
'        OrderFormID ---- ֧��ʱ�Ķ���ID
'        PaymentID ---- ����֧����֧����ID
'        Remark ---- ��ע
'        DateAndTime ---- ����ʱ��
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
