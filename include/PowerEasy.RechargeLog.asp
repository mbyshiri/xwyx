<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'**************************************************
'��������AddRechargeLog
'��  �ã����һ����Ч�ڳ�ֵ��¼
'��  ����AdminName ----����Ա����
'        strUserName ---- �û�����
'        iValidNum ----��Ч��
'        iValidUnit ----��Ч�ڵ�λ
'        iIncome_Payout ---- ������ϸ����  1--����  2--֧��
'        strRemark ---- ��ע/˵��
'**************************************************
Sub AddRechargeLog(AdminName, strUserName, iValidNum, iValidUnit, iIncome_Payout, strRemark)
    Dim rsRechargeLog, sqlRechargeLog
    sqlRechargeLog = "select top 1 * from PE_RechargeLog"
    Set rsRechargeLog = Server.CreateObject("adodb.recordset")
    rsRechargeLog.Open sqlRechargeLog, Conn, 1, 3
    rsRechargeLog.addnew
    rsRechargeLog("UserName") = strUserName
    rsRechargeLog("ValidNum") = iValidNum
    rsRechargeLog("ValidUnit") = iValidUnit
    rsRechargeLog("Income_PayOut") = iIncome_Payout
    rsRechargeLog("Remark") = strRemark
    rsRechargeLog("LogTime") = Now()
    rsRechargeLog("IP") = UserTrueIP
    rsRechargeLog("Inputer") = AdminName
    rsRechargeLog.Update
    rsRechargeLog.Close
    Set rsRechargeLog = Nothing
End Sub
%>
