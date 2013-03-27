<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'**************************************************
'方法名：AddRechargeLog
'作  用：添加一条有效期充值记录
'参  数：AdminName ----管理员名称
'        strUserName ---- 用户名称
'        iValidNum ----有效期
'        iValidUnit ----有效期单位
'        iIncome_Payout ---- 消费明细类型  1--收入  2--支出
'        strRemark ---- 备注/说明
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
