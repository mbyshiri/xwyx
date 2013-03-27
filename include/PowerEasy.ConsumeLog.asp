<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'**************************************************
'方法名：AddConsumeLog
'作  用：添加一条消费明细
'参  数：AdminName ----操作者,如无则为System
'        ModuleType---- 频道类型
'        strUserName ----用户名称
'        iInfoID ---- 项目ID,被扣除点数的信息ID
'        iPoint ---- 点数
'        Income_PayOut ---- 消费明细类型  1--收入  2--支出
'        strRemark ---- 备注/说明
'**************************************************
Sub AddConsumeLog(AdminName, ModuleType, strUserName, iInfoID, iPoint, Income_PayOut, strRemark)
    Dim rsConsumeLog, sqlConsumeLog
    sqlConsumeLog = "select top 1 * from PE_ConsumeLog"
    Set rsConsumeLog = Server.CreateObject("adodb.recordset")
    rsConsumeLog.Open sqlConsumeLog, Conn, 1, 3
    rsConsumeLog.addnew
    rsConsumeLog("UserName") = strUserName
    rsConsumeLog("ModuleType") = ModuleType
    rsConsumeLog("InfoID") = iInfoID
    rsConsumeLog("Point") = iPoint
    rsConsumeLog("Income_PayOut") = Income_PayOut
    rsConsumeLog("Remark") = strRemark
    rsConsumeLog("IP") = UserTrueIP
    rsConsumeLog("LogTime") = Now()
    rsConsumeLog("Times") = 1
    rsConsumeLog("Inputer") = AdminName
    rsConsumeLog.Update
    rsConsumeLog.Close
    Set rsConsumeLog = Nothing
End Sub
%>
