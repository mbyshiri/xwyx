<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'**************************************************
'��������AddConsumeLog
'��  �ã����һ��������ϸ
'��  ����AdminName ----������,������ΪSystem
'        ModuleType---- Ƶ������
'        strUserName ----�û�����
'        iInfoID ---- ��ĿID,���۳���������ϢID
'        iPoint ---- ����
'        Income_PayOut ---- ������ϸ����  1--����  2--֧��
'        strRemark ---- ��ע/˵��
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
