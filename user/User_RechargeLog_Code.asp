<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub Main()
    strFileName = "User_RechargeLog.asp"
    
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>ʱ��</td>"
    Response.Write "    <td width='100'>������Ч��</td>"
    Response.Write "    <td width='100'>������Ч��</td>"
    Response.Write "    <td width='40'>ժҪ</td>"
    Response.Write "    <td>��ע/˵��</td>"
    Response.Write "  </tr>"
    
    Dim rsRechargeLog, sqlRechargeLog
    sqlRechargeLog = "select * from PE_RechargeLog where UserName='" & UserName & "' order by LogID desc"
    
    Set rsRechargeLog = Server.CreateObject("Adodb.RecordSet")
    rsRechargeLog.Open sqlRechargeLog, Conn, 1, 1
    If rsRechargeLog.BOF And rsRechargeLog.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κ���Ч����ϸ��¼��</td></tr>"
    Else
        totalPut = rsRechargeLog.RecordCount
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
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsRechargeLog.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim i
        i = 0
        Do While Not rsRechargeLog.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='120' align='center'>" & rsRechargeLog("LogTime") & "</td>"
            Response.Write "    <td width='80' align='right'>"
            If rsRechargeLog("Income_Payout") = 1 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='80' align='right'>"
            If rsRechargeLog("Income_Payout") = 2 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='40' align='center'>"
            Select Case rsRechargeLog("Income_Payout")
            Case 1
                Response.Write "<font color='blue'>����</font>"
            Case 2
                Response.Write "<font color='green'>����</font>"
            Case Else
                Response.Write "����"
            End Select
            Response.Write "</td>"
            Response.Write "    <td align='left'>" & rsRechargeLog("Remark") & "</td>"
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsRechargeLog.MoveNext
        Loop
    End If
    rsRechargeLog.Close
    Set rsRechargeLog = Nothing
    
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����Ч����ϸ��¼", True)
End Sub
%>
