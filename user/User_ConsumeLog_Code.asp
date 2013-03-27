<!--#include file="CommonCode.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub Main()
    Dim rsConsumeLog, sqlConsumeLog
    Dim TotalIncome, TotalPayout
   strFileName =   "User_ConsumeLog.asp"
    TotalIncome = 0
    TotalPayout = 0

    Select Case ShowType
    Case 0
        sqlConsumeLog = "select * from PE_ConsumeLog where UserName='" & UserName & "' order by LogID desc"
    Case 1
        sqlConsumeLog = "select * from PE_ConsumeLog where UserName='" & UserName & "' and Income_Payout=1 order by LogID desc"
    Case 2
        sqlConsumeLog = "select * from PE_ConsumeLog where UserName='" & UserName & "' and Income_Payout=2 order by LogID desc"
    End Select
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>" & PointName & "ʱ��</td>"
    Response.Write "    <td width='100'>IP��ַ</td>"
    Response.Write "    <td width='60'>����" & PointName & "��</td>"
    Response.Write "    <td width='60'>֧��" & PointName & "��</td>"
    Response.Write "    <td width='40'>ժҪ</td>"
    Response.Write "    <td width='50'>�ظ�����</td>"
    Response.Write "    <td>��ע/˵��</td>"
    Response.Write "  </tr>"

    Set rsConsumeLog = Server.CreateObject("Adodb.RecordSet")
    rsConsumeLog.Open sqlConsumeLog, Conn, 1, 1
    If rsConsumeLog.BOF And rsConsumeLog.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κ�" & PointName & "��¼��</td></tr>"
    Else
        totalPut = rsConsumeLog.RecordCount
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
                rsConsumeLog.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim i
        i = 0
        Do While Not rsConsumeLog.EOF
            If rsConsumeLog("Income_Payout") = 1 Or rsConsumeLog("Income_Payout") = 3 Then
                TotalIncome = TotalIncome + rsConsumeLog("Point")
            Else
                TotalPayout = TotalPayout + rsConsumeLog("Point")
            End If

            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='120' align='center'>" & rsConsumeLog("LogTime") & "</td>"
            Response.Write "    <td width='100' align='center'>" & rsConsumeLog("IP") & "</td>"
            Response.Write "    <td width='60' align='right'>"
            If rsConsumeLog("Income_Payout") = 1 Then Response.Write rsConsumeLog("Point")
            Response.Write "</td>"
            Response.Write "    <td width='60' align='right'>"
            If rsConsumeLog("Income_Payout") = 2 Then Response.Write rsConsumeLog("Point")
            Response.Write "</td>"
            Response.Write "    <td width='40' align='center'>"
            Select Case rsConsumeLog("Income_Payout")
            Case 1
                Response.Write "<font color='blue'>����</font>"
            Case 2
                Response.Write "<font color='green'>֧��</font>"
            Case Else
                Response.Write "����"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='50' align='center'>" & rsConsumeLog("Times") & "</td>"
            Response.Write "    <td align='left'>" & rsConsumeLog("Remark") & "</td>"
            Response.Write "  </tr>"

            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsConsumeLog.MoveNext
        Loop
    End If
    rsConsumeLog.Close
    Set rsConsumeLog = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='2' align='right'>��ҳ�ϼƣ�</td>"
    Response.Write "    <td align='right'>" & TotalIncome & "</td>"
    Response.Write "    <td align='right'>" & TotalPayout & "</td>"
    Response.Write "    <td colspan='3'>&nbsp;</td>"
    Response.Write "  </tr>"

    Dim trs, TotalIncomeAll, TotalPayoutAll
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where Income_Payout=1 and UserName='" & UserName & "'")
    If IsNull(trs(0)) Then
        TotalIncomeAll = 0
    Else
        TotalIncomeAll = trs(0)
    End If
    Set trs = Nothing
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where Income_Payout=2 and UserName='" & UserName & "'")
    If IsNull(trs(0)) Then
        TotalPayoutAll = 0
    Else
        TotalPayoutAll = trs(0)
    End If
    Set trs = Nothing
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='2' align='right'>�ܼ�" & PointName & "����</td>"
    Response.Write "    <td align='right'>" & TotalIncomeAll & "</td>"
    Response.Write "    <td align='right'>" & TotalPayoutAll & "</td>"
    Response.Write "    <td colspan='3' align='center'>" & PointName & "����" & TotalIncomeAll - TotalPayoutAll & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��" & PointName & "��ϸ��¼", True)
End Sub
%>
