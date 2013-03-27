<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub Main()
    strFileName = "User_Bankroll.asp?ShowType=" & ShowType

    Dim rsBankroll, sqlBankroll
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0

    Select Case ShowType
    Case 0
        sqlBankroll = "select * from PE_BankrollItem where UserName='" & UserName & "' order by ItemID desc"
    Case 1
        sqlBankroll = "select * from PE_BankrollItem where UserName='" & UserName & "' and Money>0 order by ItemID desc"
    Case 2
        sqlBankroll = "select * from PE_BankrollItem where UserName='" & UserName & "' and Money<0 order by ItemID desc"
    End Select

    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>����ʱ��</td>"
    Response.Write "    <td width='60'>���׷�ʽ</td>"
    Response.Write "    <td width='50'>����</td>"
    Response.Write "    <td width='80'>������</td>"
    Response.Write "    <td width='80'>֧�����</td>"
    Response.Write "    <td width='60'>��������</td>"
    Response.Write "    <td>��ע/˵��</td>"
    Response.Write "  </tr>"

    Set rsBankroll = Server.CreateObject("Adodb.RecordSet")
    rsBankroll.Open sqlBankroll, Conn, 1, 1
    If rsBankroll.BOF And rsBankroll.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κη����������ʽ��¼��</td></tr>"
    Else
        totalPut = rsBankroll.RecordCount
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
                rsBankroll.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim i
        i = 0
        Do While Not rsBankroll.EOF
            If rsBankroll("Money") > 0 Then
                TotalIncome = TotalIncome + rsBankroll("Money")
            Else
                TotalPayout = TotalPayout + rsBankroll("Money")
            End If

            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='120' align='center'>" & rsBankroll("DateAndTime") & "</td>"
            Response.Write "    <td width='60' align='center'>"
            Select Case rsBankroll("MoneyType")
            Case 1
                Response.Write "�ֽ�"
            Case 2
                Response.Write "���л��"
            Case 3
                Response.Write "�ʽ���ϸ"
            Case 4
                Response.Write "�������"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='50' align='center'>"
            Select Case rsBankroll("CurrencyType")
            Case 1
                Response.Write "�����"
            Case 2
                Response.Write "��Ԫ"
            Case 3
                Response.Write "����"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='80' align='right'>"
            If rsBankroll("Money") > 0 Then Response.Write FormatNumber(rsBankroll("Money"), 2, vbTrue, vbFalse, vbTrue)
            Response.Write "</td>"
            Response.Write "    <td width='80' align='right'>"
            If rsBankroll("Money") < 0 Then Response.Write FormatNumber(Abs(rsBankroll("Money")), 2, vbTrue, vbFalse, vbTrue)
            Response.Write "</td>"
            Response.Write "    <td align='center' width='60'>"
            If rsBankroll("MoneyType") = 3 Then
                Response.Write GetPayOnlineProviderName(rsBankroll("eBankID"))
            Else
                Response.Write rsBankroll("Bank")
            End If
            Response.Write "</td>"
            Response.Write "    <td align='center'>" & rsBankroll("Remark") & "</td>"
            Response.Write "  </tr>"

            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsBankroll.MoveNext
        Loop
    End If
    rsBankroll.Close
    Set rsBankroll = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='4' align='right'>��ҳ�ϼƣ�</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncome, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayout), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='3'>&nbsp;</td>"
    Response.Write "  </tr>"

    Dim trs, TotalIncomeAll, TotalPayoutAll
    Set trs = Conn.Execute("select sum(Money) from PE_BankrollItem where Money>0 and UserName='" & UserName & "'")
    If IsNull(trs(0)) Then
        TotalIncomeAll = 0
    Else
        TotalIncomeAll = trs(0)
    End If
    Set trs = Nothing
    Set trs = Conn.Execute("select sum(Money) from PE_BankrollItem where Money<0 and UserName='" & UserName & "'")
    If IsNull(trs(0)) Then
        TotalPayoutAll = 0
    Else
        TotalPayoutAll = trs(0)
    End If
    Set trs = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='4' align='right'>�ܼƽ�</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncomeAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayoutAll), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='3' align='center'>�ʽ���" & FormatNumber(TotalIncomeAll + TotalPayoutAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "���ʽ���ϸ��¼", True)
End Sub
%>
