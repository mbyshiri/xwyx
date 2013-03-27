<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub Main()
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='80'>支付序号</td>"
    Response.Write "    <td width='70'>支付平台</td>"
    Response.Write "    <td width='120'>交易时间</td>"
    Response.Write "    <td width='80'>汇款金额</td>"
    Response.Write "    <td width='80'>实际转账金额</td>"
    Response.Write "    <td width='60'>交易状态</td>"
    Response.Write "    <td width='70'>银行信息</td>"
    Response.Write "    <td>备注</td>"
    Response.Write "  </tr>"

    Dim rsPaymentList, sqlPaymentList
    Dim TotalMoneyPay, TotalMoneyTrue
    TotalMoneyPay = 0
    TotalMoneyTrue = 0

    sqlPaymentList = "select * from PE_Payment where UserName='" & UserName & "' order by PaymentID desc"
    Set rsPaymentList = Server.CreateObject("Adodb.RecordSet")
    rsPaymentList.Open sqlPaymentList, Conn, 1, 1
    If rsPaymentList.BOF And rsPaymentList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何在线支付单！</td></tr>"
    Else
        totalPut = rsPaymentList.RecordCount
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
                rsPaymentList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        
        Dim i
        i = 0
        Do While Not rsPaymentList.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='80' align='center'>" & rsPaymentList("PaymentNum") & "</td>"
            Response.Write "    <td width='70' align='center'>" & GetPayOnlineProviderName(rsPaymentList("eBankID")) & "</td>"
            Response.Write "    <td width='120' align='center'>" & rsPaymentList("PayTime") & "</td>"
            Response.Write "    <td width='80' align='right'>" & FormatNumber(rsPaymentList("MoneyPay"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='80' align='right'>" & FormatNumber(rsPaymentList("MoneyTrue"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='60' align='center'>"
            If rsPaymentList("eBankID") <> 8 Then
                Select Case rsPaymentList("Status")
                Case 1
                    Response.Write "未提交"
                Case 2
                    Response.Write "已经提交，但未成功"
                Case 3
                    Response.Write "支付成功"
                End Select
            Else
                Select Case rsPaymentList("Status")
                Case 1
                    Response.Write "等待买家付款"
                Case 2
                    Response.Write "买家已付款"
                Case 3
                    Response.Write "交易成功"
                Case 4
                    Response.Write "卖家已发货，等待买家确认收货"
                End Select
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='70' align='center'>" & rsPaymentList("eBankInfo") & "</td>"
            Response.Write "    <td>" & rsPaymentList("Remark") & "</td>"
            Response.Write "  </tr>"
            TotalMoneyPay = TotalMoneyPay + rsPaymentList("MoneyPay")
            TotalMoneyTrue = TotalMoneyTrue + rsPaymentList("MoneyTrue")
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsPaymentList.MoveNext
        Loop
    End If
    rsPaymentList.Close
    Set rsPaymentList = Nothing
        
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='5' align='right'>合计金额：</td>"
    Response.Write "    <td width='80' align='right'>" & FormatNumber(TotalMoneyPay, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td width='80' align='right'>" & FormatNumber(TotalMoneyTrue, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='4' align='center'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条在线支付记录", True)
End Sub
%>
