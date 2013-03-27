<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Payment"   '其他权限

strFileName = "Admin_Payment.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword


Response.Write "<html><head><title>在线支付记录管理</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
Call ShowJS_Main("在线支付记录")
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
Call ShowPageTitle("在 线 支 付 记 录 管 理", 10204)
Response.Write "    <tr class='tdbg' height='30'> "
Response.Write "  <form name='form1' action='Admin_Payment.asp' method='get'>"
Response.Write "      <td>快速查找："
Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
Response.Write "          <option value='0'"
If SearchType = 0 Then Response.Write " selected"
Response.Write ">所有在线支付记录</option>"
Response.Write "          <option value='1'"
If SearchType = 1 Then Response.Write " selected"
Response.Write ">最近10天内的新在线支付记录</option>"
Response.Write "          <option value='2'"
If SearchType = 2 Then Response.Write " selected"
Response.Write ">最近一月内的新在线支付记录</option>"
Response.Write "          <option value='3'"
If SearchType = 3 Then Response.Write " selected"
Response.Write ">未提交的在线支付记录</option>"
Response.Write "          <option value='4'"
If SearchType = 4 Then Response.Write " selected"
Response.Write ">未成功的在线支付记录</option>"
Response.Write "          <option value='5'"
If SearchType = 5 Then Response.Write " selected"
Response.Write ">支付成功的在线支付记录</option>"
Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_Payment.asp'>在线支付记录首页</a></td>"
Response.Write "  </form>"
Response.Write "<form name='form2' method='post' action='Admin_Payment.asp'>"
Response.Write "    <td>高级查询："
Response.Write "      <select name='Field' id='Field'>"
Response.Write "      <option value='PaymentNum'>在线支付记录编号</option>"
Response.Write "      <option value='UserName'>用户名</option>"
Response.Write "      <option value='PayTime'>支付时间</option>"
Response.Write "      </select>"
Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='30'>"
Response.Write "      <input type='submit' name='Submit2' value=' 查 询 '>"
Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'>"
Response.Write " </td>"
Response.Write "</form>"
Response.Write "</table>"
Response.Write "<br>"
If Action = "Cancel" Then
    Call DelPayment
ElseIf Action = "Success" Then
    Call PaySuccess
Else
    Call main
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsPaymentList, sqlPaymentList, Querysql
    Dim TotalMoneyPay, TotalMoneyTrue
    TotalMoneyPay = 0
    TotalMoneyTrue = 0
    
    sqlPaymentList = "select top " & MaxPerPage & " * from PE_Payment "
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>您现在的位置：<a href='Admin_Payment.asp'>在线支付记录管理</a>&nbsp;&gt;&gt;&nbsp;"

    Querysql = Querysql & " where 1=1 "
    Select Case SearchType
        Case 0
            Response.Write "所有在线支付记录"
        Case 1
            Querysql = Querysql & " and datediff(" & PE_DatePart_D & ",PayTime," & PE_Now & ")<10"
            Response.Write "最近10天内的新在线支付记录"
        Case 2
            Querysql = Querysql & " and datediff(" & PE_DatePart_M & ",PayTime," & PE_Now & ")<1"
            Response.Write "最近一月内的新在线支付记录"
        Case 3
            Querysql = Querysql & " and Status=1"
            Response.Write "未提交的在线支付记录"
        Case 4
            Querysql = Querysql & " and Status=2"
            Response.Write "未成功的在线支付记录"
        Case 5
            Querysql = Querysql & " and Status=3"
            Response.Write "支付成功的在线支付记录"
        Case 10
            If Keyword = "" Then
                Response.Write "所有在线支付记录"
            Else
                Select Case strField
                Case "PaymentNum"
                    Querysql = Querysql & " and PaymentNum like '%" & Keyword & "%'"
                    Response.Write "在线支付记录编号中含有“ <font color=red> " & Keyword & " </font> ”的在线支付记录"
                Case "UserName"
                    Querysql = Querysql & " and UserName like '%" & Keyword & "%'"
                    Response.Write "用户名中含有“ <font color=red>" & Keyword & "</font> ”的在线支付记录"
                Case "PayTime"
                    If IsDate(Keyword) = True Then
                        If SystemDatabaseType = "SQL" Then
                            Querysql = Querysql & " and PayTime='" & Keyword & "'"
                        Else
                            Querysql = Querysql & " and PayTime=#" & Keyword & "#"
                        End If
                        Response.Write "支付时间为 <font color=red>" & Keyword & "</font> 的在线支付记录"
                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>查询的支付时间格式不正确！</li>"
                    End If
                End Select
            End If
        Case Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>错误的参数！</li>"
    End Select

    totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Payment " & Querysql)(0))
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
        Querysql = Querysql & " and PaymentID < (select min(PaymentID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " PaymentID from PE_Payment " & Querysql & " order by PaymentID desc) as QueryPayment) "
    End If
    sqlPaymentList = sqlPaymentList & Querysql & " order by PaymentID desc"

    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Payment.asp' onsubmit=""return confirm('确定要删除选定的在线支付记录吗？');"">"
    Response.Write "     <td>"
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'>选中</td>"
    Response.Write "    <td width='80'>支付序号</td>"
    Response.Write "    <td width='60'>用户名</td>"
    Response.Write "    <td width='70'>支付平台</td>"
    Response.Write "    <td width='120'>交易时间</td>"
    Response.Write "    <td width='70'>汇款金额</td>"
    Response.Write "    <td width='70'>实际转账<br>金额</td>"
    Response.Write "    <td width='60'>交易状态</td>"
    Response.Write "    <td width='70'>银行信息</td>"
    Response.Write "    <td>备注</td>"
    Response.Write "    <td>操作</td>"
    Response.Write "  </tr>"
    
    Set rsPaymentList = Server.CreateObject("Adodb.RecordSet")
    rsPaymentList.Open sqlPaymentList, Conn, 1, 1
    If rsPaymentList.BOF And rsPaymentList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何符合条件的在线支付单！</td></tr>"
    Else
        Dim i
        i = 0
        Do While Not rsPaymentList.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='30' align='center'><input name='PaymentID' type='checkbox' onclick='unselectall()' id='PaymentID' value='" & rsPaymentList("PaymentID") & "'></td>"
            Response.Write "    <td width='80' align='center'>" & rsPaymentList("PaymentNum") & "</td>"
            Response.Write "    <td width='60' align='center'><a href='Admin_User.asp?Action=Show&UserName=" & rsPaymentList("UserName") & "'>" & rsPaymentList("UserName") & "</a></td>"
            Response.Write "    <td width='70' align='center'>" & GetPayOnlineProviderName(rsPaymentList("eBankID")) & "</td>"
            Response.Write "    <td width='120' align='center'>" & rsPaymentList("PayTime") & "</td>"
            Response.Write "    <td width='70' align='right'>" & FormatNumber(rsPaymentList("MoneyPay"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='70' align='right'>" & FormatNumber(rsPaymentList("MoneyTrue"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
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
            Response.Write "    <td align='center'>"
            If rsPaymentList("Status") = 1 Then
                Response.Write "<a href='Admin_Payment.asp?Action=Cancel&PaymentID=" & rsPaymentList("PaymentID") & "' onclick=""return confirm('确定要删除这条在线支付记录吗？');"">取消</a> "
                Response.Write "<a href='Admin_Payment.asp?Action=Success&PaymentID=" & rsPaymentList("PaymentID") & "' onclick=""return confirm('确定这条在线支付记录已经支付成功了吗？');"">成功</a>"
            End If
            Response.Write "</td>"
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
    Response.Write "    <td width='70' align='right'>" & FormatNumber(TotalMoneyPay, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td width='70' align='right'>" & FormatNumber(TotalMoneyTrue, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='4' align='center'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='220' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> 选中本页显示的所有在线支付记录</td>"
    Response.Write "    <td width='560'> <input name='Action' type='hidden' id='Action' value='Cancel'> <input type='submit' name='Submit' value='删除选定的在线支付记录'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条在线支付记录", True)
End Sub


Sub DelPayment()
    Dim PaymentID
    Dim rsPayment, sqlPayment
    PaymentID = Trim(Request("PaymentID"))
    If PaymentID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定支付单ID！</li>"
        Exit Sub
    Else
        If IsValidID(PaymentID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定正确的支付单ID！</li>"
            Exit Sub
        End If
    End If
    
    sqlPayment = "select * from PE_Payment where PaymentID in (" & PaymentID & ")"
    Set rsPayment = Server.CreateObject("Adodb.RecordSet")
    rsPayment.Open sqlPayment, Conn, 1, 3
    Do While Not rsPayment.EOF
        If rsPayment("Status") = 1 Then
            rsPayment.Delete
            rsPayment.Update
        End If
        rsPayment.MoveNext
    Loop
    rsPayment.Close
    Set rsPayment = Nothing
    Call CloseConn
    Call WriteSuccessMsg("成功删除选定的在线支付记录", "Admin_Payment.asp")
End Sub

Sub PaySuccess()
    Dim PaymentID, PaymentNum, UserName, OrderFormID, MoneyReceipt, eBankID, MoneyPayout, ClientID
    Dim rsPayment, sqlPayment, trs, rsUser
    PaymentID = Trim(Request("PaymentID"))
    ClientID = 0
    If PaymentID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定支付单ID！</li>"
        Exit Sub
    Else
        PaymentID = PE_CLng(PaymentID)
    End If
    
    sqlPayment = "select * from PE_Payment where PaymentID=" & PaymentID & ""
    Set rsPayment = Server.CreateObject("Adodb.RecordSet")
    rsPayment.Open sqlPayment, Conn, 1, 3
    If rsPayment.BOF And rsPayment.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的订单！</li>"
        rsPayment.Close
        Set rsPayment = Nothing
        Exit Sub
    End If
    If rsPayment("Status") > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>此支付单已经提交给银行！</li>"
    Else
        PaymentNum = rsPayment("PaymentNum")
        UserName = rsPayment("UserName")
        OrderFormID = rsPayment("OrderFormID")
        MoneyReceipt = rsPayment("MoneyPay")
        eBankID = rsPayment("eBankID")
        rsPayment("Status") = 3
        rsPayment("eBankInfo") = "支付完成"
        rsPayment("Remark") = "未知"
        rsPayment.Update
    End If
    rsPayment.Close
    Set rsPayment = Nothing

    Set rsUser = Conn.Execute("select ClientID from PE_User where UserName='" & UserName & "'")
    If Not (rsUser.EOF And rsUser.BOF) Then ClientID = rsUser(0)
      
    If FoundErr = True Then Exit Sub
    
    '检查是否已经有记录，若已经有，跳过写入数据库的操作
    Set trs = Conn.Execute("select * from PE_BankrollItem where PaymentID=" & PaymentID & "")
    If Not (trs.BOF And trs.EOF) Then
        ErrMsg = ErrMsg & "<li>资金明细中已经有相关记录！</li>"
        FoundErr = True
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub
    
    '向资金余额中添加金额
    Conn.Execute ("update PE_User set Balance=Balance+" & MoneyReceipt & " where UserName='" & UserName & "'")
    
    ' 向资金明细表中添加收入记录
    Call AddBankrollItem("", UserName, ClientID, MoneyReceipt, 3, "", eBankID, 1, 0, PaymentID, "在线支付单号：" & PaymentNum, Now())
        
    If OrderFormID > 0 Then
        Dim rsOrder
        Set rsOrder = Server.CreateObject("adodb.recordset")
        rsOrder.Open "select * from PE_OrderForm where OrderFormID=" & OrderFormID & "", Conn, 1, 3
        If Not (rsOrder.BOF And rsOrder.EOF) Then
            If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") Then
                If rsOrder("MoneyTotal") - rsOrder("MoneyReceipt") > MoneyReceipt Then
                    MoneyPayout = MoneyReceipt
                    rsOrder("MoneyReceipt") = rsOrder("MoneyReceipt") + MoneyReceipt
                Else
                    MoneyPayout = rsOrder("MoneyTotal") - rsOrder("MoneyReceipt")
                    rsOrder("MoneyReceipt") = rsOrder("MoneyTotal")
                End If
                rsOrder.Update
                '向资金明细表中添加支付记录
                Call AddBankrollItem("", UserName, ClientID, MoneyPayout, 4, "", 0, 2, OrderFormID, 0, "支付订单费用，订单号：" & rsOrder("OrderFormNum"), Now())
                
                '从资金余额中扣除支付费用
                Conn.Execute ("update PE_User set Balance=Balance-" & MoneyPayout & " where UserName='" & UserName & "'")
            End If
        End If
        rsOrder.Close
        Set rsOrder = Nothing
    End If
    Call CloseConn
    Call WriteSuccessMsg("在线支付成功", "Admin_Payment.asp")
    If ErrMsg <> "" Then
        FoundErr = True
    End If
End Sub
%>
