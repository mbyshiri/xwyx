<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Bankroll"   '其他权限

Private sqlBankroll, Querysql, strResultTips

strFileName = "Admin_Bankroll.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword
If Action = "outExcel" Then
    Call GetSqlStr
    Call outHead2
    Call outExcel
ElseIf Action = "ShowSearchForm" Then
    Call outhead
    Call ShowSearchForm
ElseIf Action = "ShowDetail" Then
    Call outhead
    Call ShowDetail
Else
    Call GetSqlStr
    Call outhead
    Call main
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End If
Response.Write "</body></html>"
Call CloseConn

Sub outhead()
    Response.Write "<html><head><title>资金明细查询</title>"
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
    Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
    Response.Write "</head>"
    Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Call ShowPageTitle("资 金 明 细 查 询", 10205)
    Response.Write "    <tr class='tdbg' height='30'> "
    Response.Write "  <form name='form1' action='Admin_Bankroll.asp' method='get'>"
    Response.Write "      <td>快速查找："
    Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
    Response.Write "          <option value='0'"
    If SearchType = 0 Then Response.Write " selected"
    Response.Write ">所有资金明细记录</option>"
    Response.Write "          <option value='1'"
    If SearchType = 1 Then Response.Write " selected"
    Response.Write ">最近10天内的新资金明细记录</option>"
    Response.Write "          <option value='2'"
    If SearchType = 2 Then Response.Write " selected"
    Response.Write ">最近一月内的新资金明细记录</option>"
    Response.Write "          <option value='3'"
    If SearchType = 3 Then Response.Write " selected"
    Response.Write ">所有收入记录</option>"
    Response.Write "          <option value='4'"
    If SearchType = 4 Then Response.Write " selected"
    Response.Write ">所有支出记录</option>"
    Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_Bankroll.asp'>资金明细首页</a></td>"
    Response.Write "  </form>"
    Response.Write "<form name='form2' method='post' action='Admin_Bankroll.asp'>"
    Response.Write "    <td>高级查询："
    Response.Write "      <select name='Field' id='Field'>"
    Response.Write "      <option value='ClientName' selected>客户姓名</option>"
    Response.Write "      <option value='UserName'>用户名</option>"
    Response.Write "      <option value='DateAndTime'>交易时间</option>"
    Response.Write "      <option value='BankName'>银行名称</option>"
    Response.Write "      </select>"
    Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='30'>"
    Response.Write "      <input type='submit' name='Submit2' value=' 查 询 '>"
    Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_Bankroll.asp?Action=ShowSearchForm'><b>复杂查询</b></a>"
    Response.Write " </td>"
    Response.Write "</form>"
    Response.Write "</table>"
    Response.Write "<br>"
End Sub

Sub outHead2()
    Response.Write "<html><head>" & vbCrLf
    Response.ContentType = "application/vnd.ms-excel" & vbCrLf
    Response.AddHeader "Content-Disposition", "attachment"
    Response.Write "<meta http-equiv=""Content-Language"" content=""zh-cn"">" & vbCrLf
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
    Response.Write "<title>资金明细表</title>" & vbCrLf
    Response.Write "<body>" & vbCrLf
End Sub

Sub ShowSearchForm()
    Call PopCalendarInit
    Response.Write "<form method='Get' name='formSearch' action='Admin_Bankroll.asp'>"
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "<tr class='title' align='center'><td colspan='6'>资 金 明 细 复 杂 查 询</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>ＩＤ范围：</td><td>起始ＩＤ<input type='text' name='BeginID'  size='10' value=''>&nbsp;终止ＩＤ<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>日期范围：</td><td>起始日期<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.formSearch.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;结束日期<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.formSearch.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>金额范围：</td><td><input type='text' name='MinMoney'  size='10' value=''> 至 <input type='text' name='MaxMoney'  size='10' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>客户名称：</td><td><input type='text' name='ClientName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>用户名：</td><td><input type='text' name='UserName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>交易方式：</td><td>"
    Response.Write "<input type='radio' name='MoneyType' value='1'>现金<input type='radio' name='MoneyType' value='2'>银行汇款<input type='radio' name='MoneyType' value='3'>在线支付<input type='radio' name='MoneyType' value='4'>虚拟货币<input type='radio' name='MoneyType' checked value='0'>所有方式"
    Response.Write "</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>资金收支方式：</td><td>"
    Response.Write "<input type='radio' name='Income_Payout' value='1'>收入<input type='radio' name='Income_Payout' value='2'>支出<input type='radio' name='Income_Payout' checked value='0'>所有方式"
    Response.Write "</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>备注/说明：</td><td><input type='text' name='Remark'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg' height='40' align='center'><td colspan='6'><input name='SearchType' type='hidden' id='SearchType' value='99'><input name='Action' type='hidden' value='Manage'><input type='submit' name='Submit'  value=' 查 询 '> "
    Response.Write "<input type='submit' name='Submit2'  value='导出到EXCEL' onclick=""document.formSearch.Action.value='outExcel';"">"
    Response.Write "</td></tr></table></form>"

End Sub

Sub GetSqlStr()
    sqlBankroll = "select B.*,C.ShortedForm as ClientName from PE_BankrollItem B left join PE_Client C on B.ClientID=C.ClientID "

    Querysql = " where 1=1"
    Select Case SearchType
        Case 0
            strResultTips = "所有资金明细记录"
        Case 1
            Querysql = Querysql & " And datediff(" & PE_DatePart_D & ",B.DateAndTime," & PE_Now & ")<10 "
            strResultTips = "最近10天内的新资金明细记录"
        Case 2
            Querysql = Querysql & " And datediff(" & PE_DatePart_M & ",B.DateAndTime," & PE_Now & ")<1 "
            strResultTips = "最近一月内的新资金明细记录"
        Case 3
            Querysql = Querysql & " And B.Money>0 "
            strResultTips = "所有收入记录"
        Case 4
            Querysql = Querysql & " And B.Money<0 "
            strResultTips = "所有支出记录"
        Case 10
            If Keyword = "" Then
                Querysql = Querysql & ""
                strResultTips = "所有资金明细记录"
            Else
                Select Case strField
                Case "ClientName"
                    Querysql = Querysql & " And C.ClientName like '%" & Keyword & "%' "
                    strResultTips = "客户姓名中含有“ <font color=red>" & Keyword & "</font> ”的资金明细记录"
                Case "UserName"
                    Querysql = Querysql & " And B.UserName like '%" & Keyword & "%' "
                    strResultTips = "用户名中含有“ <font color=red>" & Keyword & "</font> ”的资金明细记录"
                Case "BankName"
                    Querysql = Querysql & " And B.Bank='" & Keyword & "' "
                    strResultTips = "<font color=red>" & Keyword & "</font> 的资金明细记录"
                Case "DateAndTime"
                    If IsDate(Keyword) = True Then
                        Querysql = Querysql & " And DateDiff(" & PE_DatePart_D & ",B.DateAndTime,'" & Keyword & "')=0 "
                        strResultTips = "交易时间为 <font color=red>" & Keyword & "</font> 的资金明细记录"
                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>查询的交易时间格式不正确！</li>"
                    End If
                End Select
            End If
    Case 99
        strResultTips = "销售明细复杂查询结果"
        Dim BeginID, EndID, BeginDate, EndDate, MinMoney, MaxMoney, ClientName, UserName, MoneyType, Income_Payout, Remark
        BeginID = PE_CLng(Trim(Request("BeginID")))
        EndID = PE_CLng(Trim(Request("EndID")))
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        MinMoney = PE_CDbl(Trim(Request("MinMoney")))
        MaxMoney = PE_CDbl(Trim(Request("MaxMoney")))
        ClientName = ReplaceBadChar(Trim(Request("ClientName")))
        UserName = ReplaceBadChar(Trim(Request("UserName")))
        MoneyType = PE_CLng(Trim(Request("MoneyType")))
        Income_Payout = PE_CLng(Trim(Request("Income_Payout")))
        Remark = ReplaceBadChar(Trim(Request("Remark")))

        strFileName = "Admin_Bankroll.asp?SearchType=99&BeginID=" & BeginID & "&EndID=" & EndID & "&BeginDate=" & BeginDate & "&EndDate=" & EndDate & "&MinMoney=" & MinMoney & "&MaxMoney=" & MaxMoney
        strFileName = strFileName & "&ClientName=" & ClientName
        strFileName = strFileName & "&UserName=" & UserName & "&MoneyType=" & MoneyType
        strFileName = strFileName & "&Income_Payout=" & Income_Payout & "&Remark=" & Remark

        If BeginID > 0 Then
            Querysql = Querysql & " And B.ItemID>=" & BeginID
        End If
        If EndID > 0 Then
            Querysql = Querysql & " And B.ItemID<=" & EndID
        End If

        If BeginDate <> "" Then
            BeginDate = PE_CDate(BeginDate)
            If SystemDatabaseType = "SQL" Then
                Querysql = Querysql & " And B.DateAndTime>='" & BeginDate & "'"
            Else
                Querysql = Querysql & " And B.DateAndTime>=#" & BeginDate & "#"
            End If
        End If
        If EndDate <> "" Then
            EndDate = PE_CDate(EndDate)
            If SystemDatabaseType = "SQL" Then
                Querysql = Querysql & " And B.DateAndTime<='" & EndDate & "'"
            Else
                Querysql = Querysql & " And B.DateAndTime<=#" & EndDate & "#"
            End If
        End If
        If MinMoney > 0 Then
            Querysql = Querysql & " And abs(B.Money)>=" & MinMoney
        End If
        If MaxMoney > 0 Then
            Querysql = Querysql & " And abs(B.Money)<=" & MaxMoney
        End If
        If ClientName <> "" Then
            Querysql = Querysql & " And C.ClientName like '%" & ClientName & "%'"
        End If
        If UserName <> "" Then
            Querysql = Querysql & " And B.UserName like '%" & UserName & "%'"
        End If
        If MoneyType > 0 Then
            Querysql = Querysql & " And B.MoneyType = " & MoneyType & ""
        End If
        If Income_Payout > 0 Then
            If Income_Payout = 1 Then
                Querysql = Querysql & " And B.Money > 0"
            Else
                Querysql = Querysql & " And B.Money < 0"
            End If
        End If
        If Remark <> "" Then
            Querysql = Querysql & " And B.Remark like '%" & Remark & "%'"
        End If
    Case Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误的参数！</li>"
    End Select
End Sub


Sub main()
    Dim rsBankroll
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0

    sqlBankroll = sqlBankroll & Querysql & " order by B.ItemID desc"
    
    Call PopCalendarInit
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>您现在的位置：<a href='Admin_Bankroll.asp'>资金明细记录管理</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write strResultTips
    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>交易时间</td>"
    Response.Write "    <td width='80'>客户名称</td>"
    Response.Write "    <td width='80'>用户名</td>"
    Response.Write "    <td width='60'>交易方式</td>"
    Response.Write "    <td width='50'>币种</td>"
    Response.Write "    <td width='80'>收入金额</td>"
    Response.Write "    <td width='80'>支出金额</td>"
    Response.Write "    <td width='60'>银行名称</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "    <td width='40'>操作</td>"
    Response.Write "  </tr>"
    
    Set rsBankroll = Server.CreateObject("Adodb.RecordSet")
    rsBankroll.Open sqlBankroll, Conn, 1, 1
    If rsBankroll.BOF And rsBankroll.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何符合条件的资金记录！</td></tr>"
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
            Response.Write "    <td width='80' align='center'><a href='Admin_Client.asp?Action=Show&InfoType=3&ClientID=" & rsBankroll("ClientID") & "'>" & rsBankroll("ClientName") & "</a></td>"
            Response.Write "    <td width='80' align='center'><a href='Admin_User.asp?Action=Show&InfoType=1&UserName=" & rsBankroll("UserName") & "'>" & rsBankroll("UserName") & "</a></td>"
            Response.Write "    <td width='60' align='center'>"
            Select Case rsBankroll("MoneyType")
            Case 1
                Response.Write "现金"
            Case 2
                Response.Write "银行汇款"
            Case 3
                Response.Write "在线支付"
            Case 4
                Response.Write "虚拟货币"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='50' align='center'>"
            Select Case rsBankroll("CurrencyType")
            Case 1
                Response.Write "人民币"
            Case 2
                Response.Write "美元"
            Case 3
                Response.Write "其他"
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
            Response.Write "    <td align='center' width='40'><a href='Admin_Bankroll.asp?Action=ShowDetail&ItemID=" & rsBankroll("ItemID") & "'>查看</a></td>"
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsBankroll.MoveNext
        Loop
    End If
    rsBankroll.Close
    Set rsBankroll = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='5' align='right'>本页合计：</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncome, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayout), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='4'>&nbsp;</td>"
    Response.Write "  </tr>"

    Dim trs, TotalIncomeAll, TotalPayoutAll
    Set trs = Conn.Execute("select sum(Money) from PE_BankrollItem where Money>0")
    If IsNull(trs(0)) Then
        TotalIncomeAll = 0
    Else
        TotalIncomeAll = trs(0)
    End If
    Set trs = Nothing
    Set trs = Conn.Execute("select sum(Money) from PE_BankrollItem where Money<0")
    If IsNull(trs(0)) Then
        TotalPayoutAll = 0
    Else
        TotalPayoutAll = trs(0)
    End If
    Set trs = Nothing
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='5' align='right'>总计金额：</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncomeAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayoutAll), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='4' align='center'>资金余额：" & FormatNumber(TotalIncomeAll + TotalPayoutAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条资金明细记录", True)
End Sub

Sub ShowDetail()
    Dim rs, crs
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>您现在的位置：<a href='Admin_Bankroll.asp'>资金明细记录管理</a>&nbsp;&gt;&gt;&nbsp;资金明细详情"
    Response.Write "</td></tr></table>"
    Set rs = Conn.Execute("select * from PE_BankrollItem where ItemID=" & PE_CLng(Request("ItemID")) & "")
    If rs.BOF And rs.EOF Then
        Response.Write "<p align='center'>找不到指定的资金明细记录！</p>"
    Else
        
        Response.Write "    <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "        <tr class='title'>"
        Response.Write "          <td align='center' colspan='4'>查看资金明细记录详情</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>时间：</td>"
        Response.Write "          <td>" & rs("DateAndTime") & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>用户名：</td>"
        Response.Write "          <td><a href='Admin_User.asp?Action=Show&InfoType=1&UserName=" & rs("UserName") & "'>" & rs("UserName") & "</a></td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>客户名称：</td>"
        Response.Write "          <td>"
        Set crs = Conn.Execute("select ClientName from PE_Client where ClientID=" & PE_Clng(rs("ClientID")) & "")
        If Not (crs.BOF And crs.EOF) Then
            Response.Write "<a href='Admin_Client.asp?Action=Show&ClientID=" & rs("ClientID") & "'>" & crs("ClientName") & "</a>"
        End If
        crs.Close
        Set crs = Nothing
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg' valign='top'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>交易方式：</td>"
        Response.Write "    <td width='60' align='center'>"
        Select Case rs("MoneyType")
        Case 1
            Response.Write "现金"
        Case 2
            Response.Write "银行汇款"
        Case 3
            Response.Write "在线支付"
        Case 4
            Response.Write "虚拟货币"
        End Select
        Response.Write "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>币种：</td>"
        Response.Write "    <td width='50' align='center'>"
        Select Case rs("CurrencyType")
        Case 1
            Response.Write "人民币"
        Case 2
            Response.Write "美元"
        Case 3
            Response.Write "其他"
        End Select
        Response.Write "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>金额：</td>"
        Response.Write "          <td>" & FormatNumber(rs("Money"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>银行名称：</td>"
        Response.Write "          <td>"
        If rs("MoneyType") = 3 Then
            Response.Write GetPayOnlineProviderName(rs("eBankID"))
        Else
            Response.Write rs("Bank")
        End If
        Response.Write "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>关联订单：</td>"
        Response.Write "          <td>"
        Set crs = Conn.Execute("select OrderFormNum from PE_OrderForm where OrderFormID=" & rs("OrderFormID") & "")
        If Not (crs.BOF And crs.EOF) Then
            Response.Write "<a href='Admin_Order.asp?Action=ShowOrder&OrderFormID=" & rs("OrderFormID") & "'>" & crs("OrderFormNum") & "</a>"
        End If
        crs.Close
        Set crs = Nothing
        Response.Write "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>备注/说明：</td>"
        Response.Write "          <td colspan='3'>" & (rs("Remark")) & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>录入者：</td>"
        Response.Write "          <td>" & rs("Inputer") & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='15%' class='tdbg5' align='right'>IP:</td>"
        Response.Write "          <td>" & rs("IP") & "</td>"
        Response.Write "        </tr>"
        Response.Write "    </table>"
    End If
    Response.Write "<br><div align='center'><input type='button' name='button' value=' 返 回 ' onclick='javascript:history.go(-1)'></div>"
End Sub

Sub outExcel()
    Dim BeginDate, EndDate, BeginID, EndID, SelectType
    Dim Sqlout, Rsout

    If SearchType <> 99 Then
        SelectType = Trim(Request("SelectType"))
        BeginDate = Trim(Request("BeginDate"))
        If BeginDate = "" Then
            BeginDate = "1900-1-1"
        Else
            BeginDate = ReplaceBadChar(BeginDate)
        End If
        EndDate = Trim(Request("EndDate"))
        If EndDate = "" Then
            EndDate = FormatDateTime(Date, 2)
        Else
            EndDate = ReplaceBadChar(EndDate)
        End If
        BeginID = PE_CLng(Trim(Request("BeginID")))
        EndID = PE_CLng(Trim(Request("EndID")))
        If IsDate(BeginDate) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入正确的起始日期！</li>"
        End If
        If IsDate(EndDate) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入正确的结束日期！</li>"
        End If
        If FoundErr = True Then
            Exit Sub
        End If
        
        Sqlout = "select B.*,C.ShortedForm as ClientName from PE_BankrollItem B left join PE_Client C on B.ClientID=C.ClientID"
        Select Case SelectType
        Case "Date"
            If SystemDatabaseType = "SQL" Then
                Sqlout = Sqlout & " where B.DateAndTime Between '" & BeginDate & "' and '" & EndDate & "'"
            Else
                Sqlout = Sqlout & " where B.DateAndTime Between #" & BeginDate & "# and #" & EndDate & "#"
            End If
        Case "ID"
            If BeginID <> 0 And EndID <> 0 Then
                Sqlout = Sqlout & " where B.ItemID Between " & BeginID & " and " & EndID
            End If
        End Select
        Sqlout = Sqlout & " order by B.ItemID"
    Else
        Sqlout = sqlBankroll & Querysql & " order by B.ItemID desc"
    End If

    Set Rsout = Conn.Execute(Sqlout)
    If Rsout.BOF And Rsout.EOF Then
        Response.Write "无指定的明细单！"
    Else
        Response.Write "<table border=""0"" cellspacing=""1"" style=""border-collapse: collapse;table-layout:fixed"" id=""AutoNumber1"" height=""32"">" & vbCrLf
        Response.Write "<tr>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>客户姓名</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>交易时间</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>交易方式</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>币种</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>收入金额</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>支出金额</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>银行名称</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>备注/说明</b></span></td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Do While Not Rsout.EOF
            Response.Write "<tr>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & Rsout("ClientName") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & FormatDateTime(Rsout("DateAndTime"), 2) & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            Select Case Rsout("MoneyType")
            Case 1
                Response.Write "现金"
            Case 2
                Response.Write "银行汇款"
            Case 3
                Response.Write "资金明细"
            Case 4
                Response.Write "虚拟货币"
            End Select
            Response.Write "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            Select Case Rsout("CurrencyType")
            Case 1
                Response.Write "人民币"
            Case 2
                Response.Write "美元"
            Case 3
                Response.Write "其他"
            End Select
            Response.Write "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            If Rsout("Money") > 0 Then Response.Write FormatNumber(Rsout("Money"), 2, vbTrue, vbFalse, vbTrue)
            Response.Write "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            If Rsout("Money") < 0 Then Response.Write FormatNumber(Abs(Rsout("Money")), 2, vbTrue, vbFalse, vbTrue)
            Response.Write "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            If Rsout("MoneyType") = 3 Then
                Select Case Rsout("eBankID")
                Case 1
                    Response.Write "网银在线"
                Case 2
                    Response.Write "中国在线支付网"
                Case 3
                    Response.Write "上海环迅"
                End Select
            Else
                Response.Write Rsout("Bank")
            End If
            Response.Write "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & Rsout("Remark") & "</span></td>" & vbCrLf
            Response.Write "</tr>" & vbCrLf
            Rsout.MoveNext
        Loop
    End If
    Rsout.Close
    Set Rsout = Nothing
End Sub
%>
