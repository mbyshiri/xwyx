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
Const PurviewLevel_Others = "RechargeLog"   '其他权限

strFileName = "Admin_RechargeLog.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword

Response.Write "<html><head><title>有效期明细查询</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
Call ShowPageTitle("有 效 期 明 细 查 询", 10045)
Response.Write "    <tr class='tdbg' height='30'> "
Response.Write "  <form name='form1' action='Admin_RechargeLog.asp' method='get'>"
Response.Write "      <td>快速查找："
Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
Response.Write "          <option value='0'"
If SearchType = 0 Then Response.Write " selected"
Response.Write ">所有有效期明细记录</option>"
Response.Write "          <option value='1'"
If SearchType = 1 Then Response.Write " selected"
Response.Write ">最近10天内的新有效期明细记录</option>"
Response.Write "          <option value='2'"
If SearchType = 2 Then Response.Write " selected"
Response.Write ">最近一月内的新有效期明细记录</option>"
Response.Write "          <option value='3'"
If SearchType = 3 Then Response.Write " selected"
Response.Write ">所有收入记录</option>"
Response.Write "          <option value='4'"
If SearchType = 4 Then Response.Write " selected"
Response.Write ">所有支出记录</option>"
Response.Write "          <option value='5'"
If SearchType = 5 Then Response.Write " selected"
Response.Write ">所有非开户记录</option>"
Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_RechargeLog.asp'>有效期明细首页</a></td>"
Response.Write "  </form>"
Response.Write "<form name='form2' method='post' action='Admin_RechargeLog.asp'>"
Response.Write "    <td>高级查询："
Response.Write "      <select name='Field' id='Field'>"
Response.Write "      <option value='UserName'>用户名</option>"
Response.Write "      <option value='LogTime'>时间</option>"
Response.Write "      </select>"
Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='30'>"
Response.Write "      <input type='submit' name='Submit2' value=' 查 询 '>"
Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'>"
Response.Write " </td>"
Response.Write "</form>"
Response.Write "</table>"
Response.Write "<br>"

If Action = "Del" Then
    Call Del
Else
    Call main
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsRechargeLog, sqlRechargeLog
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    
    sqlRechargeLog = "select * from PE_RechargeLog "
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>您现在的位置：<a href='Admin_Bankroll.asp'>有效期明细记录管理</a>&nbsp;&gt;&gt;&nbsp;"
    Select Case SearchType
        Case 0
            sqlRechargeLog = sqlRechargeLog & " order by LogID desc"
            Response.Write "所有有效期明细记录"
        Case 1
            sqlRechargeLog = sqlRechargeLog & " where datediff(" & PE_DatePart_D & ",LogTime," & PE_Now & ")<10 order by LogID desc"
            Response.Write "最近10天内的新有效期明细记录"
        Case 2
            sqlRechargeLog = sqlRechargeLog & " where datediff(" & PE_DatePart_M & ",LogTime," & PE_Now & ")<1 order by LogID desc"
            Response.Write "最近一月内的新有效期明细记录"
        Case 3
            sqlRechargeLog = sqlRechargeLog & " where Income_Payout=1 order by LogID desc"
            Response.Write "所有收入记录"
        Case 4
            sqlRechargeLog = sqlRechargeLog & " where Income_Payout=2 order by LogID desc"
            Response.Write "所有支出记录"
        Case 10
            If Keyword = "" Then
                sqlRechargeLog = sqlRechargeLog & " order by LogID desc"
                Response.Write "所有有效期明细记录"
            Else
                Select Case strField
                Case "UserName"
                    sqlRechargeLog = sqlRechargeLog & " where UserName like '%" & Keyword & "%' order by LogID desc"
                    Response.Write "用户名中含有“ <font color=red>" & Keyword & "</font> ”的有效期明细记录"
                Case "LogTime"
                    If IsDate(Keyword) Then
                        If SystemDatabaseType = "SQL" Then
                            sqlRechargeLog = sqlRechargeLog & " where LogTime='" & Keyword & "'  order by LogID desc"
                        Else
                            sqlRechargeLog = sqlRechargeLog & " where LogTime=#" & Keyword & "#  order by LogID desc"
                        End If
                        Response.Write "消费时间为 <font color=red>" & Keyword & "</font> 的有效期明细记录"
                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>查询的消费时间必须为日期格式！</li>"
                    End If
                End Select
            End If
        Case Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>错误的参数！</li>"
    End Select
    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>时间</td>"
    Response.Write "    <td width='80'>用户名</td>"
    Response.Write "    <td width='120'>IP地址</td>"
    Response.Write "    <td width='50'>增加有效期</td>"
    Response.Write "    <td width='50'>减少有效期</td>"
    Response.Write "    <td width='40'>摘要</td>"
    Response.Write "    <td width='60'>操作员</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "  </tr>"
    
    Set rsRechargeLog = Server.CreateObject("Adodb.RecordSet")
    rsRechargeLog.Open sqlRechargeLog, Conn, 1, 1
    If rsRechargeLog.BOF And rsRechargeLog.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何符合条件的消费记录！</td></tr>"
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
            Response.Write "    <td width='80' align='center'><a href='Admin_User.asp?Action=Show&InfoType=3&UserName=" & rsRechargeLog("UserName") & "'>" & rsRechargeLog("UserName") & "</a></td>"
            Response.Write "    <td width='120' align='center'>" & rsRechargeLog("IP") & "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsRechargeLog("Income_Payout") = 1 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsRechargeLog("Income_Payout") = 2 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='40' align='center'>"
            Select Case rsRechargeLog("Income_Payout")
            Case 1
                Response.Write "<font color='blue'>增加</font>"
            Case 2
                Response.Write "<font color='green'>减少</font>"
            Case Else
                Response.Write "其他"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>" & rsRechargeLog("Inputer") & "</td>"
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
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条有效期明细记录", True)

    Response.Write "<form name='myform' method='post' action='Admin_RechargeLog.asp' onsubmit=""return confirm('确实要删除有关记录吗？一旦删除这些记录，会出现会员查看原来已经付过费的收费信息时重复收费等问题。请慎重！')"">"
    Response.Write "如果点券明细记录太多，影响了系统性能，可以删除一定时间段前的记录以加快速度。但可能会带来会员在查看以前收过费的信息时重复收费（这样会引发众多消费纠纷问题），无法通过点券明细记录来真实分析会员的消费习惯等问题。<br>"
    Response.Write "时间范围：<input type='radio' name='DatepartType' value='0'>10天前&nbsp;&nbsp;<input type='radio' name='DatepartType' value='1'>1个月前&nbsp;&nbsp;<input type='radio' name='DatepartType' value='2'>2个月前&nbsp;&nbsp;<input type='radio' name='DatepartType' value='3'>3个月前&nbsp;&nbsp;<input type='radio' name='DatepartType' value='4'>6个月前&nbsp;&nbsp;<input type='radio' name='DatepartType' value='5' checked>1年前&nbsp;&nbsp;<input type='submit' name='submit1' value='删除记录'>"
    Response.Write "<input type='hidden' name='Action' value='Del'></form>"
End Sub

Sub Del()
    Dim DatepartType, TempDate, strDatepart
    DatepartType = PE_CLng(Trim(Request("DatepartType")))
    Select Case DatepartType
    Case 0
        TempDate = DateAdd("D", -10, Date)
        strDatepart = "10天前"
    Case 1
        TempDate = DateAdd("M", -1, Date)
        strDatepart = "1个月前"
    Case 2
        TempDate = DateAdd("M", -2, Date)
        strDatepart = "2个月前"
    Case 3
        TempDate = DateAdd("M", -3, Date)
        strDatepart = "3个月前"
    Case 4
        TempDate = DateAdd("M", -6, Date)
        strDatepart = "6个月前"
    Case 5
        TempDate = DateAdd("yyyy", -1, Date)
        strDatepart = "1年前"
    End Select
    If SystemDatabaseType = "SQL" Then
        Conn.Execute ("delete from PE_RechargeLog where LogTime<'" & TempDate & "'")
    Else
        Conn.Execute ("delete from PE_RechargeLog where LogTime<#" & TempDate & "#")
    End If
    Call WriteSuccessMsg("成功删除了 " & strDatepart & " 的记录！", "Admin_RechargeLog.asp")
End Sub
%>
