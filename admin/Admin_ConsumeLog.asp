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
Const PurviewLevel_Others = "ConsumeLog"   '其他权限
Dim BeginDate,EndDate
BeginDate = Trim(Request("BeginDate"))
EndDate = Trim(Request("EndDate"))

strFileName = "Admin_ConsumeLog.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword &"&BeginDate="&BeginDate&"&EndDate="&EndDate

Response.Write "<html><head><title>" & PointName & "明细查询</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0'  marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
Call ShowPageTitle(PointName & "明细查询", 10044)
Response.Write "    <tr class='tdbg' height='30'> "
Response.Write "  <form name='form1' action='Admin_ConsumeLog.asp' method='get'>"
Response.Write "      <td width=400>快速查找："
Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
Response.Write "          <option value='0'"
If SearchType = 0 Then Response.Write " selected"
Response.Write ">所有" & PointName & "明细记录</option>"
Response.Write "          <option value='1'"
If SearchType = 1 Then Response.Write " selected"
Response.Write ">最近10天内的新" & PointName & "明细记录</option>"
Response.Write "          <option value='2'"
If SearchType = 2 Then Response.Write " selected"
Response.Write ">最近一月内的新" & PointName & "明细记录</option>"
Response.Write "          <option value='3'"
If SearchType = 3 Then Response.Write " selected"
Response.Write ">所有收入记录</option>"
Response.Write "          <option value='4'"
If SearchType = 4 Then Response.Write " selected"
Response.Write ">所有支出记录</option>"
Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_ConsumeLog.asp'>" & PointName & "明细首页</a></td>"& vbCrLf
Response.Write "  </form>"& vbCrLf
Response.Write " <script language='javascript'>"& vbCrLf
Response.Write " function ChangeSearch(type)"& vbCrLf
Response.Write " {"& vbCrLf
Response.Write "  if (type=='LogTime')"& vbCrLf
Response.Write "     {"& vbCrLf
Response.Write "      document.getElementById(""UserNameTable"").style.display=""none""; "& vbCrLf
Response.Write "      document.getElementById(""LogTimeTable"").style.display=""""; "& vbCrLf
Response.Write "      }"& vbCrLf
Response.Write "   else"& vbCrLf
Response.Write "     {"& vbCrLf
Response.Write "     document.getElementById(""UserNameTable"").style.display=""""; "& vbCrLf
Response.Write "     document.getElementById(""LogTimeTable"").style.display=""none""; "& vbCrLf
Response.Write "     }"& vbCrLf
Response.Write " }"& vbCrLf
Response.Write " </script>"& vbCrLf
Response.Write "<form name='form2' method='post' action='Admin_ConsumeLog.asp'>"
Response.Write "    <td>"
Response.Write "<table><tr><td>高级查询："
Response.Write "      <select name='Field' onchange=ChangeSearch(this.options[this.selectedIndex].value) id='Field'>"
Response.Write "      <option value='UserName'>用户名</option>"
Response.Write "      <option value='LogTime'>消费时间</option>"
Response.Write "      </select>"
Response.Write "      <td><Table id=UserNameTable style=""DISPLAY""><tr><td><input name='Keyword' style='display' type='text' id='Keyword' size='20' maxlength='30'></td></tr></Table></td>"
Response.Write "      <td><table id=LogTimeTable style=""DISPLAY: none""><tr><td>起始日期<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.form2.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;结束日期<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.form2.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr></Table></td>"
Response.Write "      <td>"	
Response.Write "      <input type='submit' name='Submit2' value=' 查 询 '>"
Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'></td></tr></table>"
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
    Dim rsConsumeLog, sqlConsumeLog
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    sqlConsumeLog = "select * from PE_ConsumeLog "
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>您现在的位置：<a href='Admin_ConsumeLog.asp'>" & PointName & "明细记录管理</a>&nbsp;&gt;&gt;&nbsp;"
    Select Case SearchType
        Case 0
            sqlConsumeLog = sqlConsumeLog & " order by LogID desc"
            Response.Write "所有" & PointName & "明细记录"
        Case 1
            sqlConsumeLog = sqlConsumeLog & " where datediff(" & PE_DatePart_D & ",LogTime," & PE_Now & ")<10 order by LogID desc"
            Response.Write "最近10天内的新" & PointName & "明细记录"
        Case 2
            sqlConsumeLog = sqlConsumeLog & " where datediff(" & PE_DatePart_M & ",LogTime," & PE_Now & ")<1 order by LogID desc"
            Response.Write "最近一月内的新" & PointName & "明细记录"
        Case 3
            sqlConsumeLog = sqlConsumeLog & " where Income_Payout=1 order by LogID desc"
            Response.Write "所有收入记录"
        Case 4
            sqlConsumeLog = sqlConsumeLog & " where Income_Payout=2 order by LogID desc"
            Response.Write "所有支出记录"
        Case 5
            sqlConsumeLog = sqlConsumeLog & " where Income_Payout<=2 order by LogID desc"
            Response.Write "所有非开户记录"
        Case 10
            If Keyword = "" and BeginDate = "" and EndDate = "" Then
                sqlConsumeLog = sqlConsumeLog & " order by LogID desc"
                Response.Write "所有" & PointName & "明细记录"
            Else

                Select Case strField
                Case "UserName"
                    sqlConsumeLog = sqlConsumeLog & " where UserName like '%" & Keyword & "%' order by LogID desc"
                    Response.Write "用户名中含有“ <font color=red>" & Keyword & "</font> ”的" & PointName & "明细记录"
                Case "LogTime"
                    sqlConsumeLog = sqlConsumeLog & " where 1=1"
                    If (IsDate(BeginDate) and EndDate="") Or (IsDate(BeginDate) and IsDate(EndDate)) Or (IsDate(EndDate) and BeginDate="") Then
                        If SystemDatabaseType = "SQL"  Then 
                            If BeginDate<>"" Then
                                 sqlConsumeLog = sqlConsumeLog & " and LogTime>='" & BeginDate &"'"
                            End If
                            If EndDate<>"" Then
                                sqlConsumeLog = sqlConsumeLog & " and LogTime<='" & EndDate &"'"
                            End If                      
                        Else
                            If BeginDate<>"" Then
                                sqlConsumeLog = sqlConsumeLog & " and LogTime>=#" & BeginDate &"#"
                            End If
								
                            If EndDate<>"" Then
                                sqlConsumeLog = sqlConsumeLog & " and LogTime<=#" & EndDate &"#"
                            End If 
                        End If
                            sqlConsumeLog = sqlConsumeLog &"  order by LogID desc"
                            If (IsDate(BeginDate) and EndDate="") Then Response.Write "消费时间为 <font color=red>" & BeginDate & "之后</font> 的" & PointName & "明细记录"
                            If (IsDate(BeginDate) and IsDate(EndDate)) Then Response.Write "消费时间为 <font color=red>" & BeginDate & "</font> 与 <font color=red>"& EndDate &"</font> 之间的" & PointName & "明细记录"
                            If (IsDate(EndDate) and BeginDate="") Then Response.Write "消费时间为 <font color=red>" & EndDate & "之前</font> 的" & PointName & "明细记录"
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
    Call PopCalendarInit
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>消费时间</td>"
    Response.Write "    <td width='80'>用户名</td>"
    Response.Write "    <td width='100'>IP地址</td>"
    Response.Write "    <td width='50'>收入" & PointName & "数</td>"
    Response.Write "    <td width='50'>支出" & PointName & "数</td>"
    Response.Write "    <td width='40'>摘要</td>"
    Response.Write "    <td width='60'>重复次数</td>"
    Response.Write "    <td width='60'>操作员</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "  </tr>"
    
    Set rsConsumeLog = Server.CreateObject("Adodb.RecordSet")
    rsConsumeLog.Open sqlConsumeLog, Conn, 1, 1
    If rsConsumeLog.BOF And rsConsumeLog.EOF Then
        TotalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何符合条件的消费记录！</td></tr>"
    Else
        TotalPut = rsConsumeLog.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
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
            Response.Write "    <td width='80' align='center'><a href='Admin_User.asp?Action=Show&InfoType=2&UserName=" & rsConsumeLog("UserName") & "'>" & rsConsumeLog("UserName") & "</a></td>"
            Response.Write "    <td width='100' align='center'>" & rsConsumeLog("IP") & "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsConsumeLog("Income_Payout") = 1 Then Response.Write rsConsumeLog("Point")
            Response.Write "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsConsumeLog("Income_Payout") = 2 Then Response.Write rsConsumeLog("Point")
            Response.Write "</td>"
            Response.Write "    <td width='40' align='center'>"
            Select Case rsConsumeLog("Income_Payout")
            Case 1
                Response.Write "<font color='blue'>收入</font>"
            Case 2
                Response.Write "<font color='green'>支出</font>"
            Case Else
                Response.Write "其他"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>" & rsConsumeLog("Times") & "</td>"
            Response.Write "    <td width='60' align='center'>" & rsConsumeLog("Inputer") & "</td>"
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
    Response.Write "    <td colspan='3' align='right'>本页合计：</td>"
    Response.Write "    <td align='right'>" & TotalIncome & "</td>"
    Response.Write "    <td align='right'>" & TotalPayout & "</td>"
    Response.Write "    <td colspan='4'>&nbsp;</td>"
    Response.Write "  </tr>"

    Dim trs, TotalIncomeAll, TotalPayoutAll
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where Income_Payout=1")
    If IsNull(trs(0)) Then
        TotalIncomeAll = 0
    Else
        TotalIncomeAll = trs(0)
    End If
    Set trs = Nothing
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where Income_Payout=2")
    If IsNull(trs(0)) Then
        TotalPayoutAll = 0
    Else
        TotalPayoutAll = trs(0)
    End If
    Set trs = Nothing
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='3' align='right'>总计金额：</td>"
    Response.Write "    <td align='right'>" & TotalIncomeAll & "</td>"
    Response.Write "    <td align='right'>" & TotalPayoutAll & "</td>"
    Response.Write "    <td colspan='4' align='center'>" & PointName & "数余额：" & TotalIncomeAll - TotalPayoutAll & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "条" & PointName & "明细记录", True)
    Response.Write "<form name='myform' method='post' action='Admin_ConsumeLog.asp' onsubmit=""return confirm('确实要删除有关记录吗？一旦删除这些记录，会出现会员查看原来已经付过费的收费信息时重复收费等问题。请慎重！')"">"
    Response.Write "如果" & PointName & "明细记录太多，影响了系统性能，可以删除一定时间段前的记录以加快速度。但可能会带来会员在查看以前收过费的信息时重复收费（这样会引发众多消费纠纷问题），无法通过" & PointName & "明细记录来真实分析会员的消费习惯等问题。<br>"
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
        Conn.Execute ("delete from PE_ConsumeLog where LogTime<'" & TempDate & "'")
    Else
        Conn.Execute ("delete from PE_ConsumeLog where LogTime<#" & TempDate & "#")
    End If
    Call WriteSuccessMsg("成功删除了 " & strDatepart & " 的记录！", "Admin_ConsumeLog.asp")
End Sub
%>
