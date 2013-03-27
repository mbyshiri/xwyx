<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim LogType
LogType = PE_CLng(Trim(Request("LogType")))

FileName = "Admin_Log.asp?LogType=" & LogType
strFileName = FileName & "&Field=" & strField & "&keyword=" & Keyword

Response.Write "<html><head><title>网站日志管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("网 站 日 志 管 理", 10028)
Response.Write "<tr class='tdbg'><td width='70'>管理导航：</td><td>"
Response.Write " <a href='Admin_Log.asp'>全部日志</a> | "
Response.Write " <a href='Admin_Log.asp?LogType=1'>重要操作</a> | "
Response.Write " <a href='Admin_Log.asp?LogType=2'>系统操作</a> | "
Response.Write " <a href='Admin_Log.asp?LogType=3'>频道操作</a> | "
Response.Write " <a href='Admin_Log.asp?LogType=4'>登录失败</a> | "
Response.Write " <a href='Admin_Log.asp?LogType=5'>管理错误</a> | "
Response.Write " <a href='Admin_Log.asp?LogType=6'>越权操作</a> | "
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Show"
    Call Show
Case "DelLog"
    Call DelLog
Case "ClearLog"
    Call ClearLog
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim strAction, sqlstr
    Select Case LogType
    Case 0
        strAction = "全部日志"
    Case 1
        strAction = "重要操作日志：管理员成功登录、管理员管理、网站配置、数据库管理"
    Case 2
        strAction = "有关系统的其它操作日志"
    Case 3
        strAction = "有关频道的操作日志"
    Case 4
        strAction = "登录失败记录"
    Case 5
        strAction = "后台管理错误记录"
    Case 6
        strAction = "越权操作的记录"
    Case Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
        Exit Sub
    End Select

    Call ShowJS_Main("日志")
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：网站日志管理&nbsp;&gt;&gt;&nbsp;"
    Response.Write strAction
    Response.Write "</td></tr></table>"
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Log.asp'>"
    Response.Write "      <td>"
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' align='center' height='22'>"
    Response.Write "          <td width='30'><strong>选中</strong></td>"
    Response.Write "          <td width='300'><strong>访问地址</strong></td>"
    Response.Write "          <td><strong>操作信息</strong></td>"
    Response.Write "          <td width='120'><strong>操作时间</strong></td>"
    Response.Write "          <td width='90'><strong>IP地址</strong></td>"
    Response.Write "          <td width='60'><strong>操作人</strong></td>"
    Response.Write "          <td width='40'><strong>详细</strong></td>"
    Response.Write "        </tr>"
    
    Dim rsLog, sqlLog
    sqlLog = "select * from PE_Log where 1=1 "
    If LogType > 0 Then
        sqlLog = sqlLog & " and LogType=" & LogType
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "UserName"
            sqlLog = sqlLog & " and UserName like '%" & Keyword & "%' "
        Case "LogContent"
            sqlLog = sqlLog & " and LogContent like '%" & Keyword & "%' "
        Case "UserIP"
            sqlLog = sqlLog & " and UserIP like '%" & Keyword & "%' "
        Case Else
            sqlLog = sqlLog & " and UserName like '%" & Keyword & "%' "
        End Select
    End If
    sqlLog = sqlLog & " order by LogTime desc"
    Set rsLog = Server.CreateObject("adodb.recordset")
    rsLog.Open sqlLog, Conn, 1, 1
    totalPut = rsLog.RecordCount
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
            rsLog.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Dim LogNum
    LogNum = 0
    Do While Not rsLog.EOF
        Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "        <td width='30' align='center'><input name='LogID' type='checkbox' onclick=""unselectall()"" value='" & rsLog("LogID") & "'></td>"
        Response.Write "        <td width='300'>" & PE_HtmlEncode(rsLog("ScriptName")) & "</td>"
        Response.Write "        <td>" & rsLog("LogContent") & "</td>"
        Response.Write "        <td width='120' align='center'>" & rsLog("LogTime") & "</td>"
        Response.Write "        <td width='90' align='center'>" & rsLog("UserIP") & "</td>"
        Response.Write "        <td width='60' align='center'>" & rsLog("UserName") & "</td>"
        Response.Write "        <td width='40' align='center'><a href='Admin_Log.asp?Action=Show&LogID=" & rsLog("LogID") & "'>查看</a></td>"
        Response.Write "      </tr>"
        LogNum = LogNum + 1
        If LogNum >= MaxPerPage Then Exit Do
        rsLog.MoveNext
    Loop
    rsLog.Close
    Set rsLog = Nothing
    
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> 选中本页显示的所有日志记录</td>"
    Response.Write "    <td><input type='hidden' name='LogType' value='" & LogType & "'><input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "    <input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='DelLog'"" value='删除选中的日志记录'>&nbsp;"
    Response.Write "    <input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='ClearLog'"" value='清空日志记录'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条记录", True)

    Response.Write "<br>"
    Response.Write "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>日志搜索：</strong></td>"
    Response.Write "   <td>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='UserName' selected>操 作 人</option>"
    Response.Write "<option value='LogContent'>操作信息</option>"
    Response.Write "<option value='UserIP'>IP 地 址</option>"
    Response.Write "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='搜索'>"
    Response.Write "</td></tr></table></form>"
End Sub

Sub Show()
    Dim LogID, rsLog
    LogID = PE_CLng(Trim(Request("LogID")))
    If LogID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定相关日志！</li>"
    End If
    Set rsLog = Conn.Execute("select * from PE_Log where LogID=" & LogID)
    If rsLog.BOF And rsLog.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的日志！</li>"
        Set rsLog = Nothing
        Exit Sub
    End If
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：网站日志管理&nbsp;&gt;&gt;&nbsp;显示日志详细信息</td></tr></table>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='2' align='center'><strong>详 细 信 息</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='20%' align='center' class='tdbg5'><strong>操 作 人：</strong></td>" & vbCrLf
    If rsLog("UserName") = "" Then
        Response.Write "    <td>未知</td>" & vbCrLf
    Else
        Response.Write "    <td>" & rsLog("UserName") & "</td>" & vbCrLf
    End If
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='20%' align='center' class='tdbg5'><strong>操作时间：</strong></td>" & vbCrLf
    Response.Write "    <td>" & rsLog("LogTime") & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='20%' align='center' class='tdbg5'><strong>IP 地 址：</strong></td>" & vbCrLf
    Response.Write "    <td>" & rsLog("UserIP") & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='20%' align='center' class='tdbg5'><strong>操作信息：</strong></td>" & vbCrLf
    Response.Write "    <td>" & rsLog("LogContent") & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='20%' align='center' class='tdbg5'><strong>访问地址：</strong></td>" & vbCrLf
    Response.Write "    <td>" & PE_HtmlEncode(rsLog("ScriptName")) & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='20%' align='center' class='tdbg5'><strong>提交参数：</strong></td>" & vbCrLf
    Response.Write "    <td style='word-break:break-all;Width:fixed'>" & rsLog("PostString") & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    rsLog.Close
    Set rsLog = Nothing
End Sub

Sub DelLog()
    Dim LogID
    LogID = Request("LogID")
    If IsValidID(LogID) = False Then
        LogID = ""
    End If
    If LogID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定相关日志！</li>"
        Exit Sub
    End If
    Conn.Execute ("delete from PE_Log where Datediff(" & PE_DatePart_D & ",LogTime, " & PE_Now & ") > 2 and LogID in (" & LogID & ")")
    Call WriteSuccessMsg("成功删除了指定的日志。注意：两天内的日志会被系统保留。", ComeUrl)
End Sub

Sub ClearLog()
    If LogType = 0 Then
        Conn.Execute ("delete from PE_Log Where Datediff(" & PE_DatePart_D & ",LogTime, " & PE_Now & ") > 2")
    Else
        Conn.Execute ("delete from PE_Log where  Datediff(" & PE_DatePart_D & ",LogTime, " & PE_Now & ") > 2 and LogType=" & LogType & "")
    End If
    Call WriteEntry(1, AdminName, "清空日志")
    Call WriteSuccessMsg("成功清空了两天前的日志。注意：两天内的日志会被系统保留。", ComeUrl)
End Sub
%>
