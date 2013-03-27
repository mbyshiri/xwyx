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

Dim BeginDate, EndDate, i
BeginDate = Trim(Request("BeginDate"))
EndDate = Trim(Request("EndDate"))
If IsDate(BeginDate) Then
    BeginDate = CDate(BeginDate)
Else
    BeginDate = CDate(Year(Date) & "-1-1")
End If
If IsDate(EndDate) Then
    EndDate = CDate(EndDate)
Else
    EndDate = Date
End If

Dim iYear, iMonth, iDate1, iDate2, iCount, jcount, j
iYear = Year(BeginDate)
iMonth = Month(BeginDate)
iCount = DateDiff("m", BeginDate, EndDate) + 1
Action = Trim(Request("Action"))
Dim tUserName, TableName
tUserName = ReplaceBadChar(Trim(Request("UserName")))

Dim arrCount()
Dim iTemp
GroupID = PE_Clng(Trim(Request("GroupID")))

Response.Write "<html><head><title>网站统计</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'> " & vbCrLf
Response.Write "    <td height='22' colspan='10' align='center'><b>网 站 统 计</b></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td height='30'>"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByChannelMonth'>按频道/月份统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByChannelUser'>按频道/录入者统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByChannelEditor'>按频道/审核人统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByClassMonth'>按栏目/月份统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByClassUser'>按栏目/录入者统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByClassEditor'>按栏目/审核人统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByUserMonth'>按录入者/月份统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByEditorMonth'>按审核人/月份统计</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_SiteCount.asp?Action=CountByChannelUserGroup'>按频道/会员组统计</a> | "
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "<script language='JavaScript' src='PopCalendar.js'></script>" & vbCrLf
Response.Write "<script language='JavaScript'>" & vbCrLf
Response.Write "PopCalendar = getCalendarInstance()" & vbCrLf
Response.Write "PopCalendar.startAt = 0 // 0 - sunday ; 1 - monday" & vbCrLf
Response.Write "PopCalendar.showWeekNumber = 0 // 0 - don't show; 1 - show" & vbCrLf
Response.Write "PopCalendar.showTime = 0 // 0 - don't show; 1 - show" & vbCrLf
Response.Write "PopCalendar.showToday = 0 // 0 - don't show; 1 - show" & vbCrLf
Response.Write "PopCalendar.showWeekend = 1 // 0 - don't show; 1 - show" & vbCrLf
Response.Write "PopCalendar.showHolidays = 1 // 0 - don't show; 1 - show" & vbCrLf
Response.Write "PopCalendar.showSpecialDay = 1 // 0 - don't show, 1 - show" & vbCrLf
Response.Write "PopCalendar.selectWeekend = 0 // 0 - don't Select; 1 - Select" & vbCrLf
Response.Write "PopCalendar.selectHoliday = 0 // 0 - don't Select; 1 - Select" & vbCrLf
Response.Write "PopCalendar.addCarnival = 0 // 0 - don't Add; 1- Add to Holiday" & vbCrLf
Response.Write "PopCalendar.addGoodFriday = 0 // 0 - don't Add; 1- Add to Holiday" & vbCrLf
Response.Write "PopCalendar.language = 0 // 0 - Chinese; 1 - English" & vbCrLf
Response.Write "PopCalendar.defaultFormat = 'yyyy-mm-dd' //Default Format dd-mm-yyyy" & vbCrLf
Response.Write "PopCalendar.fixedX = -1 // x position (-1 if to appear below control)" & vbCrLf
Response.Write "PopCalendar.fixedY = -1 // y position (-1 if to appear below control)" & vbCrLf
Response.Write "PopCalendar.fade = .5 // 0 - don't fade; .1 to 1 - fade (Only IE) " & vbCrLf
Response.Write "PopCalendar.shadow = 1 // 0  - don't shadow, 1 - shadow" & vbCrLf
Response.Write "PopCalendar.move = 1 // 0  - don't move, 1 - move (Only IE)" & vbCrLf
Response.Write "PopCalendar.saveMovePos = 1  // 0  - don't save, 1 - save" & vbCrLf
Response.Write "PopCalendar.centuryLimit = 40 // 1940 - 2039" & vbCrLf
Response.Write "PopCalendar.initCalendar()" & vbCrLf
Response.Write "</script>" & vbCrLf

Response.Write "<form method='post' name='form1' action='Admin_SiteCount.asp'>"

Select Case Action
Case "CountByChannelMonth"
    Call CountByChannelMonth
Case "CountByChannelUser"
    Call CountByChannelUser(1)
Case "CountByChannelEditor"
    Call CountByChannelUser(2)
Case "CountByClassMonth"
    Call CountByClassMonth
Case "CountByClassUser"
    Call CountByClassUser(1)
Case "CountByClassEditor"
    Call CountByClassUser(2)
Case "CountByUserMonth"
    Call CountByUserMonth(1)
Case "CountByEditorMonth"
    Call CountByUserMonth(2)
Case "CountByChannelUserGroup"
    Call CountByChannelUserGroup
Case Else
    Call CountByChannelMonth
End Select
Response.Write "</form>"
Response.Write "</body></html>"

Sub ShowDateField()
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>统计日期范围：</td><td>起始日期<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10' value='" & BeginDate & "'><a style='cursor:hand;' onClick='PopCalendar.show(document.form1.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;结束日期<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10' value='" & EndDate & "'><a style='cursor:hand;' onClick='PopCalendar.show(document.form1.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
    Response.Write "<tr class='tdbg'><td colspan='2' align='center'><input type='hidden' name='Action' value='" & Action & "'><input type='submit' name='submit' value='开始统计'></td></tr>"
End Sub

Sub CountByChannelMonth()
    Response.Write "<p align='center'>按 频 道 / 月 份 统 计</p>"
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Call ShowUserField
    Call ShowDateField
    Response.Write "</table><br>"

    If Trim(Request("BeginDate")) = "" Then
        Exit Sub
    End If

    ReDim arrCount(iCount + 2)

    Response.Write "<table border='0' align='center' cellpadding='2' cellspacing='1' class='border' width='" & iCount * 60 + 160 & "'><tr align='center' class='title'><td width='100'>频道名</td>"
    For i = 1 To iCount
        Response.Write "<td width='60'>" & iYear & "-" & iMonth & "</td>"
        iMonth = iMonth + 1
        If iMonth > 12 Then
            iYear = iYear + 1
            iMonth = 1
        End If
    Next
    Response.Write "<td width='60'>合计</td></tr>"


    Dim rsChannel
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelType<=1 and Disabled=0 and ModuleType<4 order by OrderID")
    Do While Not rsChannel.EOF
        Select Case rsChannel("ModuleType")
        Case 1
            TableName = "PE_Article"
        Case 2
            TableName = "PE_Soft"
        Case 3
            TableName = "PE_Photo"
        End Select
        Response.Write "<tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'"">"
        Response.Write "<td><a href='Admin_SiteCount.asp?Action=CountByClassMonth&ChannelID=" & rsChannel("ChannelID") & "&UserName=" & tUserName & "&BeginDate=" & BeginDate & "&EndDate=" & EndDate & "'>" & rsChannel("ChannelName") & "</a></td>"
        iYear = Year(BeginDate)
        iMonth = Month(BeginDate)

        For i = 1 To iCount
            iDate1 = iYear & "-" & iMonth & "-1"
            iDate2 = DateAdd("m", 1, CDate(iDate1))
            If iDate2 > DateAdd("d", 1, CDate(EndDate)) Then iDate2 = DateAdd("d", 1, CDate(EndDate))				
            If SystemDatabaseType = "SQL" Then
                iDate1 = "'" & iDate1 & "'"
                iDate2 = "'" & iDate2 & "'"
            Else
                iDate1 = "#" & iDate1 & "#"
                iDate2 = "#" & iDate2 & "#"
            End If
            If tUserName = "" Then
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Else
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Inputer='" & tUserName & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            End If
            arrCount(i) = arrCount(i) + iTemp
            Response.Write "<td>" & iTemp & "</td>"
            iMonth = iMonth + 1
            If iMonth > 12 Then
                iYear = iYear + 1
                iMonth = 1
            End If
        Next
        If SystemDatabaseType = "SQL" Then
            iDate1 = "'" & BeginDate & "'"
            iDate2 = "'" & DateAdd("d", 1, CDate(EndDate)) & "'"
        Else
            iDate1 = "#" & BeginDate & "#"
            iDate2 = "#" & DateAdd("d", 1, CDate(EndDate)) & "#"			
        End If

        If tUserName = "" Then
            iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
        Else
            iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Inputer='" & tUserName & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
        End If
        arrCount(i) = arrCount(i) + iTemp
        Response.Write "<td>" & iTemp & "</td>"
        Response.Write "</tr>"
        rsChannel.movenext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "<tr align='center' class='tdbg2'><td>合计</td>"
    For i = 1 To iCount + 1
        Response.Write "<td>" & arrCount(i) & "</td>"
    Next
    Response.Write "</tr>"
    Response.Write "</table>"
End Sub

Sub CountByChannelUser(UserType)
    If UserType = 1 Then
        Response.Write "<p align='center'>按 频 道 / 录 入 者 统 计</p>"
    Else
        Response.Write "<p align='center'>按 频 道 / 审　核　人 统 计</p>"
    End If
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Call ShowDateField
    Response.Write "</table><br>"

    If Trim(Request("BeginDate")) = "" Then
        Exit Sub
    End If
    Dim rsUser
    If UserType = 1 Then
        Set rsUser = Conn.Execute("select distinct(Inputer) from PE_Article")
        If SystemDatabaseType = "SQL" Then
            iCount = Conn.Execute("select Count(distinct(Inputer)) from PE_Article")(0)
        Else
            iCount = Conn.Execute("select Count(0) from (select distinct(Inputer) from PE_Article)")(0)
        End If
    Else
        Set rsUser = Conn.Execute("select distinct(Editor) from PE_Article")
        If SystemDatabaseType = "SQL" Then
            iCount = Conn.Execute("select Count(distinct(Editor)) from PE_Article")(0)
        Else
            iCount = Conn.Execute("select Count(0) from (select distinct(Editor) from PE_Article)")(0)
        End If
    End If
    ReDim arrCount(iCount)
    If rsUser.Bof And rsUser.EOF Then
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

    Response.Write "<table border='0' align='center' cellpadding='2' cellspacing='1' class='border' width='" & iCount * 60 + 160 & "'><tr align='center' class='title'><td width='100'>频道名</td>"
    Do While Not rsUser.EOF
        Response.Write "<td width='60'>" & rsUser(0) & "</td>"
        rsUser.movenext
    Loop
    Response.Write "<td width='60'>合计</td></tr>"
    If SystemDatabaseType = "SQL" Then
        iDate1 = "'" & BeginDate & "'"
        iDate2 = "'" & DateAdd("d", 1, CDate(EndDate)) & "'"
    Else
        iDate1 = "#" & BeginDate & "#"
        iDate2 = "#" & DateAdd("d", 1, CDate(EndDate)) & "#"
    End If

    Dim rsChannel
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelType<=1 and Disabled=0 and ModuleType<4 order by OrderID")
    Do While Not rsChannel.EOF
        Select Case rsChannel("ModuleType")
        Case 1
            TableName = "PE_Article"
        Case 2
            TableName = "PE_Soft"
        Case 3
            TableName = "PE_Photo"
        End Select
        Response.Write "<tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'""><td>" & rsChannel("ChannelName") & "</td>"

        i = 0
        rsUser.MoveFirst
        Do While Not rsUser.EOF
            If UserType = 1 Then
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=0 and Status=3 and Inputer='" & rsUser(0) & "' and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Else
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=0 and Status=3 and Editor='" & rsUser(0) & "' and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            End If
            If iTemp = 0 Then
                Response.Write "<td>0</td>"
            Else
                Response.Write "<td><a href='Admin_SiteCount.asp?Action=CountByChannelMonth&ChannelID=" & rsChannel("ChannelID") & "&UserName=" & rsUser(0) & "&BeginDate=" & BeginDate & "&EndDate=" & EndDate & "'>" & iTemp & "</a></td>"
            End If
            arrCount(i) = arrCount(i) + iTemp
            i = i + 1
            rsUser.movenext
        Loop
        iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
        arrCount(i) = arrCount(i) + iTemp
        Response.Write "<td>" & iTemp & "</td>"
        Response.Write "</tr>"
        rsChannel.movenext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    rsUser.Close
    Set rsUser = Nothing

    Response.Write "<tr align='center' class='tdbg2'><td>合计</td>"
    For i = 0 To iCount
        Response.Write "<td>" & arrCount(i) & "</td>"
    Next
    Response.Write "</tr>"
    Response.Write "</table>"
End Sub


Sub CountByClassMonth()
    Response.Write "<p align='center'>按 栏 目 / 月 份 统 计</p>"
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Call ShowChannelField
    Call ShowUserField
    Call ShowDateField
    Response.Write "</table><br>"

    If ChannelID = 0 Or Trim(Request("BeginDate")) = "" Then
        Exit Sub
    End If
    ReDim arrCount(iCount + 2)

    Response.Write "<table border='0' align='center' cellpadding='2' cellspacing='1' class='border' width='" & iCount * 60 + 360 & "'><tr align='center' class='title'><td width='300'>栏目名称</td>"
    For i = 1 To iCount
        Response.Write "<td width='60'>" & iYear & "-" & iMonth & "</td>"
        iMonth = iMonth + 1
        If iMonth > 12 Then
            iYear = iYear + 1
            iMonth = 1
        End If
    Next
    Response.Write "<td width='60'>合计</td></tr>"

    Dim arrShowLine(20), i
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    Dim sqlClass, rsClass, iDepth, ClassDir, ClassItemDir
    Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID")
    Do While Not rsClass.EOF
        Response.Write "<tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'""><td align='left'>"
        iDepth = rsClass("Depth")
        If rsClass("NextID") > 0 Then
            arrShowLine(iDepth) = True
        Else
            arrShowLine(iDepth) = False
        End If
        If iDepth > 0 Then
            For i = 1 To iDepth
                If i = iDepth Then
                    If rsClass("NextID") > 0 Then
                        Response.Write "<img src='../images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
                    Else
                        Response.Write "<img src='../images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
                    End If
                Else
                    If arrShowLine(i) = True Then
                        Response.Write "<img src='../images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
                    Else
                        Response.Write "<img src='../images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
                    End If
                End If
            Next
        End If
        If rsClass("Child") > 0 Then
            Response.Write "<img src='../images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
        Else
            Response.Write "<img src='../images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
        End If
        If rsClass("Depth") = 0 Then
            Response.Write "<b>"
        End If
        Response.Write "" & rsClass("ClassName") & ""
        If rsClass("Child") > 0 Then
            Response.Write "（" & rsClass("Child") & "）"
        End If
        If rsClass("ClassType") = 2 Then
            Response.Write " <font color=blue>（外）</font>"
        Else
            'Response.Write " [" & rsClass("ClassDir") & "]"
        End If
        Response.Write "</td>"

        If rsClass("ClassType") = 1 Then
            iYear = Year(BeginDate)
            iMonth = Month(BeginDate)
            For i = 1 To iCount
                iDate1 = iYear & "-" & iMonth & "-1"
                iDate2 = DateAdd("m", 1, CDate(iDate1))
                If iDate2 > DateAdd("d", 1, CDate(EndDate)) Then iDate2 = DateAdd("d", 1, CDate(EndDate))
                If SystemDatabaseType = "SQL" Then
                    iDate1 = "'" & iDate1 & "'"
                    iDate2 = "'" & iDate2 & "'"
                Else
                    iDate1 = "#" & iDate1 & "#"
                    iDate2 = "#" & iDate2 & "#"
                End If
                If tUserName = "" Then
                    iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                Else
                    iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Inputer='" & tUserName & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                End If
                Response.Write "<td>" & iTemp & "</td>"
                arrCount(i) = arrCount(i) + iTemp
                iMonth = iMonth + 1
                If iMonth > 12 Then
                    iYear = iYear + 1
                    iMonth = 1
                End If
            Next
            If SystemDatabaseType = "SQL" Then
                iDate1 = "'" & BeginDate & "'"
                iDate2 = "'" & DateAdd("d", 1, CDate(EndDate)) & "'"
            Else
                iDate1 = "#" & BeginDate & "#"
                iDate2 = "#" & DateAdd("d", 1, CDate(EndDate)) & "#"
            End If
            If tUserName = "" Then
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Else
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Inputer='" & tUserName & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            End If
            arrCount(i) = arrCount(i) + iTemp
            Response.Write "<td>" & iTemp & "</td>"
        Else
            For i = 1 To iCount + 1
                Response.Write "<td>0</td>"
            Next
        End If
        Response.Write "</tr>"
        rsClass.movenext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    Response.Write "<tr align='center' class='tdbg2'><td>合计</td>"
    For i = 1 To iCount + 1
        Response.Write "<td>" & arrCount(i) & "</td>"
    Next
    Response.Write "</tr>"
    Response.Write "</table>"
End Sub

Sub CountByClassUser(UserType)
    If UserType = 1 Then
        Response.Write "<p align='center'>按 栏 目 / 录 入 者 统 计</p>"
    Else
        Response.Write "<p align='center'>按 栏 目 / 审 核 人 统 计</p>"
    End If
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Call ShowChannelField
    Call ShowDateField
    Response.Write "</table><br>"

    If ChannelID = 0 Or Trim(Request("BeginDate")) = "" Then
        Exit Sub
    End If
    Dim rsUser
    Set rsUser = Conn.Execute("select AdminName from PE_Admin")
    iCount = Conn.Execute("select Count(0) from PE_Admin")(0)
    ReDim arrCount(iCount)
    

    Response.Write "<table border='0' align='center' cellpadding='2' cellspacing='1' class='border' width='" & iCount * 60 + 360 & "'><tr align='center' class='title'><td width='300'>栏目名称</td>"
    Do While Not rsUser.EOF
        Response.Write "<td width='60'>" & rsUser(0) & "</td>"
        rsUser.movenext
    Loop
    Response.Write "<td width='60'>合计</td></tr>"

    If SystemDatabaseType = "SQL" Then
        iDate1 = "'" & BeginDate & "'"
        iDate2 = "'" & DateAdd("d", 1, CDate(EndDate)) & "'"
    Else
        iDate1 = "#" & BeginDate & "#"
        iDate2 = "#" & DateAdd("d", 1, CDate(EndDate)) & "#"
    End If

    Dim arrShowLine(20), i
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    Dim sqlClass, rsClass, iDepth, ClassDir, ClassItemDir
    Set rsClass = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID")
    Do While Not rsClass.EOF
        Response.Write "<tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'""><td align='left'>"
        iDepth = rsClass("Depth")
        If rsClass("NextID") > 0 Then
            arrShowLine(iDepth) = True
        Else
            arrShowLine(iDepth) = False
        End If
        If iDepth > 0 Then
            For i = 1 To iDepth
                If i = iDepth Then
                    If rsClass("NextID") > 0 Then
                        Response.Write "<img src='../images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
                    Else
                        Response.Write "<img src='../images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
                    End If
                Else
                    If arrShowLine(i) = True Then
                        Response.Write "<img src='../images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
                    Else
                        Response.Write "<img src='../images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
                    End If
                End If
            Next
        End If
        If rsClass("Child") > 0 Then
            Response.Write "<img src='../images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
        Else
            Response.Write "<img src='../images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
        End If
        If rsClass("Depth") = 0 Then
            Response.Write "<b>"
        End If
        Response.Write "<a href='Admin_Class.asp?Action=Modify&ChannelID=" & ChannelID & "&ClassID=" & rsClass("ClassID") & "' title='" & nohtml(rsClass("Tips")) & "'>" & rsClass("ClassName") & "</a>"
        If rsClass("Child") > 0 Then
            Response.Write "（" & rsClass("Child") & "）"
        End If
        If rsClass("ClassType") = 2 Then
            Response.Write " <font color=blue>（外）</font>"
        Else
            'Response.Write " [" & rsClass("ClassDir") & "]"
        End If
        Response.Write "</td>"

        If rsClass("ClassType") = 1 Then
            i = 0
            rsUser.MoveFirst
            Do While Not rsUser.EOF
                If UserType = 1 Then
                    iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Deleted=0 and Status=3 and Inputer='" & rsUser(0) & "' and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                Else
                    iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Deleted=0 and Status=3 and Editor='" & rsUser(0) & "' and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                End If
                If iTemp = 0 Then
                    Response.Write "<td>0</td>"
                Else
                    Response.Write "<td><a href='Admin_SiteCount.asp?Action=CountByClassMonth&ChannelID=" & ChannelID & "&UserName=" & rsUser(0) & "&BeginDate=" & BeginDate & "&EndDate=" & EndDate & "'>" & iTemp & "</a></td>"
                End If
                arrCount(i) = arrCount(i) + iTemp
                i = i + 1
                rsUser.movenext
            Loop
            iTemp = Conn.Execute("select count(0) from " & TableName & " where ClassID=" & rsClass("ClassID") & " and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Response.Write "<td>" & iTemp & "</td>"
        Else
            For i = 0 To iCount
                Response.Write "<td>0</td>"
            Next
        End If
        Response.Write "</tr>"
        rsClass.movenext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    Response.Write "<tr align='center' class='tdbg2'><td>合计</td>"
    For i = 0 To iCount
        Response.Write "<td>" & arrCount(i) & "</td>"
    Next
    Response.Write "</tr>"
    Response.Write "</table>"
End Sub

Sub CountByUserMonth(UserType)
    If UserType = 1 Then
        Response.Write "<p align='center'>按 录 入 者 / 月 份 统 计</p>"
    Else
        Response.Write "<p align='center'>按 审　核　人 / 月 份 统 计</p>"
    End If
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Call ShowChannelField
    Call ShowDateField
    Response.Write "</table><br>"

    If Trim(Request("BeginDate")) = "" Then
        Exit Sub
    End If
    If ChannelID = 0 Then TableName = "PE_Article"
    Dim rsUser
    If UserType = 1 Then
        Set rsUser = Conn.Execute("select distinct(Inputer) from " & TableName & "")
    Else
        Set rsUser = Conn.Execute("select distinct(Editor) from " & TableName & "")
    End If
    ReDim arrCount(iCount + 2)
    

    Response.Write "<table border='0' align='center' cellpadding='2' cellspacing='1' class='border' width='" & iCount * 60 + 160 & "'><tr align='center' class='title'><td width='100'>用户名</td>"
    For i = 1 To iCount
        Response.Write "<td width='60'>" & iYear & "-" & iMonth & "</td>"
        iMonth = iMonth + 1
        If iMonth > 12 Then
            iYear = iYear + 1
            iMonth = 1
        End If
    Next
    Response.Write "<td width='60'>合计</td></tr>"

    Do While Not rsUser.EOF
        Response.Write "<tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'"">"
        Response.Write "<td>" & rsUser(0) & "</a></td>"
        iYear = Year(BeginDate)
        iMonth = Month(BeginDate)

        For i = 1 To iCount
            iDate1 = iYear & "-" & iMonth & "-1"
            iDate2 = DateAdd("m", 1, CDate(iDate1))
            If iDate2 > DateAdd("d", 1, CDate(EndDate)) Then iDate2 = DateAdd("d", 1, CDate(EndDate))			
            If SystemDatabaseType = "SQL" Then
                iDate1 = "'" & iDate1 & "'"
                iDate2 = "'" & iDate2 & "'"
            Else
                iDate1 = "#" & iDate1 & "#"
                iDate2 = "#" & iDate2 & "#"
            End If
            If UserType = 1 Then
                If ChannelID > 0 Then
                    iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & ChannelID & " and Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                Else
                    iTemp = Conn.Execute("select count(0) from PE_Article where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                    iTemp = iTemp + Conn.Execute("select count(0) from PE_Soft where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                    iTemp = iTemp + Conn.Execute("select count(0) from PE_Photo where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                End If
            Else
                If ChannelID > 0 Then
                    iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & ChannelID & " and Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                Else
                    iTemp = Conn.Execute("select count(0) from PE_Article where Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                    iTemp = iTemp + Conn.Execute("select count(0) from PE_Soft where Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                    iTemp = iTemp + Conn.Execute("select count(0) from PE_Photo where Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                End If
            End If
            arrCount(i) = arrCount(i) + iTemp
            Response.Write "<td>" & iTemp & "</td>"
            iMonth = iMonth + 1
            If iMonth > 12 Then
                iYear = iYear + 1
                iMonth = 1
            End If
        Next
        If SystemDatabaseType = "SQL" Then
            iDate1 = "'" & BeginDate & "'"
            iDate2 = "'" & DateAdd("d", 1, CDate(EndDate)) & "'"
        Else
            iDate1 = "#" & BeginDate & "#"
            iDate2 = "#" & DateAdd("d", 1, CDate(EndDate)) & "#"
        End If
        If UserType = 1 Then
            If ChannelID > 0 Then
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & ChannelID & " and Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Else
                iTemp = Conn.Execute("select count(0) from PE_Article where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                iTemp = iTemp + Conn.Execute("select count(0) from PE_Soft where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                iTemp = iTemp + Conn.Execute("select count(0) from PE_Photo where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            End If
        Else
            If ChannelID > 0 Then
                iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & ChannelID & " and Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Else
                iTemp = Conn.Execute("select count(0) from PE_Article where Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                iTemp = iTemp + Conn.Execute("select count(0) from PE_Soft where Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
                iTemp = iTemp + Conn.Execute("select count(0) from PE_Photo where Editor='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            End If
        End If
        arrCount(i) = arrCount(i) + iTemp
        Response.Write "<td>" & iTemp & "</td>"
        Response.Write "</tr>"
        rsUser.movenext
    Loop
    rsUser.Close
    Set rsUser = Nothing
    Response.Write "<tr align='center' class='tdbg2'><td>合计</td>"
    For i = 1 To iCount + 1
        Response.Write "<td>" & arrCount(i) & "</td>"
    Next
    Response.Write "</tr>"

    Response.Write "</table>"
End Sub

Sub CountByChannelUserGroup()
    Response.Write "<p align='center'>按 频 道 / 会 员 组 统 计</p>"
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Call ShowChannelField
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>指定会员组：</td><td><select name='GroupID'><option value='0'>请选择要统计的会员组</option>"
    Dim rsUserGroup
    Set rsUserGroup = Conn.Execute("select GroupID,GroupName from PE_UserGroup where GroupID<>-1")
    Do While Not rsUserGroup.EOF
        If rsUserGroup("GroupID") = GroupID Then
            Response.Write "<option value='" & rsUserGroup("GroupID") & "' selected>" & rsUserGroup("GroupName") & "</option>"
        Else
            Response.Write "<option value='" & rsUserGroup("GroupID") & "'>" & rsUserGroup("GroupName") & "</option>"
        End If
        rsUserGroup.movenext
    Loop
    rsUserGroup.Close
    Set rsUserGroup = Nothing
    Response.Write "</select></td></tr>"

    Call ShowDateField
    Response.Write "</table><br>"

    If ChannelID = 0 Or Trim(Request("BeginDate")) = "" Then
        Exit Sub
    End If
    
    If ChannelID = 0 Then TableName = "PE_Article"
    Dim rsUser
    Set rsUser = Conn.Execute("select a.UserName,b.TrueName from PE_User as a left join PE_contacter as b on a.ContacterID=b.ContacterID where GroupID=" & GroupID & " order by a.UserID")
    jcount = Conn.Execute("select count(0) from PE_User where GroupID=" & GroupID & "")(0)
    ReDim tt(jcount)
    ReDim arrCount(iCount + 2)
    

    Response.Write "<table border='0' align='center' cellpadding='2' cellspacing='1' class='border' width='" & iCount * 60 + 160 & "'><tr align='center' class='title'><td width='100'>用户名</td><td width='130'>真实姓名</td>"
    For i = 1 To iCount
        Response.Write "<td width='80'>" & iYear & "-" & iMonth & "</td>"
        iMonth = iMonth + 1
        If iMonth > 12 Then
            iYear = iYear + 1
            iMonth = 1
        End If
    Next
    Response.Write "<td width='60'>合计</td></tr>"
    j = 1
    Do While Not rsUser.EOF
       Response.Write "<tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'"">"
        Response.Write "<td>" & rsUser(0) & "</td><td>" & rsUser(1) & "</td>"
        iYear = Year(BeginDate)
        iMonth = Month(BeginDate)
        For i = 1 To iCount
            iDate1 = iYear & "-" & iMonth & "-1"
            iDate2 = DateAdd("m", 1, CDate(iDate1))
            If iDate2 > DateAdd("d", 1, CDate(EndDate)) Then iDate2 = DateAdd("d", 1, CDate(EndDate))			
            If SystemDatabaseType = "SQL" Then
                iDate1 = "'" & iDate1 & "'"
                iDate2 = "'" & iDate2 & "'"
            Else
                iDate1 = "#" & iDate1 & "#"
                iDate2 = "#" & iDate2 & "#"
            End If
            If ChannelID > 0 Then
               iTemp = Conn.Execute("select count(0) from " & TableName & " where ChannelID=" & ChannelID & " and Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            Else
               iTemp = Conn.Execute("select count(0) from PE_Article where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
               iTemp = iTemp + Conn.Execute("select count(0) from PE_Soft where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
               iTemp = iTemp + Conn.Execute("select count(0) from PE_Photo where Inputer='" & rsUser(0) & "' and Deleted=0 and Status=3 and UpdateTime>=" & iDate1 & " and updatetime<" & iDate2 & "")(0)
            End If
            arrCount(i) = arrCount(i) + iTemp
            Response.Write "<td>" & iTemp & "</td>"
            iMonth = iMonth + 1
            If iMonth > 12 Then
                iYear = iYear + 1
                iMonth = 1
            End If
            tt(j) = tt(j) + iTemp
        Next
        Response.Write "<td>" & tt(j) & "</td>"
        Response.Write "</tr>"
        arrCount(iCount + 1) = arrCount(iCount + 1) + tt(j)
        j = j + 1
        rsUser.movenext
    Loop
    rsUser.Close
    Set rsUser = Nothing
    Response.Write "<tr align='center' class='tdbg2'><td colspan='2'>合计</td>"
    For i = 1 To iCount + 1
        Response.Write "<td>" & arrCount(i) & "</td>"
    Next
Response.Write "</tr>"
Response.Write "</table>"
Response.Write "</form>"
Response.Write "</body></html>"

End Sub
Sub ShowChannelField()
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>统计频道：</td><td><select name='ChannelID'><option value='0'>请选择要统计的频道</option>"
    Dim rsChannel
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelType<=1 and Disabled=0 and ModuleType<4 order by OrderID")
    Do While Not rsChannel.EOF
        If rsChannel("ChannelID") = ChannelID Then
            Response.Write "<option value='" & rsChannel("ChannelID") & "' selected>" & rsChannel("ChannelName") & "</option>"
            Select Case rsChannel("ModuleType")
            Case 1
                TableName = "PE_Article"
            Case 2
                TableName = "PE_Soft"
            Case 3
                TableName = "PE_Photo"
            End Select
        Else
            Response.Write "<option value='" & rsChannel("ChannelID") & "'>" & rsChannel("ChannelName") & "</option>"
        End If
        rsChannel.movenext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "</select></td></tr>"
End Sub

Sub ShowUserField()
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>指定录入者：</td><td><select name='UserName'><option value=''>所有录入者</option>"
    Dim rsUser
    Set rsUser = Conn.Execute("select distinct(Inputer) from PE_Article")
    Do While Not rsUser.EOF
        If rsUser(0) = tUserName Then
            Response.Write "<option value='" & rsUser(0) & "' selected>" & rsUser(0) & "</option>"
        Else
            Response.Write "<option value='" & rsUser(0) & "'>" & rsUser(0) & "</option>"
        End If
        rsUser.movenext
    Loop
    rsUser.Close
    Set rsUser = Nothing
    Response.Write "</select></td></tr>"
End Sub

%>
