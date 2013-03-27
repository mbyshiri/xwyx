<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<!--#include file="../Include/PowerEasy.ConsumeLog.asp"-->
<!--#include file="../Include/PowerEasy.RechargeLog.asp"-->
<!--#include file="../Include/PowerEasy.SMS.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../Include/PowerEasy.MD5_New.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim xmlHttp

'检查管理员操作权限
If AdminPurview > 1 Then
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "User_View")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "User_ModifyInfo")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "User_MofidyPurview")
    arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "User_Lock")
    arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "User_Del")
    arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "User_Update")
    arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "User_Money")
    arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "User_Point")
    arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "User_Valid")
    arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "UserGroup")
    arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Card")
    arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "ConsumeLog")
    arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "RechargeLog")
    arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Message")
    arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "MailList")
    arrPurview(15) = CheckPurview_Other(AdminPurview_Others, "AddPayment")
    For PurviewIndex = 0 To 15
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            Exit For
        End If
    Next
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Response.End
    End If
End If

Dim iGroupType
Response.Write "<html><head><title>会员管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
'call ShowJS_Check()
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <form name='searchmyform' action='Admin_User.asp' method='get'>"
Call ShowPageTitle("会 员 管 理", 10041)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='100' height='30'>快速查找会员：</td>" & vbCrLf
Response.Write "    <td width='687' height='30'>"
Response.Write "          <select size=1 name='SearchType' onChange=""javascript:submit()"">"
Response.Write "          <option value='0' "
If SearchType = 0 Then Response.Write " selected"
Response.Write ">列出所有会员</option>"
Response.Write "          <option value='1' "
If SearchType = 1 Then Response.Write " selected"
Response.Write ">文章最多TOP100</option>"
Response.Write "          <option value='2' "
If SearchType = 2 Then Response.Write " selected"
Response.Write ">文章最少的100个会员</option>"
Response.Write "          <option value='3' "
If SearchType = 3 Then Response.Write " selected"
Response.Write ">最近24小时内登录的会员</option>"
Response.Write "          <option value='4' "
If SearchType = 4 Then Response.Write " selected"
Response.Write ">最近24小时内注册的会员</option>"
Response.Write "          <option value='5' "
If SearchType = 5 Then Response.Write " selected"
Response.Write ">所有被锁住的会员</option>"
Response.Write "          <option value='6' "
If SearchType = 6 Then Response.Write " selected"
Response.Write ">" & PointName & "数大于0的会员</option>"
Response.Write "          <option value='7' "
If SearchType = 7 Then Response.Write " selected"
Response.Write ">积分大于0的会员</option>"
Response.Write "          <option value='8' "
If SearchType = 8 Then Response.Write " selected"
Response.Write ">资金余额大于0的会员</option>"
Response.Write "          <option value='9' "
If SearchType = 9 Then Response.Write " selected"
Response.Write ">资金余额小于0的会员</option>"
Response.Write "          <option value='21' "
If SearchType = 21 Then Response.Write " selected"
Response.Write ">最近一个月没有登录过的会员</option>"
Response.Write "          <option value='22' "
If SearchType = 22 Then Response.Write " selected"
Response.Write ">最近三个月没有登录过的会员</option>"
Response.Write "          <option value='23' "
If SearchType = 23 Then Response.Write " selected"
Response.Write ">最近半年没有登录过的会员</option>"
Response.Write "        </select>"
Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_User.asp'>会员管理首页</a>&nbsp;|&nbsp;<a href='Admin_User.asp?Action=AddUser'>添加新会员</a>&nbsp;|&nbsp;<a href='Admin_User.asp?Action=Update'>更新会员数据</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  </form>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "AddUser", "Modify"
    Call ShowForm
Case "SaveAddUser", "SaveModify"
    Call SaveUser
Case "ModifyPurview"
    Call ModifyPurview
Case "SavePurview"
    Call SavePurview
Case "Show"
    Call Show
Case "Lock", "UnLock"
    Call LockUser
Case "BatchMove"
    Call BatchMove
Case "DoBatchMove"
    Call DoBatchMove
Case "Update"
    Call UpdateUser
Case "DoUpdate"
    Call DoUpdate
Case "ExchangePoint", "AddPoint", "MinusPoint", "ExchangeValid", "AddValid", "MinusValid"
    Call Exchange
Case "DoExchangePoint", "DoAddPoint", "DoMinusPoint", "DoExchangeValid", "DoAddValid", "DoMinusValid"
    Call SaveExchange
Case "AddRemit"
    Call AddIncome(1)
Case "SaveRemit"
    Call SaveRemit
Case "AddPayment"
    Call AddPayment
Case "SavePayment"
    Call SavePayment
Case "AddIncome"
    Call AddIncome(2)
Case "SaveIncome"
    Call SaveIncome
Case "BatchAddMoney", "BatchMinusMoney", "BatchAddPoint", "BatchMinusPoint", "BatchAddValid", "BatchMinusValid", "BatchDel"
    Call Batch
Case "DoBatchAddMoney", "DoBatchMinusMoney", "DoBatchAddPoint", "DoBatchMinusPoint", "DoBatchAddValid", "DoBatchMinusValid", "DoBatchDel"
    Call DoBatch
Case "RegCompany"
    Call RegCompany
Case "Join"
    Call JoinCompany
Case "SaveRegCompany"
    Call SaveRegCompany
Case "Up2Client"
    Call Up2Client
Case "SaveClient"
    Call SaveClient
Case "Agree", "Reject", "RemoveFromCompany", "AddToAdmin", "RemoveFromAdmin"
    Call MemberManage
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim GroupID
    Dim sql, Querysql, rsUserList
    Dim OrderType, strOrderType
    GroupID = PE_CLng(Trim(Request.QueryString("GroupID")))
    OrderType = Trim(Request("OrderType"))
    If OrderType <> "" Then
        OrderType = ReplaceBadChar(OrderType)
    End If
    strFileName = "Admin_User.asp?SearchType=" & SearchType & "&Field=" & strField & "&keyword=" & Keyword & "&GroupID=" & GroupID
    
    Call ShowJS_Main("会员")
    
    Dim rsGroup, i
    i = 1
    Response.Write "<br><table width='100%' class='border' border='0' cellpadding='2' cellspacing='1'><tr class='title'><td>| <a href='Admin_User.asp'>所有会员</a> |"
    Set rsGroup = Conn.Execute("select GroupID,GroupName,GroupIntro from PE_UserGroup order by GroupType asc,GroupID asc")
    Do While Not rsGroup.EOF
        Response.Write " <a href='Admin_User.asp?SearchType=11&GroupID=" & rsGroup(0) & "' title='" & rsGroup(2) & "'>" & rsGroup(1) & "</a> |"
        rsGroup.MoveNext
        i = i + 1
        If i Mod 10 = 0 And Not rsGroup.EOF Then Response.Write "<br>|"
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
    Response.Write "</td></tr></table>"
    
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_User.asp'>注册会员管理</a>&nbsp;&gt;&gt;&nbsp;"

    Select Case OrderType
        Case "Balance"
            strOrderType = "U.Balance desc,"
        Case "Point"
            strOrderType = "U.UserPoint desc,"
        Case "UserExp"
            strOrderType = "U.UserExp desc,"
        Case Else
            strOrderType = ""
    End Select

    Select Case SearchType
    Case 1, 2
        sql = "select top 100 "
    Case 3, 4
        sql = "select "
    Case Else
        If strOrderType = "" Then
            sql = "select top " & MaxPerPage & " "
        Else
            sql = "select "
        End If
    End Select

    sql = sql & " U.*,G.GroupName from PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID "

    Querysql = Querysql & " where 1=1 "
    Select Case SearchType
    Case 0
        Response.Write "所有会员"
    Case 1
        Response.Write "添加信息最多的前100个会员"
    Case 2
        Response.Write "添加信息最少的100个会员"
    Case 3
        Querysql = Querysql & " and datediff(" & PE_DatePart_H & ",LastLoginTime," & PE_Now & ")<25"
        Response.Write "最近24小时内登录的会员"
    Case 4
        Querysql = Querysql & " and datediff(" & PE_DatePart_H & ",RegTime," & PE_Now & ")<25"
        Response.Write "最近24小时内注册的会员"
    Case 5
        Querysql = Querysql & " and U.IsLocked=" & PE_True
        Response.Write "所有被锁住的会员"
    Case 6
        Querysql = Querysql & " and U.UserPoint>0"
        Response.Write PointName & "数大于0的会员"
    Case 7
        Querysql = Querysql & " and U.UserExp>0"
        Response.Write "积分大于0的会员"
    Case 8
        Querysql = Querysql & " and U.Balance>0"
        Response.Write "资金余额大于0的会员"
    Case 9
        Querysql = Querysql & " and U.Balance<0"
        Response.Write "资金余额小于0的会员"
    Case 11
        Querysql = Querysql & " and U.GroupID=" & GroupID & ""
        Response.Write GetGroupName(GroupID)
    Case 21
        Querysql = Querysql & " and datediff(" & PE_DatePart_M & ",LastLoginTime," & PE_Now & ")>=1"
        Response.Write "最近一个月没有登录过的会员"
    Case 22
        Querysql = Querysql & " and datediff(" & PE_DatePart_M & ",LastLoginTime," & PE_Now & ")>=3"
        Response.Write "最近三个月没有登录过的会员"
    Case 23
        Querysql = Querysql & " and datediff(" & PE_DatePart_M & ",LastLoginTime," & PE_Now & ")>=6"
        Response.Write "最近半年没有登录过的会员"

    Case 10
        If Keyword = "" Then
            Response.Write "所有会员"
        Else
            Select Case strField
            Case "UserID"
                If IsNumeric(Keyword) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>会员ID必须是整数</li>"
                Else
                    Querysql = Querysql & " and U.UserID=" & PE_CLng(Keyword)
                    Response.Write "会员ID等于<font color=red> " & PE_CLng(Keyword) & " </font>的会员"
                End If
            Case "UserName"
                Querysql = Querysql & " and U.UserName like '%" & Keyword & "%'"
                Response.Write "用户名中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Email"
                Querysql = Querysql & " and U.Email like '%" & Keyword & "%'"
                Response.Write "电子邮件中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Homepage"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Homepage like '%" & Keyword & "%')"
                Response.Write "个人主页中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "TrueName"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where TrueName like '%" & Keyword & "%')"
                Response.Write "真实姓名中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "IDCard"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where IDCard like '%" & Keyword & "%')"
                Response.Write "身份证号码中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Company"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Company like '%" & Keyword & "%')"
                Response.Write "单位名称中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Address"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Address like '%" & Keyword & "%')"
                Response.Write "联系地址中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Mobile"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Mobile like '%" & Keyword & "%')"
                Response.Write "手机号码中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "OfficePhone"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where OfficePhone like '%" & Keyword & "%')"
                Response.Write "办公电话号码中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "HomePhone"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where HomePhone like '%" & Keyword & "%')"
                Response.Write "家庭电话号码中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "PHS"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where PHS like '%" & Keyword & "%')"
                Response.Write "小灵通号码中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Fax"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Fax like '%" & Keyword & "%')"
                Response.Write "传真号码中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "QQ"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where QQ='" & Keyword & "')"
                Response.Write "QQ号为“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "ICQ"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where ICQ='" & Keyword & "')"
                Response.Write "ICQ号为“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "MSN"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where MSN='" & Keyword & "')"
                Response.Write "MSN帐号为“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Yahoo"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Yahoo='" & Keyword & "')"
                Response.Write "雅虎通帐号为“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "UC"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where UC='" & Keyword & "')"
                Response.Write "UC号为“ <font color=red>" & Keyword & "</font> ”的会员"
            Case "Aim"
                Querysql = Querysql & " and U.ContacterID in (select ContacterID from PE_Contacter where Aim='" & Keyword & "')"
                Response.Write "Aim帐号为“ <font color=red>" & Keyword & "</font> ”的会员"
            Case Else
                Querysql = Querysql & " and U.UserName like '%" & Keyword & "%'"
                Response.Write "用户名中含有“ <font color=red>" & Keyword & "</font> ”的会员"
            End Select
        End If
    End Select
    totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_User U " & Querysql)(0))
    If (SearchType = 1 Or SearchType = 2) And totalPut > 100 Then
        totalPut = 100
    End If
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
    
    Select Case SearchType
    Case 1
        sql = sql & " order by " & strOrderType & "U.PostItems desc,U.UserID desc"
    Case 2
        sql = sql & " order by " & strOrderType & "U.PostItems asc,U.UserID desc"
    Case 3
        sql = sql & Querysql & " order by " & strOrderType & "U.LastLoginTime desc,U.UserID desc"
    Case 4
        sql = sql & Querysql & " order by " & strOrderType & "U.RegTime desc,U.UserID desc"
    Case Else
        If strOrderType = "" Then
            If CurrentPage > 1 Then
                Querysql = Querysql & " and U.UserID < (select min(UserID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.UserID from PE_User U " & Querysql & " order by U.UserID desc) as QueryUser) "
            End If
            sql = sql & Querysql & " order by U.UserID desc"
        Else
            sql = sql & Querysql & " order by " & strOrderType & "U.UserID desc"
        End If
    End Select

    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_User.asp'>"
    Response.Write "      <td>"
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' height='22' align='center'>"
    Response.Write "          <td width='30'>选中</td>"
    Response.Write "          <td width='70'> 用户名</td>"
    Response.Write "          <td>会员类型</td>"
    Response.Write "          <td>所属会员组</td>"
    Response.Write "          <td width='60'><a href='" & strFileName & "&MaxPerPage=" & MaxPerPage & "&OrderType=Balance'>资金余额<a></td>"
    Response.Write "          <td width='60'><a href='" & strFileName & "&MaxPerPage=" & MaxPerPage & "&OrderType=Point'>可用" & PointName & "数</a></td>"
    Response.Write "          <td width='60'>剩余天数</td>"
    Response.Write "          <td width='60'><a href='" & strFileName & "&MaxPerPage=" & MaxPerPage & "&OrderType=UserExp'>可用积分</a></td>"
    Response.Write "          <td width='120'>最后登录IP<br>最后登录时间</td>"
    Response.Write "          <td width='40'>登录<br>次数</td>"
    Response.Write "          <td width='40'>状态</td>"
    Response.Write "          <td width='40'>聚合</td>"
    Response.Write "        </tr>"
    Set rsUserList = Server.CreateObject("Adodb.RecordSet")
    rsUserList.Open sql, Conn, 1, 1
    If rsUserList.BOF And rsUserList.EOF Then
        Response.Write "<tr><td colspan='20' height='50' align='center'>共找到 <font color=red>0</font> 个会员</td></tr>"
    Else
        If (SearchType = 1 Or SearchType = 2 Or SearchType = 3 Or SearchType = 4 Or strOrderType <> "") And CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsUserList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim UserNum
        UserNum = 0
        strFileName = strFileName & "&OrderType=" & OrderType
        Do While Not rsUserList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" align=center>"
            Response.Write "        <td><input name='UserID' type='checkbox' onclick=""unselectall()"" id='UserID' value='" & CStr(rsUserList("UserID")) & "'></td>"
            Response.Write "        <td><a href='Admin_User.asp?Action=Show&UserID=" & rsUserList("UserID") & "'>" & rsUserList("UserName") & "</a></td>"
            Response.Write "        <td>"
            If PE_CLng(rsUserList("UserType")) > 4 Then
                Response.Write arrUserType(0)
            Else
                Response.Write arrUserType(PE_CLng(rsUserList("UserType")))
            End If
            Response.Write "        </td>"
            Response.Write "        <td>" & rsUserList("GroupName") & "</td>"
            Response.Write "        <td align='right'>" & FormatNumber(PE_CDbl(rsUserList("Balance")), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "        <td>"
            If rsUserList("UserPoint") <= 0 Then
                Response.Write "<font color=red>" & rsUserList("UserPoint") & "</font> " & PointUnit & ""
            Else
                If rsUserList("UserPoint") <= 10 Then
                    Response.Write "<font color=blue>" & rsUserList("UserPoint") & "</font> " & PointUnit & ""
                Else
                    Response.Write rsUserList("UserPoint") & " " & PointUnit & ""
                End If
            End If
            Response.Write "</td>"
            Response.Write "<td>"
            If rsUserList("ValidNum") = -1 Then
                Response.Write "无限期"
            Else
                ValidDays = ChkValidDays(rsUserList("ValidNum"), rsUserList("ValidUnit"), rsUserList("BeginTime"))
                If ValidDays <= 0 Then
                    Response.Write "<font color='red'>" & ValidDays & "</font> 天"
                Else
                    Response.Write ValidDays & " 天"
                End If
            End If
            Response.Write "        </td>"
            Response.Write "        <td>" & PE_CLng(rsUserList("UserExp")) & "分</td>"
            Response.Write "        <td>" & rsUserList("LastLoginIP") & "<br>" & rsUserList("LastLoginTime") & "</td>"
            Response.Write "        <td>"
            If rsUserList("LoginTimes") <> "" Then
                Response.Write rsUserList("LoginTimes")
            Else
                Response.Write "0"
            End If
            Response.Write "        </td>"
            Response.Write "        <td>"
            If rsUserList("IsLocked") = True Then
              Response.Write "<font color=red>已锁定</font>"
            Else
              Response.Write "正常"
            End If
            Response.Write "        </td>"
            Response.Write "        <td><a href='Admin_SpaceManage.asp?UserID=" & rsUserList("UserID") & "'>列</a> <a href='Admin_SpaceManage.asp?Action=Add&UserID=" & rsUserList("UserID") & "'>增</a></td>"
            Response.Write "      </tr>"

            UserNum = UserNum + 1
            If UserNum >= MaxPerPage Then Exit Do
            rsUserList.MoveNext
        Loop
    End If
    rsUserList.Close
    Set rsUserList = Nothing
    Response.Write "      </table>"
    If totalPut > 0 Then
        Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "        <tr height='60'>"
        Response.Write "          <td width='200'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form);' value='checkbox'>"
        Response.Write "          选中本页显示的所有会员</td>"
        Response.Write "          <td><input type='hidden' name='Action' value=''>"
        If AdminPurview = 1 Or arrPurview(3) = True Then
            Response.Write "          <input name='Lock' type='submit' value=' 批量锁定 ' onClick=""document.myform.Action.value='Lock';return confirm('确定要锁定选中的会员吗？');"">&nbsp;"
            Response.Write "          <input name='UnLock' type='submit' value=' 批量解锁 ' onClick=""document.myform.Action.value='UnLock';return confirm('确定要将选定的会员解锁吗？');"">&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        If AdminPurview = 1 Or arrPurview(1) = True Then
            Response.Write "          <input name='BatchMove' type='submit' value=' 批量移动 ' onClick=""document.myform.Action.value='BatchMove'"">"
        End If
        If AdminPurview = 1 Or arrPurview(4) = True Then
            Response.Write "          <input name='BatchDel' type='submit' value=' 批量删除 ' onClick=""document.myform.Action.value='BatchDel';"">&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "<br><br>"
        If AdminPurview = 1 Or arrPurview(6) = True Then
            Response.Write "    <input type='submit' name='Submit11' value='发奖金' onClick=""document.myform.Action.value='BatchAddMoney'""> "
            Response.Write "    <input type='submit' name='Submit12' value='扣奖金' onClick=""document.myform.Action.value='BatchMinusMoney'"">&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        If AdminPurview = 1 Or arrPurview(7) = True Then
        Response.Write "    <input type='submit' name='Submit13' value='奖励" & PointName & "' onClick=""document.myform.Action.value='BatchAddPoint'""> "
        Response.Write "    <input type='submit' name='Submit14' value='扣除" & PointName & "' onClick=""document.myform.Action.value='BatchMinusPoint'"">&nbsp;&nbsp;&nbsp;&nbsp;"
    End If
        If AdminPurview = 1 Or arrPurview(8) = True Then
         Response.Write "    <input type='submit' name='Submit15' value='奖励有效期' onClick=""document.myform.Action.value='BatchAddValid'""> "
         Response.Write "    <input type='submit' name='Submit16' value='扣除有效期' onClick=""document.myform.Action.value='BatchMinusValid'"">"
    End If
        Response.Write "        </tr>"
        Response.Write "      </table>"
    End If
    Response.Write "      </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个会员", True)
    End If


    Response.Write "<br>"
    Call ShowSearch
End Sub

Sub ShowSearch()
    Response.Write "<form name='SearchForm' action='Admin_User.asp' method='post'>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='1' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='80'>会员查询：</td>" & vbCrLf
    Response.Write "    <td>" & vbCrLf
    Response.Write "      <select name='Field' size='1'>" & vbCrLf
    Response.Write "        <option value='UserID'>会员ID</option>" & vbCrLf
    Response.Write "        <option value='UserName' selected>用户名</option>" & vbCrLf
    Response.Write "        <option value='Email'>电子邮件</option>" & vbCrLf
    Response.Write "        <option value='Homepage'>个人主页</option>" & vbCrLf
    Response.Write "        <option value='TrueName'>真实姓名</option>" & vbCrLf
    Response.Write "        <option value='IDCard'>身份证号码</option>" & vbCrLf
    Response.Write "        <option value='Company'>单位名称</option>" & vbCrLf
    Response.Write "        <option value='Address'>联系地址</option>" & vbCrLf
    Response.Write "        <option value='Mobile'>手机号码</option>" & vbCrLf
    Response.Write "        <option value='OfficePhone'>办公电话</option>" & vbCrLf
    Response.Write "        <option value='HomePhone'>家庭电话</option>" & vbCrLf
    Response.Write "        <option value='PHS'>小灵通</option>" & vbCrLf
    Response.Write "        <option value='Fax'>传真号码</option>" & vbCrLf
    Response.Write "        <option value='QQ'>QQ号</option>" & vbCrLf
    Response.Write "        <option value='ICQ'>ICQ号</option>" & vbCrLf
    Response.Write "        <option value='MSN'>MSN帐号</option>" & vbCrLf
    Response.Write "        <option value='UC'>UC号</option>" & vbCrLf
    Response.Write "        <option value='Aim'>Aim帐号</option>" & vbCrLf
    Response.Write "      </select>" & vbCrLf
    Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='40'>" & vbCrLf
    Response.Write "      <input type='submit' name='Submit' value='搜 索' id='Submit'>" & vbCrLf
    Response.Write "      <input type='hidden' name='SearchType' value='10'>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "" & vbCrLf
End Sub


Sub ShowForm()
    Dim GroupID, UserType, ContacterID, ClientID, CompanyID
    Dim UserID, UserName, Email, Question, Privacy, UserFace, FaceWidth, FaceHeight, Sign
    Dim rsUser
    If Action = "AddUser" Then
        GroupID = 1
        UserType = 0
        UserID = 0
        ContacterID = 0
        ClientID = 0
        CompanyID = 0
    Else
        If AdminPurview > 1 And arrPurview(1) = False Then
            Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
            Call WriteEntry(6, AdminName, "越权操作")
            Exit Sub
        End If

        UserID = PE_CLng(Trim(Request("UserID")))
        If UserID <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
            Exit Sub
        End If
        Set rsUser = Conn.Execute("select * from PE_User where UserID=" & UserID & "")
        If rsUser.BOF And rsUser.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
            rsUser.Close
            Set rsUser = Nothing
            Exit Sub
        Else
            GroupID = rsUser("GroupID")
            UserType = rsUser("UserType")
            UserName = rsUser("UserName")
            Email = rsUser("Email")
            Question = rsUser("Question")
            Privacy = rsUser("Privacy")
            UserFace = rsUser("UserFace")
            FaceWidth = rsUser("FaceWidth")
            FaceHeight = rsUser("FaceHeight")
            Sign = rsUser("Sign")
            ContacterID = rsUser("ContacterID")
            ClientID = rsUser("ClientID")
            CompanyID = rsUser("CompanyID")
        End If
        rsUser.Close
        Set rsUser = Nothing
    End If
    
    Call PopCalendarInit
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    If Action = "AddUser" Then
        Response.Write "    if(document.myform.UserName.value==''){" & vbCrLf
        Response.Write "        alert('用户名不能为空！');" & vbCrLf
        Response.Write "        document.myform.UserName.focus();" & vbCrLf
        Response.Write "        return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if(document.myform.UserPassword.value==''){" & vbCrLf
        Response.Write "        alert('用户密码不能为空！');" & vbCrLf
        Response.Write "        document.myform.UserPassword.focus();" & vbCrLf
        Response.Write "        return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if(document.myform.Answer.value==''){" & vbCrLf
        Response.Write "        alert('密码提示答案不能为空！');" & vbCrLf
        Response.Write "        document.myform.Answer.focus();" & vbCrLf
        Response.Write "        return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
    End If
    If FoundInArr(RegFields_MustFill, "TrueName", ",") = True Then
        Response.Write "    if(document.myform.TrueName.value==''){" & vbCrLf
        Response.Write "        alert('用户名不能为空！');" & vbCrLf
        Response.Write "        document.myform.TrueName.focus();" & vbCrLf
        Response.Write "        return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
    End If
    Response.Write "    if(document.myform.Question.value==''){" & vbCrLf
    Response.Write "        alert('密码提示问题不能为空！');" & vbCrLf
    Response.Write "        document.myform.Question.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(document.myform.Email.value==''){" & vbCrLf
    Response.Write "        alert('电子邮件不能为空！');" & vbCrLf
    Response.Write "        document.myform.Email.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.myform.Country1.value=frm1.document.regionform.Country.value;" & vbCrLf
    Response.Write "    document.myform.Province1.value=frm1.document.regionform.Province.value;" & vbCrLf
    Response.Write "    document.myform.City1.value=frm1.document.regionform.City.value;" & vbCrLf
    If UserType = 1 Then
        Response.Write "    document.myform.Country2.value=frm2.document.regionform.Country.value;" & vbCrLf
        Response.Write "    document.myform.Province2.value=frm2.document.regionform.Province.value;" & vbCrLf
        Response.Write "    document.myform.City2.value=frm2.document.regionform.City.value;" & vbCrLf
    End If
    Response.Write "}" & vbCrLf

    Response.Write "function SelectClient(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=ClientList','','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        var ss=arr.split('$$$');" & vbCrLf
    Response.Write "        document.myform.ClientName.value=ss[0];" & vbCrLf
    Response.Write "        document.myform.ClientID.value=ss[1];" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_User.asp'>会员管理</a>&nbsp;&gt;&gt;&nbsp;"
    If Action = "AddUser" Then
        Response.Write "添加会员"
    Else
        Response.Write "修改会员信息"
    End If
    Response.Write "</td></tr></table>"
    Response.Write "<form name='myform' id='myform' action='Admin_User.asp' method='post' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>会员信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>联系信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>个人信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>业务信息</td>" & vbCrLf
    If UserType = 1 Then
        Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>单位信息</td>" & vbCrLf
    End If
    Response.Write "            <td>&nbsp;</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "  <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>会 员 组：</td>" & vbCrLf
    Response.Write "        <td width='38%'><select name='GroupID' id='GroupID'>" & GetUserGroup_Option(GroupID) & "</select></td>" & vbCrLf
    Response.Write "        <td width='12%' align='right' class='tdbg5'>会员类别：</td>" & vbCrLf
    Response.Write "        <td width='38%'>"
    If PE_CLng(UserType) > 4 Then
        Response.Write arrUserType(0)
    Else
        Response.Write arrUserType(PE_CLng(UserType))
    End If

    Response.Write "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>用 户 名：</td>" & vbCrLf
    If Action = "AddUser" Then
        Response.Write "        <td width='38%'><input type='text' name='UserName' size='20' maxlength='20' value='" & UserName & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Else
        Response.Write "        <td width='38%'><input type='text' name='UserName' size='20' maxlength='20' value='" & UserName & "' disabled> <font color='#FF0000'>*</font></td>" & vbCrLf
    End If
    Response.Write "        <td width='12%' class='tdbg5' align='right'>用户密码：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='UserPassword' type='text' id='UserPassword' size='20' maxlength='20'>"
    If Action = "AddUser" Then
        Response.Write " <font color='#FF0000'>*</font>"
    Else
        Response.Write " <font color='#FF0000'>不修改请留空</font>"
    End If
    Response.Write "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>提示问题：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='Question' type='text' id='Question' value='" & Question & "'  size='35'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>提示答案：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='Answer' type='text' id='Answer' size='20'>"
    If Action = "AddUser" Then
        Response.Write " <font color='#FF0000'>*</font>"
    Else
        Response.Write " <font color='#FF0000'>不修改请留空</font>"
    End If
    Response.Write "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>电子邮件：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='Email' type='text' id='Email' value='" & Email & "' size='35' maxlength='255'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>隐私设定：</td>" & vbCrLf
    Response.Write "        <td width='38%'><Input type=radio name='Privacy' " & RadioValue(Privacy, 0) & ">全部公开 "
    Response.Write "<Input type=radio name='Privacy'" & RadioValue(Privacy, 1) & ">部分公开 "
    Response.Write "<Input type=radio name='Privacy'" & RadioValue(Privacy, 2) & ">完全保密"
    Response.Write "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>头像地址：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='UserFace' type='text' value='" & UserFace & "' size='35' maxlength='255'></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>头像宽度：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='FaceWidth' type='text' value='" & FaceWidth & "' size='6' maxlength='3'> 象素</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' valign='top'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>签名信息：</td>" & vbCrLf
    Response.Write "        <td width='38%'><textarea name='Sign' cols='35' rows='5' id='Sign'>" & Sign & "</textarea></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>头像高度：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='FaceHeight' type='text' id='FaceHeight' value='" & FaceHeight & "' size='6' maxlength='3'> 象素</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Dim arrEducation, arrIncome
    arrEducation = GetArrFromDictionary("PE_Contacter", "Education")
    arrIncome = GetArrFromDictionary("PE_Contacter", "Income")

    Dim rsContacter, sqlContacter
    Dim TrueName, Title, Company, Department, Position, Operation, CompanyAddress
    Dim Country, Province, City, Address, ZipCode
    Dim Mobile, OfficePhone, Homephone, Fax1, PHS
    Dim HomePage, Email1, QQ, ICQ, MSN, Yahoo, UC, Aim
    Dim IDCard, Birthday, NativePlace, Nation, Sex, Marriage, Income, Education, GraduateFrom, Family
    Dim InterestsOfLife, InterestsOfCulture, InterestsOfAmusement, InterestsOfSport, InterestsOfOther

    sqlContacter = "select * from PE_Contacter where ContacterID=" & ContacterID & ""
    Set rsContacter = Conn.Execute(sqlContacter)
    If rsContacter.BOF And rsContacter.EOF Then
        Sex = -1
        Marriage = 0
        Income = -1
    Else
        TrueName = rsContacter("TrueName")
        Title = rsContacter("Title")
        Country = rsContacter("Country")
        Province = rsContacter("Province")
        City = rsContacter("City")
        ZipCode = rsContacter("ZipCode")
        Address = rsContacter("Address")
        OfficePhone = rsContacter("OfficePhone")
        Homephone = rsContacter("HomePhone")
        Mobile = rsContacter("Mobile")
        Fax1 = rsContacter("Fax")
        PHS = rsContacter("PHS")
        HomePage = rsContacter("HomePage")
        Email1 = rsContacter("Email")
        QQ = rsContacter("QQ")
        ICQ = rsContacter("ICQ")
        MSN = rsContacter("MSN")
        Yahoo = rsContacter("Yahoo")
        UC = rsContacter("UC")
        Aim = rsContacter("Aim")
        IDCard = rsContacter("IDCard")
        Birthday = rsContacter("Birthday")
        NativePlace = rsContacter("NativePlace")
        Nation = rsContacter("Nation")
        Sex = rsContacter("Sex")
        Marriage = rsContacter("Marriage")
        Income = rsContacter("Income")
        Education = rsContacter("Education")
        GraduateFrom = rsContacter("GraduateFrom")
        InterestsOfLife = rsContacter("InterestsOfLife")
        InterestsOfCulture = rsContacter("InterestsOfCulture")
        InterestsOfAmusement = rsContacter("InterestsOfAmusement")
        InterestsOfSport = rsContacter("InterestsOfSport")
        InterestsOfOther = rsContacter("InterestsOfOther")
        Company = rsContacter("Company")
        Department = rsContacter("Department")
        Position = rsContacter("Position")
        Operation = rsContacter("Operation")
        CompanyAddress = rsContacter("CompanyAddress")
    End If
    rsContacter.Close
    Set rsContacter = Nothing
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >真实姓名：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='TrueName' type='text' size='35' maxlength='200' value='" & TrueName & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >称谓：</td>" & vbCrLf
    Response.Write "                        <td><input name='Title' type='text' size='35' maxlength='20' value='" & Title & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>通讯地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <iframe name='frm1' id='frm1' src='../Region.asp?Action=Modify&Country=" & Country & "&Province=" & Province & "&City=" & City & "' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "                            <input name='Country1' type='hidden'> <input name='Province1' type='hidden'> <input name='City1' type='hidden'>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
    Response.Write "                                    <td><input name='Address1' type='text' size='60' maxlength='255' value='" & Address & "'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td align='right' class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
    Response.Write "                                    <td><input name='ZipCode1' type='text' size='35' maxlength='10' value='" & ZipCode & "'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                            </table>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>办公电话：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='OfficePhone' type='text' size='35' maxlength='30' value='" & OfficePhone & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>住宅电话：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='HomePhone' type='text' size='35' maxlength='30' value='" & Homephone & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >移动电话：</td>" & vbCrLf
    Response.Write "                        <td><input name='Mobile' type='text' size='35' maxlength='30' value='" & Mobile & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >传真号码：</td>" & vbCrLf
    Response.Write "                        <td><input name='Fax1' type='text' size='35' maxlength='30' value='" & Fax1 & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >小灵通：</td>" & vbCrLf
    Response.Write "                        <td><input name='PHS' type='text' size='35' maxlength='30' value='" & PHS & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' ></td>" & vbCrLf
    Response.Write "                        <td></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >个人主页：</td>" & vbCrLf
    Response.Write "                        <td><input name='Homepage1' type='text' size='35' maxlength='255' value='" & HomePage & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Email地址：</td>" & vbCrLf
    Response.Write "                        <td><input name='Email1' type='text' size='35' maxlength='255' value='" & Email1 & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >QQ号码：</td>" & vbCrLf
    Response.Write "                        <td><input name='QQ' type='text' size='35' maxlength='20' value='" & QQ & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >MSN帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='MSN' type='text' size='35' maxlength='255' value='" & MSN & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ICQ号码：</td>" & vbCrLf
    Response.Write "                        <td><input name='ICQ' type='text' size='35' maxlength='25' value='" & ICQ & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >雅虎通帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='Yahoo' type='text' size='35' maxlength='255' value='" & Yahoo & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >UC帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='UC' type='text' size='35' maxlength='255' value='" & UC & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Aim帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='Aim' type='text' size='35' maxlength='255' value='" & Aim & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>出生日期：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Birthday' type='text' size='35' maxlength='10' value='" & Birthday & "' onFocus=""PopCalendar.show(document.myform.Birthday, 'yyyy-mm-dd', null, null, null, '11');""></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>证件号码：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='IDCard' type='text' size='35' maxlength='20' value='" & IDCard & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >籍贯：</td>" & vbCrLf
    Response.Write "                        <td><input name='NativePlace' type='text' size='35' maxlength='50' value='" & NativePlace & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >民族：</td>" & vbCrLf
    Response.Write "                        <td><input name='Nation' type='text' size='35' maxlength='50' value='" & Nation & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >性别：</td>" & vbCrLf
    Response.Write "                        <td><input name='Sex' type='radio' value='0' "
    If Sex <= 0 Or Sex > 2 Then Response.Write " checked"
    Response.Write ">保密 <input name='Sex' type='radio' value='1'"
    If Sex = 1 Then Response.Write " checked"
    Response.Write ">男 <input name='Sex' type='radio' value='2'"
    If Sex = 2 Then Response.Write " checked"
    Response.Write ">女</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >婚姻状况：</td>" & vbCrLf
    Response.Write "                        <td><input name='Marriage' type='radio' value='0'"
    If Marriage = 0 Then Response.Write " checked"
    Response.Write ">保密 <input name='Marriage' type='radio' value='1'"
    If Marriage = 1 Then Response.Write " checked"
    Response.Write ">未婚 <input name='Marriage' type='radio' value='2'"
    If Marriage = 2 Then Response.Write " checked"
    Response.Write ">已婚 <input name='Marriage' type='radio' value='3'"
    If Marriage = 3 Then Response.Write " checked"
    Response.Write ">离异</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >学历：</td>" & vbCrLf
    Response.Write "                        <td><select name='Education'>" & Array2Option(arrEducation, Education) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >毕业学校：</td>" & vbCrLf
    Response.Write "                        <td><input name='GraduateFrom' type='text' size='35' maxlength='255' value='" & GraduateFrom & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >生活爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfLife' type='text' size='35' maxlength='255' value='" & InterestsOfLife & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >文化爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfCulture' type='text' size='35' maxlength='255' value='" & InterestsOfCulture & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >娱乐休闲爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfAmusement' type='text' size='35' maxlength='255' value='" & InterestsOfAmusement & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >体育爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfSport' type='text' size='35' maxlength='255' value='" & InterestsOfSport & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >其他爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfOther' type='text' size='35' maxlength='255' value='" & InterestsOfOther & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >月 收 入：</td>" & vbCrLf
    Response.Write "                        <td><select name='Income'>" & Array2Option(arrIncome, Income) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf

    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>单位名称：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Company' type='text' size='35' maxlength='100' value='" & Company & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>所属部门：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Department' type='text' size='35' maxlength='30' value='" & Department & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >职位：</td>" & vbCrLf
    Response.Write "                        <td><input name='Position' type='text' size='35' maxlength='30' value='" & Position & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >负责业务：</td>" & vbCrLf
    Response.Write "                        <td><input name='Operation' type='text' size='35' maxlength='30' value='" & Operation & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >单位地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'><input name='CompanyAddress' type='text' size='35' maxlength='200' value='" & CompanyAddress & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf

    Dim Company2, Phone, Fax2, Country2, Province2, City2, Address2, ZipCode2, HomePage2
    Dim BankOfDeposit, BankAccount, TaxNum, StatusInField, CompanySize, BusinessScope, AnnualSales, ManagementForms, RegisteredCapital
    Dim CompamyPic, CompanyIntro
    Dim arrStatusInField, arrCompanySize, arrManagementForms
    arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
    arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
    arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")
    If UserType = 1 Then
        Dim rsCompany
        Set rsCompany = Conn.Execute("select * from PE_Company where CompanyID=" & CompanyID & "")
        If rsCompany.BOF And rsCompany.EOF Then
            StatusInField = -1
            CompanySize = -1
            ManagementForms = -1
        Else
            Company2 = rsCompany("CompanyName")
            Address2 = rsCompany("Address")
            Country2 = rsCompany("Country")
            Province2 = rsCompany("Province")
            City2 = rsCompany("City")
            ZipCode2 = rsCompany("ZipCode")
            Phone = rsCompany("Phone")
            Fax2 = rsCompany("Fax")
            HomePage2 = rsCompany("Homepage")
            BankOfDeposit = rsCompany("BankOfDeposit")
            BankAccount = rsCompany("BankAccount")
            TaxNum = rsCompany("TaxNum")
            StatusInField = rsCompany("StatusInField")
            CompanySize = rsCompany("CompanySize")
            BusinessScope = rsCompany("BusinessScope")
            AnnualSales = rsCompany("AnnualSales")
            ManagementForms = rsCompany("ManagementForms")
            RegisteredCapital = rsCompany("RegisteredCapital")
            CompamyPic = rsCompany("CompamyPic")
            CompanyIntro = rsCompany("CompanyIntro")
        End If
        rsCompany.Close
        Set rsCompany = Nothing
        Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>单位名称：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Company2' type='text' size='35' maxlength='30' value='" & Company2 & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'></td>" & vbCrLf
        Response.Write "                        <td width='38%'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>通讯地址：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & vbCrLf
        Response.Write "                            <iframe name='frm2' id='frm2' src='../Region.asp?Action=Modify&Country=" & Country2 & "&Province=" & Province2 & "&City=" & City2 & "' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
        Response.Write "                            <input name='Country2' type='hidden'> <input name='Province2' type='hidden'> <input name='City2' type='hidden'>" & vbCrLf
        Response.Write "                        </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & vbCrLf
        Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
        Response.Write "                                <tr class='tdbg'>" & vbCrLf
        Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
        Response.Write "                                    <td><input name='Address2' type='text' size='60' maxlength='255' value='" & Address2 & "'></td>" & vbCrLf
        Response.Write "                                </tr>" & vbCrLf
        Response.Write "                                <tr class='tdbg'>" & vbCrLf
        Response.Write "                                    <td align='right' class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
        Response.Write "                                    <td><input name='ZipCode2' type='text' size='35' maxlength='10' value='" & ZipCode2 & "'></td>" & vbCrLf
        Response.Write "                                </tr>" & vbCrLf
        Response.Write "                            </table>" & vbCrLf
        Response.Write "                        </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>联系电话：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Phone' type='text' size='35' maxlength='30' value='" & Phone & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>传真号码：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Fax2' type='text' size='35' maxlength='30' value='" & Fax2 & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>开户银行：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='BankOfDeposit' type='text' size='35' maxlength='255' value='" & BankOfDeposit & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>银行帐号：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='BankAccount' type='text' size='35' maxlength='255' value='" & BankAccount & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >税号：</td>" & vbCrLf
        Response.Write "                        <td><input name='TaxNum' type='text' size='35' maxlength='20' value='" & TaxNum & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >网址：</td>" & vbCrLf
        Response.Write "                        <td><input name='Homepage2' type='text' size='35' maxlength='100' value='" & HomePage2 & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >行业地位：</td>" & vbCrLf
        Response.Write "                        <td><select name='StatusInField'>" & Array2Option(arrStatusInField, StatusInField) & "</select></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司规模：</td>" & vbCrLf
        Response.Write "                        <td><select name='CompanySize'>" & Array2Option(arrCompanySize, CompanySize) & "</select></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >业务范围：</td>" & vbCrLf
        Response.Write "                        <td><input name='BusinessScope' type='text' size='35' maxlength='255' value='" & BusinessScope & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >年销售额：</td>" & vbCrLf
        Response.Write "                        <td><input name='AnnualSales' type='text' size='15' maxlength='20' value='" & AnnualSales & "'> 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >经营状态：</td>" & vbCrLf
        Response.Write "                        <td><select name='ManagementForms'>" & Array2Option(arrManagementForms, ManagementForms) & "</select></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >注册资本：</td>" & vbCrLf
        Response.Write "                        <td><input name='RegisteredCapital' type='text' size='15' maxlength='20' value='" & RegisteredCapital & "'> 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司照片：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'><input name='CompamyPic' type='text' size='35' maxlength='255' value='" & CompamyPic & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司简介：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'><textarea name='CompanyIntro' cols='75' rows='5' id='CompanyIntro'>" & PE_ConvertBR(CompanyIntro) & "</textarea></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                </tbody>" & vbCrLf
    End If


    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1'><tr align='center'><td height='40'>" & vbCrLf
    Response.Write "    <input type='hidden' name='action' value='Save" & Action & "'><input type='hidden' name='UserID' value='" & UserID & "'>" & vbCrLf
    Response.Write "    <input type='submit' name='Submit' value='保存会员信息'>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub ModifyPurview()
    If AdminPurview > 1 And arrPurview(2) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID
    Dim rsUser, sqlUser
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    End If
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 1
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "    if(document.myform.SpecialPermission[1].checked==true){" & vbCrLf
    
    Dim rsChannel, ChannelDir
    Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType<>4 And ModuleType<>5  And ModuleType<>7 And ModuleType<>8 And Disabled=" & PE_False & " ORDER BY OrderID")
    If Not (rsChannel.BOF And rsChannel.EOF) Then
        Do While Not rsChannel.EOF
            ChannelDir = rsChannel(0)
            Response.Write "    if(document.myform." & ChannelDir & "purview[1].checked==true){" & vbCrLf
            Response.Write "        document.myform.arrClass_Browse_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "        document.myform.arrClass_View_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "        document.myform.arrClass_Input_" & ChannelDir & ".value='';" & vbCrLf
            Response.Write "        for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Browse.length;i++){" & vbCrLf
            Response.Write "            if(frm" & ChannelDir & ".document.myform.Purview_Browse[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_Browse[i].checked==true){" & vbCrLf
            Response.Write "                if(document.myform.arrClass_Browse_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                    document.myform.arrClass_Browse_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Browse[i].value;" & vbCrLf
            Response.Write "                else" & vbCrLf
            Response.Write "                    document.myform.arrClass_Browse_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Browse[i].value;" & vbCrLf
            Response.Write "            }" & vbCrLf
            Response.Write "        }" & vbCrLf
            Response.Write "        for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_View.length;i++){" & vbCrLf
            Response.Write "            if(frm" & ChannelDir & ".document.myform.Purview_View[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_View[i].checked==true){" & vbCrLf
            Response.Write "                if(document.myform.arrClass_View_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                    document.myform.arrClass_View_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
            Response.Write "                else" & vbCrLf
            Response.Write "                    document.myform.arrClass_View_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
            Response.Write "            }" & vbCrLf
            Response.Write "        }" & vbCrLf
            Response.Write "        for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Input.length;i++){" & vbCrLf
            Response.Write "            if(frm" & ChannelDir & ".document.myform.Purview_Input[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_Input[i].checked==true){" & vbCrLf
            Response.Write "                if(document.myform.arrClass_Input_" & ChannelDir & ".value=='')" & vbCrLf
            Response.Write "                    document.myform.arrClass_Input_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
            Response.Write "                else" & vbCrLf
            Response.Write "                    document.myform.arrClass_Input_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
            Response.Write "            }" & vbCrLf
            Response.Write "        }" & vbCrLf
            Response.Write "    }" & vbCrLf
            rsChannel.MoveNext
        Loop
    End If
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Set rsChannel = Nothing
    Response.Write "</script>" & vbCrLf

    Response.Write "<form name='myform' id='myform' action='Admin_User.asp' method='post' onSubmit='CheckSubmit();'>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td height='22' colspan='6'>修 改 会 员 权 限</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' width='15%' class='tdbg5'>用户名：</td><td>" & rsUser("UserName") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' width='15%' class='tdbg5'>会员组别：</td><td>" & GetGroupName(rsUser("GroupID")) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td align='right' width='15%' class='tdbg5'>会员权限：</td><td><input name='SpecialPermission' type='radio' value='0' onClick=""tablep.style.display='none'"""
    If rsUser("SpecialPermission") = False Then Response.Write " checked"
    Response.Write "> 会员组默认 <input type='radio' name='SpecialPermission' value='1' onClick=""tablep.style.display='block'"""
    If rsUser("SpecialPermission") = True Then Response.Write " checked"
    Response.Write "> 单独设置具体权限</td></tr>" & vbCrLf

    If rsUser("SpecialPermission") = True Then
        UserSetting = Split(rsUser("UserSetting") & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        arrClass_Browse = rsUser("arrClass_Browse")
        arrClass_View = rsUser("arrClass_View")
        arrClass_Input = rsUser("arrClass_Input")
    End If
    Response.Write "                <tbody id='tablep'"
    If rsUser("SpecialPermission") = False Then Response.Write "style='display:none;'"
    Response.Write ">" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>发布权限：</td>" & vbCrLf
    Response.Write "      <td><input name='UserSetting1' type='checkbox' " & RadioValue(PE_CLng(UserSetting(1)), 1) & ">在发布信息需要审核的频道，此组会员发布信息不需要审核<br>" & vbCrLf
    Response.Write "<input name='UserSetting2' type='checkbox' " & RadioValue(PE_CLng(UserSetting(2)), 1) & ">可以修改和删除已审核的（自己的）信息<br>" & vbCrLf
    Response.Write "<input name='UserSetting21' type='checkbox' " & RadioValue(PE_CLng(UserSetting(21)), 1) & ">发表信息时可以设置标题前缀<br>" & vbCrLf
    Response.Write "<input name='UserSetting22' type='checkbox' " & RadioValue(PE_CLng(UserSetting(22)), 1) & ">发表信息时可以设置是否显示评论链接<br>" & vbCrLf
    Response.Write "<input name='UserSetting23' type='checkbox' " & RadioValue(PE_CLng(UserSetting(23)), 1) & ">发表信息时可以设置转向链接<br>" & vbCrLf
    Response.Write "<input name='UserSetting24' type='checkbox' " & RadioValue(PE_CLng(UserSetting(24)), 1) & ">发表信息时HTML编辑器为高级模式（默认为简洁模式）<br>" & vbCrLf
    Response.Write "每天最多发布<input name='UserSetting3' type='text' value='" & UserSetting(3) & "' size='6' maxlength='6' style='text-align: center;'>条信息（不想限制请设置为<b>0</b>）。<br>"
    Response.Write "发布信息时获取积分为栏目设置的<input name='UserSetting4' type='text' value='" & UserSetting(4) & "' size='5' maxlength='5' style='text-align: center;'>倍<br>"
    
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>评论权限：</td>" & vbCrLf
    Response.Write "         <td><input name='UserSetting5' type='checkbox' " & RadioValue(PE_CLng(UserSetting(5)), 1) & ">在禁止发表评论的栏目里仍然可发表评论<br><input name='UserSetting6' type='checkbox' " & RadioValue((UserSetting(6)), 1) & ">在评论需要审核的栏目里发表评论不需要审核</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>短消息权限：</td>" & vbCrLf
    Response.Write "      <td>每次最多可同时向<input name='UserSetting7' type='text' value='" & UserSetting(7) & "' size='4' maxlength='4' style='text-align: center;'>人发送短消息（如果为0，则不允许发送短消息）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>收藏夹权限：</td>" & vbCrLf
    Response.Write "      <td>会员收藏夹内最多可收录<input name='UserSetting8' type='text' value='" & UserSetting(8) & "' size='5' maxlength='5' style='text-align: center;'>条信息（如果为0，则没有收藏权限）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>上传文件权限：</td>" & vbCrLf
    Response.Write "      <td><input name='UserSetting9' type='checkbox' " & RadioValue(PE_CLng(UserSetting(9)), 1) & ">允许在开放上传的频道上传文件<br>最大允许上传<input name='UserSetting10' type='text' value='" & UserSetting(10) & "' size='5' style='text-align: center;'>K的文件（当所设置值大于某一频道的设置时，以频道设置为准。）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>商城权限：</td>" & vbCrLf
    Response.Write "      <td>购物时可以享受的折扣率：<input name='UserSetting11' type='text' value='" & UserSetting(11) & "' size='5' maxlength='5' style='text-align: center;'>%<br><input name='UserSetting12' type='checkbox' " & RadioValue(PE_CLng(UserSetting(12)), 1) & ">是否可以享受折上折优惠（对指定会员价的商品无效）<br> 允许透支的最大额度：<input name='UserSetting13' type='text' value='" & UserSetting(13) & "' size='6' maxsize='6' style='text-align: center;'>元人民币" & vbCrLf
    Response.Write "        <br><input name='UserSetting30' type='checkbox' " & RadioValue(PE_CLng(UserSetting(30)), 1) & ">是否可以批发商品<br>"
    Response.Write "    </td></tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>计费方式：</td>" & vbCrLf
    Response.Write "      <td><input name='UserSetting14' type='radio' " & RadioValue(PE_CLng(UserSetting(14)), 0) & ">只判断" & PointName & "：有" & PointName & "时，即使有效期已经到期，仍可以查看收费内容；" & PointName & "用完后，即使有效期没有到期，也不能查看收费内容。<br>" & vbCrLf
    Response.Write "          <input type='radio' name='UserSetting14' " & RadioValue(PE_CLng(UserSetting(14)), 1) & ">只判断有效期：只要在有效期内，" & PointName & "用完后仍可以查看收费内容；过期后，即使会员有" & PointName & "也不能查看收费内容。<br>" & vbCrLf
    Response.Write "          <input type='radio' name='UserSetting14' " & RadioValue(PE_CLng(UserSetting(14)), 2) & ">同时判断" & PointName & "和有效期：" & PointName & "用完或有效期到期后，就不可查看收费内容。<br>" & vbCrLf
    Response.Write "          <input type='radio' name='UserSetting14' " & RadioValue(PE_CLng(UserSetting(14)), 3) & ">同时判断" & PointName & "和有效期：" & PointName & "用完并且有效期到期后，才不能查看收费内容。" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>扣" & PointName & "方式：</td>" & vbCrLf
    Response.Write "      <td><input name='UserSetting15' type='radio' " & RadioValue(PE_CLng(UserSetting(15)), 0) & ">有效期内，查看收费内容不扣" & PointName & "，也不做记录。<br>" & vbCrLf
    Response.Write "          <input type='radio' name='UserSetting15' " & RadioValue(PE_CLng(UserSetting(15)), 1) & ">有效期内，查看收费内容不扣" & PointName & "，但做记录。<br>" & vbCrLf
    Response.Write "          <input type='radio' name='UserSetting15' " & RadioValue(PE_CLng(UserSetting(15)), 2) & ">有效期内，查看收费内容也扣" & PointName & "。<br>" & vbCrLf
    Response.Write "有效期内，总共可以看<input name='UserSetting16' type='text' value='" & UserSetting(16) & "' size='10' maxlength='10' style='text-align: center;'> 条收费信息（如果为0，则不限制）<br>" & vbCrLf
    Response.Write "有效期内，每天最多可以看<input name='UserSetting17' type='text' value='" & UserSetting(17) & "' size='10' maxlength='10' style='text-align: center;'> 条收费信息（如果为0，则不限制）" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>自助充值：</td>" & vbCrLf
    Response.Write "         <td><input name='UserSetting18' type='checkbox' " & RadioValue(PE_CLng(UserSetting(18)), 1) & ">允许自助兑换" & PointName & "<br><input name='UserSetting19' type='checkbox' " & RadioValue(PE_CLng(UserSetting(19)), 1) & ">允许自助兑换有效期<br><input name='UserSetting20' type='checkbox' " & RadioValue(PE_CLng(UserSetting(20)), 1) & ">允许将" & PointName & "赠送给他人</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>聚合空间：</td>" & vbCrLf
    Response.Write "         <td><input name='UserSetting25' type='checkbox' " & RadioValue(PE_CLng(UserSetting(25)), 1) & ">启用聚合空间<br>" & vbCrLf
    Response.Write "         <input name='UserSetting26' type='checkbox' " & RadioValue(PE_CLng(UserSetting(26)), 1) & ">申请时无须管理员审核<br>" & vbCrLf
    Response.Write " 聚合空间容量为:<input name='UserSetting27' type='text' value=' " & PE_CLng(UserSetting(27)) & "' size='4' maxlength='10' style='text-align: center;'>M<br>" & vbCrLf
    Response.Write "         <input name='UserSetting28' type='checkbox' " & RadioValue(PE_CLng(UserSetting(28)), 1) & ">用户可以自主更换皮肤" & vbCrLf
    Response.Write "    </td></tr>" & vbCrLf
    If AdminPurview = 1 Then
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td colspan='3'>" & vbCrLf
        Response.Write "        <table width='100%' border='0' cellspacing='10' cellpadding='0'>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td colspan='2' align='center'>频 道 权 限 详 细 设 置</td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
 

       Dim arrPurviews, IsNoPurview
        arrPurviews = rsUser("arrClass_Browse") & "," & rsUser("arrClass_View") & "," & rsUser("arrClass_Input")
       Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType<>4 And ModuleType<>5 And ModuleType<>7 And ModuleType<>8 AND Disabled=" & PE_False & " ORDER BY OrderID")
        Do While Not rsChannel.EOF
            IsNoPurview = FoundInArr(arrPurviews, rsChannel("ChannelDir") & "none", ",")
            Response.Write "          <tr valign='top'>" & vbCrLf
            Response.Write "           <td><fieldset>" & vbCrLf
            Response.Write "   <legend>此会员组在【<font color='red'>" & rsChannel("ChannelName") & "</font>】频道的权限：</legend>" & vbCrLf
            Response.Write "    <table width='100%' cellspacing='1'>" & vbCrLf
            Response.Write "        <tr class='tdbg'>" & vbCrLf
            Response.Write "                <td width='50%'><input type='radio' name='" & rsChannel("ChannelDir") & "purview' onClick=""table" & rsChannel("ChannelID") & ".style.display='none'"""
            If IsNoPurview = True Then Response.Write "checked"
            Response.Write ">无任何权限(开放栏目除外)"
            Response.Write "&nbsp;&nbsp;<input type='radio' name='" & rsChannel("ChannelDir") & "purview' onClick=""table" & rsChannel("ChannelID") & ".style.display='block'"""
            If IsNoPurview = False Then Response.Write "checked"
            Response.Write ">设置会员在该频道的权限</td>" & vbCrLf
            Response.Write "             <td></td>" & vbCrLf
            Response.Write "        <tr class='tdbg' id='table" & rsChannel("ChannelID") & "' style='display:"
            If IsNoPurview = True Then
                Response.Write "none"
            Else
                Response.Write "block"
            End If
            Response.Write "'>" & vbCrLf
            Response.Write "         <td width='50%'>" & vbCrLf
            Response.Write "         <iframe id='frm" & rsChannel("ChannelDir") & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=User&Action=Modify&ChannelID=" & rsChannel("ChannelID") & "&UserID=" & rsUser("UserID") & "'></iframe>" & vbCrLf
            Response.Write "         <input name='arrClass_Browse_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Browse_" & rsChannel("ChannelDir") & "' value='" & rsChannel("ChannelDir") & "none'>" & vbCrLf
            Response.Write "         <input name='arrClass_View_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_View_" & rsChannel("ChannelDir") & "' value='" & rsChannel("ChannelDir") & "none'>" & vbCrLf
            Response.Write "         <input name='arrClass_Input_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Input_" & rsChannel("ChannelDir") & "' value='" & rsChannel("ChannelDir") & "none'></td>" & vbCrLf
            Response.Write "         <td width='50%'><font color='#0000FF'>注：</font><br>1、栏目权限采用继承制度，即在某一栏目拥有某项权限，则在此栏目的所有子栏目中都拥有这项权限，并可在子栏目中指定更多的权限。<br>2、灰色并且选中的项目，说明该栏目为开放栏目，会员组在此栏目拥有浏览和查看权限。<br><br><font color='red'>权限含义：</font><br>浏览－－指可以浏览此栏目的信息列表<br>查看－－指可以查看此栏目中的信息的内容<br>发布－－指可以在此栏目中发布信息</td>" & vbCrLf
            Response.Write "        </tr>" & vbCrLf
            Response.Write "   </table>" & vbCrLf
            Response.Write "   </fieldset></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            rsChannel.MoveNext
        Loop
        rsChannel.Close
        Set rsChannel = Nothing

        Response.Write "             </table>" & vbCrLf
        Response.Write "           </td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "      </tbody>" & vbCrLf
    End If	
    Response.Write "        <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "            <td colspan='4' height='40'>" & vbCrLf
    Response.Write "                <input type='hidden' name='action' value='SavePurview'><input type='hidden' name='UserID' value='" & UserID & "'>" & vbCrLf
    Response.Write "                <input type='submit' name='Submit' value='保存修改结果'>" & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub Show()
    Dim UserID, UserName, ClientID
    Dim rsUser, sqlUser
    Dim ValidDays
    UserID = PE_CLng(Trim(Request("UserID")))
    UserName = ReplaceBadChar(Trim(Request("UserName")))
    If UserID <= 0 And UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    End If
    If UserID <> 0 Then
        sqlUser = "select * from PE_User where UserID=" & UserID
    Else
        sqlUser = "select * from PE_User where UserName='" & UserName & "'"
    End If
    Set rsUser = Conn.Execute(sqlUser)
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If
    UserID = PE_CLng(rsUser("UserID"))
    ClientID = PE_CLng(rsUser("ClientID"))
    ValidDays = ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime"))

    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_User.asp'>会员管理</a>&nbsp;&gt;&gt;&nbsp;查看会员信息：" & rsUser("UserName") & "</td></tr></table>"
    Response.Write "<br><table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>会员信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>联系信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>个人信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>业务信息</td>" & vbCrLf
    If rsUser("UserType") > 0 And rsUser("UserType") < 4 Then
        Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>单位信息</td>" & vbCrLf
        Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>单位成员</td>" & vbCrLf
    End If
    Response.Write "            <td>&nbsp;"
    If AdminPurview = 1 Or arrPurview(1) = True Then
        If rsUser("UserType") = 0 Then
            Response.Write "<input type='button' name='Reg' value='升级为企业会员' onclick=""window.location.href='Admin_User.asp?Action=RegCompany&UserID=" & rsUser("UserID") & "'""> "
        End If
    End If
    If AdminPurview = 1 Or arrPurview(5) = True Then
        If rsUser("ClientID") = 0 Then
            Response.Write "<input type='button' name='Up2Client' value=' 升级为客户 ' onclick=""window.location.href='Admin_User.asp?Action=Up2Client&UserID=" & rsUser("UserID") & "'"">"
        End If
    End If
    Response.Write " </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "  <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>用 户 名：</td><td>" & rsUser("UserName") & "</td>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>邮箱地址：</td><td width='38%'>" & rsUser("Email") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>会员组别：</td><td width='38%'>" & GetGroupName(rsUser("GroupID")) & "</td>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>会员类别：</td><td width='38%'>"
    If PE_CLng(rsUser("UserType")) > 4 Then
        Response.Write arrUserType(0)
    Else
        Response.Write arrUserType(PE_CLng(rsUser("UserType")))
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>资金余额：</td><td width='38%'>" & rsUser("Balance") & "元</td>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>可用" & PointName & "数：</td><td width='38%'>" & rsUser("UserPoint") & "" & PointUnit & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>可用积分：</td><td width='38%'>" & rsUser("UserExp") & "分</td>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>剩余天数：</td><td width='38%'>"
    If rsUser("ValidNum") = -1 Then
        Response.Write "无限期"
    Else
        Response.Write ValidDays & "天"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>待签文章：</td><td width='38%'>"
    
    If rsUser("UnsignedItems") <> "" Then
        Dim UnsignedItemNum, arrUser
        arrUser = Split(rsUser("UnsignedItems"), ",")
        UnsignedItemNum = UBound(arrUser) + 1
        Response.Write " <b><font color=red>" & UnsignedItemNum & "</font></b> 篇"
    Else
        Response.Write " <b><font color=gray>0</font></b> 篇"
    End If
    Response.Write "</td>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>待阅短信：</td><td width='38%'>"

    If rsUser("UnreadMsg") <> "" And PE_CLng(rsUser("UnreadMsg")) > 0 Then
        Response.Write " <b><font color=red>" & rsUser("UnreadMsg") & "</font></b> 条"
    Else
        Response.Write " <b><font color=gray>0</font></b> 条"
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>会员权限：</td><td width='38%'>"
    If rsUser("SpecialPermission") = True Then
        Response.Write "自定义"
    Else
        Response.Write "会员组默认"
    End If
    Response.Write "</td>"
    If rsUser("UserType") = 1 Or rsUser("UserType") = 2 Then
        Response.Write "    <td width='12%' align='right' class='tdbg5'>待审成员：</td><td width='38%'>"
        Dim trs
        Set trs = Conn.Execute("select count(0) from PE_User where UserType=4 and CompanyID=" & rsUser("CompanyID") & "")
        If trs(0) > 0 Then
            Response.Write " <b><font color=red>" & trs(0) & "</font></b> 名"
        Else
            Response.Write " <b><font color=gray>0</font></b> 名"
        End If
        Response.Write "</td>"
    Else
        Response.Write "    <td width='12%' align='right' class='tdbg5'></td><td width='38%'></td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='12%' align='right' class='tdbg5'>注册日期：</td><td width='38%'>" & rsUser("RegTime") & "</td>" & vbCrLf
    Response.Write "    <td width='12%' align='right' class='tdbg5'>加入日期：</td><td width='38%'>" & rsUser("JoinTime") & "</td>" & vbCrLf
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='12%' align='right' class='tdbg5'>最后登录时间：</td><td width='38%'>" & rsUser("LastLoginTime") & "</td>" & vbCrLf
    Response.Write "    <td width='12%' align='right' class='tdbg5'>最后登录IP：</td><td width='38%'>" & rsUser("LastLoginIP") & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf



    Dim rsContacter, sqlContacter
    Dim TrueName, Title, Company, Department, Position, Operation, CompanyAddress
    Dim Country, Province, City, Address, ZipCode
    Dim Mobile, OfficePhone, Homephone, Fax, PHS
    Dim HomePage, Email, QQ, ICQ, MSN, Yahoo, UC, Aim
    Dim IDCard, Birthday, NativePlace, Nation, Sex, Marriage, Income, Education, GraduateFrom, Family
    Dim InterestsOfLife, InterestsOfCulture, InterestsOfAmusement, InterestsOfSport, InterestsOfOther
    Dim arrEducation, arrIncome
    arrEducation = GetArrFromDictionary("PE_Contacter", "Education")
    arrIncome = GetArrFromDictionary("PE_Contacter", "Income")


    sqlContacter = "select * from PE_Contacter where ContacterID=" & rsUser("ContacterID") & ""
    Set rsContacter = Conn.Execute(sqlContacter)
    If rsContacter.BOF And rsContacter.EOF Then
        Sex = -1
        Marriage = 0
        Income = -1
    Else
        TrueName = rsContacter("TrueName")
        Title = rsContacter("Title")
        Country = rsContacter("Country")
        Province = rsContacter("Province")
        City = rsContacter("City")
        ZipCode = rsContacter("ZipCode")
        Address = rsContacter("Address")
        OfficePhone = rsContacter("OfficePhone")
        Homephone = rsContacter("HomePhone")
        Mobile = rsContacter("Mobile")
        Fax = rsContacter("Fax")
        PHS = rsContacter("PHS")
        HomePage = rsContacter("HomePage")
        Email = rsContacter("Email")
        QQ = rsContacter("QQ")
        ICQ = rsContacter("ICQ")
        MSN = rsContacter("MSN")
        Yahoo = rsContacter("Yahoo")
        UC = rsContacter("UC")
        Aim = rsContacter("Aim")
        IDCard = rsContacter("IDCard")
        Birthday = rsContacter("Birthday")
        NativePlace = rsContacter("NativePlace")
        Nation = rsContacter("Nation")
        Sex = rsContacter("Sex")
        Marriage = rsContacter("Marriage")
        Income = rsContacter("Income")
        Education = rsContacter("Education")
        GraduateFrom = rsContacter("GraduateFrom")
        InterestsOfLife = rsContacter("InterestsOfLife")
        InterestsOfCulture = rsContacter("InterestsOfCulture")
        InterestsOfAmusement = rsContacter("InterestsOfAmusement")
        InterestsOfSport = rsContacter("InterestsOfSport")
        InterestsOfOther = rsContacter("InterestsOfOther")
        
        Company = rsContacter("Company")
        Department = rsContacter("Department")
        Position = rsContacter("Position")
        Operation = rsContacter("Operation")
        CompanyAddress = rsContacter("CompanyAddress")
    End If
    rsContacter.Close
    Set rsContacter = Nothing
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >真实姓名：</td>" & vbCrLf
    Response.Write "                        <td>" & TrueName & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >称谓：</td>" & vbCrLf
    Response.Write "                        <td>" & Title & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>国家/地区：</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Country & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>省/市：</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Province & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >市/县/区：</td>" & vbCrLf
    Response.Write "                        <td>" & City & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
    Response.Write "                        <td>" & ZipCode & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & Address & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >办公电话：</td>" & vbCrLf
    Response.Write "                        <td>" & OfficePhone & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >住宅电话：</td>" & vbCrLf
    Response.Write "                        <td>" & Homephone & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >移动电话：</td>" & vbCrLf
    Response.Write "                        <td>" & Mobile & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >传真号码：</td>" & vbCrLf
    Response.Write "                        <td>" & Fax & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >小灵通：</td>" & vbCrLf
    Response.Write "                        <td>" & PHS & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' ></td>" & vbCrLf
    Response.Write "                        <td></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >个人主页：</td>" & vbCrLf
    Response.Write "                        <td>" & HomePage & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Email地址：</td>" & vbCrLf
    Response.Write "                        <td>" & Email & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >QQ号码：</td>" & vbCrLf
    Response.Write "                        <td>" & QQ & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >MSN帐号：</td>" & vbCrLf
    Response.Write "                        <td>" & MSN & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ICQ号码：</td>" & vbCrLf
    Response.Write "                        <td>" & ICQ & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >雅虎通帐号：</td>" & vbCrLf
    Response.Write "                        <td>" & Yahoo & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >UC帐号：</td>" & vbCrLf
    Response.Write "                        <td>" & UC & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Aim帐号：</td>" & vbCrLf
    Response.Write "                        <td>" & Aim & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf

    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>出生日期：</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Birthday & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>证件号码：</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & IDCard & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >籍贯：</td>" & vbCrLf
    Response.Write "                        <td>" & NativePlace & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >民族：</td>" & vbCrLf
    Response.Write "                        <td>" & Nation & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >性别：</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(Array("保密", "男", "女"), Sex) & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >婚姻状况：</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(Array("保密", "未婚", "已婚", "离异"), Marriage) & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >学历：</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(arrEducation, Education) & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >毕业学校：</td>" & vbCrLf
    Response.Write "                        <td>" & GraduateFrom & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >生活爱好：</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfLife & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >文化爱好：</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfCulture & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >娱乐休闲爱好：</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfAmusement & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >体育爱好：</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfSport & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >其他爱好：</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfOther & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >月 收 入：</td>" & vbCrLf
    Response.Write "                        <td>"
    If Income > 6 Then
        Response.Write Income
    Else
        Response.Write GetArrItem(arrIncome, Income)
    End If
    Response.Write "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf
    
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>单位名称：</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Company & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>所属部门：</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Department & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >职位：</td>" & vbCrLf
    Response.Write "                        <td>" & Position & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >负责业务：</td>" & vbCrLf
    Response.Write "                        <td>" & Operation & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >单位地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & CompanyAddress & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf


    If rsUser("UserType") > 0 And rsUser("UserType") < 4 Then
        Dim CompanyName, Phone, Fax2, HomePage2
        Dim BankOfDeposit, BankAccount, TaxNum, StatusInField, CompanySize, BusinessScope, AnnualSales, ManagementForms, RegisteredCapital
        Dim CompanyIntro, CompamyPic
        Dim arrStatusInField, arrCompanySize, arrManagementForms
        arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
        arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
        arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")
        Dim rsCompany
        Set rsCompany = Conn.Execute("select * from PE_Company where CompanyID=" & rsUser("CompanyID") & "")
        If rsCompany.BOF And rsCompany.EOF Then
            Country = ""
            Province = ""
            City = ""
            ZipCode = ""
            Address = ""
            StatusInField = -1
            CompanySize = -1
            ManagementForms = -1
        Else
            CompanyName = rsCompany("CompanyName")
            Address = rsCompany("Address")
            Country = rsCompany("Country")
            Province = rsCompany("Province")
            City = rsCompany("City")
            ZipCode = rsCompany("ZipCode")
            Phone = rsCompany("Phone")
            Fax2 = rsCompany("Fax")
            BankOfDeposit = rsCompany("BankOfDeposit")
            BankAccount = rsCompany("BankAccount")
            TaxNum = rsCompany("TaxNum")
            StatusInField = rsCompany("StatusInField")
            CompanySize = rsCompany("CompanySize")
            BusinessScope = rsCompany("BusinessScope")
            AnnualSales = rsCompany("AnnualSales")
            ManagementForms = rsCompany("ManagementForms")
            RegisteredCapital = rsCompany("RegisteredCapital")
            HomePage2 = rsCompany("Homepage")
            CompanyIntro = rsCompany("CompanyIntro")
            CompamyPic = rsCompany("CompamyPic")
        End If
        rsCompany.Close
        Set rsCompany = Nothing
        Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>单位名称：</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & CompanyName & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>联系地址：</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & Address & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>国家/地区：</td>" & vbCrLf
        Response.Write "                        <td>" & Country & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>省/市：</td>" & vbCrLf
        Response.Write "                        <td>" & Province & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>市/县/区：</td>" & vbCrLf
        Response.Write "                        <td>" & City & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>邮政编码：</td>" & vbCrLf
        Response.Write "                        <td>" & ZipCode & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>联系电话：</td>" & vbCrLf
        Response.Write "                        <td>" & Phone & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>传真号码：</td>" & vbCrLf
        Response.Write "                        <td>" & Fax2 & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>开户银行：</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & BankOfDeposit & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>银行帐号：</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & BankAccount & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >税号：</td>" & vbCrLf
        Response.Write "                        <td>" & TaxNum & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >网址：</td>" & vbCrLf
        Response.Write "                        <td>" & HomePage2 & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >行业地位：</td>" & vbCrLf
        Response.Write "                        <td>" & GetArrItem(arrStatusInField, StatusInField) & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司规模：</td>" & vbCrLf
        Response.Write "                        <td>" & GetArrItem(arrCompanySize, CompanySize) & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >业务范围：</td>" & vbCrLf
        Response.Write "                        <td>" & BusinessScope & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >年销售额：</td>" & vbCrLf
        Response.Write "                        <td>" & AnnualSales & " 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >经营状态：</td>" & vbCrLf
        Response.Write "                        <td>" & GetArrItem(arrManagementForms, ManagementForms) & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >注册资本：</td>" & vbCrLf
        Response.Write "                        <td>" & RegisteredCapital & " 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司照片：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & CompamyPic & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司简介：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & CompanyIntro & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                </tbody>" & vbCrLf
    End If
    Response.Write "  </table>" & vbCrLf
    
    If rsUser("UserType") > 0 And rsUser("UserType") < 4 Then
        arrUserType = Array("个人会员", "创建者", "管理员", "普通成员", "待审核成员")
        
        Response.Write "<table id='Tabs' style='display:none' width='100%'><tr class='title' align='center'><td>会员名</td><td>真实姓名</td><td>邮政编码</td><td>联系地址</td><td>状态</td><td>操作</td></tr>"
        Dim rsMember
        Set rsMember = Conn.Execute("select U.UserID,U.UserName,U.UserType,C.TrueName,C.ZipCode,C.Address from PE_User U left join PE_Contacter C on U.ContacterID=C.ContacterID where U.CompanyID=" & rsUser("CompanyID") & " order by U.UserType")
        Do While Not rsMember.EOF
            Response.Write "<tr><td align='center'><a href='Admin_User.asp?Action=Show&UserID=" & rsMember("UserID") & "' target='MemberInfo'>" & rsMember("UserName") & "</a></td>"
            Response.Write "<td align='center'><a href='Admin_User.asp?Action=Show&UserID=" & rsMember("UserID") & "' target='MemberInfo'>" & rsMember("TrueName") & "</a></td>"
            Response.Write "<td align='center'>" & rsMember("ZipCode") & "</td>"
            Response.Write "<td>" & rsMember("Address") & "</td>"
            Response.Write "<td align='center'>"
            If PE_CLng(rsMember("UserType")) > 4 Then
                Response.Write arrUserType(0)
            Else
                Response.Write arrUserType(PE_CLng(rsMember("UserType")))
            End If
            Response.Write "</td>"
            Response.Write "<td align='center'>"
            Select Case rsMember("UserType")
            Case 4
                Response.Write "<a href='Admin_User.asp?Action=Agree&UserID=" & rsUser("UserID") & "&MemberID=" & rsMember("UserID") & "'>批准加入</a> "
                Response.Write "<a href='Admin_User.asp?Action=Reject&UserID=" & rsUser("UserID") & "&MemberID=" & rsMember("UserID") & "'>拒绝加入</a>"
            Case 3
                Response.Write "<a href='Admin_User.asp?Action=RemoveFromCompany&UserID=" & rsUser("UserID") & "&MemberID=" & rsMember("UserID") & "'>从企业中删除</a> "
                Response.Write "<a href='Admin_User.asp?Action=AddToAdmin&UserID=" & rsUser("UserID") & "&MemberID=" & rsMember("UserID") & "'>升级为管理员</a>"
            Case 2
                Response.Write "<a href='Admin_User.asp?Action=RemoveFromCompany&UserID=" & rsUser("UserID") & "&MemberID=" & rsMember("UserID") & "'>从企业中删除</a> "
                Response.Write "<a href='Admin_User.asp?Action=RemoveFromAdmin&UserID=" & rsUser("UserID") & "&MemberID=" & rsMember("UserID") & "'>降为普通成员</a>"
            End Select
            Response.Write "</td>"
            Response.Write "</tr>"
            rsMember.MoveNext
        Loop
        rsMember.Close
        Set rsMember = Nothing
        Response.Write "</table>"
    End If
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "<table width='100%' height='60'><tr align='center'><td>"
    If AdminPurview = 1 Or arrPurview(1) = True Then
        Response.Write "    <input type='button' name='Submit' value='修改会员信息' onClick=""window.location.href='Admin_User.asp?Action=Modify&UserID=" & UserID & "'""> "
    End If
    If AdminPurview = 1 Or arrPurview(2) = True Then
        Response.Write "    <input type='button' name='Submit' value='修改会员权限' onClick=""window.location.href='Admin_User.asp?Action=ModifyPurview&UserID=" & UserID & "'""> "
    End If
    If AdminPurview = 1 Or arrPurview(3) = True Then
        If rsUser("IsLocked") = True Then
            Response.Write "    <input type='button' name='Submit' value='  解除锁定  ' onClick=""if(confirm('确定要将此会员解锁吗？')){window.location.href='Admin_User.asp?Action=UnLock&UserID=" & UserID & "';}""> "
        Else
            Response.Write "    <input type='button' name='Submit' value=' 锁定此会员 ' onClick=""if(confirm('确定要锁定此会员吗？')){window.location.href='Admin_User.asp?Action=Lock&UserID=" & UserID & "';}""> "
        End If
    End If
    If AdminPurview = 1 Or arrPurview(4) = True Then
        Response.Write "    <input type='button' name='Submit' value=' 删除此会员 ' onClick=""window.location.href='Admin_User.asp?Action=BatchDel&UserID=" & UserID & "'""> "
    End If
    Response.Write "    <input type='button' name='Submit' value=' 发送短消息 ' onClick=""window.location.href='Admin_Message.asp?Action=Send&UserType=2&UserName=" & rsUser("UserName") & "'""> "
    If AdminPurview = 1 Or arrPurview(6) = True Then
        Response.Write "    <input type='button' name='Submit' value='添加银行汇款' onClick=""window.location.href='Admin_User.asp?Action=AddRemit&UserID=" & UserID & "'""> "
        Response.Write "    <input type='button' name='Submit' value='添加其他收入' onClick=""window.location.href='Admin_User.asp?Action=AddIncome&UserID=" & UserID & "'""> "
        Response.Write "    <input type='button' name='Submit' value='添加支出信息' onClick=""window.location.href='Admin_User.asp?Action=AddPayment&UserID=" & UserID & "'""> "
    End If
    If AdminPurview = 1 Or arrPurview(7) = True Then
        Response.Write "    <input type='button' name='Submit' value='  " & PointName & "兑换  ' onClick=""window.location.href='Admin_User.asp?Action=ExchangePoint&UserID=" & UserID & "'""> "
        Response.Write "    <input type='button' name='Submit' value='  奖励" & PointName & "  ' onClick=""window.location.href='Admin_User.asp?Action=AddPoint&UserID=" & UserID & "'""> "
        Response.Write "    <input type='button' name='Submit' value='  扣除" & PointName & "  '"
        If rsUser("UserPoint") <= 0 Then Response.Write " disabled"
        Response.Write " onClick=""window.location.href='Admin_User.asp?Action=MinusPoint&UserID=" & UserID & "'""> "
    End If
    If AdminPurview = 1 Or arrPurview(8) = True Then
        Response.Write "    <input type='button' name='Submit' value=' 兑换有效期 '"
        If rsUser("ValidNum") = -1 Then Response.Write " disabled"
        Response.Write " onClick=""window.location.href='Admin_User.asp?Action=ExchangeValid&UserID=" & UserID & "'""> "
        Response.Write "    <input type='button' name='Submit' value=' 奖励有效期 '"
        If rsUser("ValidNum") = -1 Then Response.Write " disabled"
        Response.Write " onClick=""window.location.href='Admin_User.asp?Action=AddValid&UserID=" & UserID & "'""> "
        Response.Write "    <input type='button' name='Submit' value=' 扣除有效期 '"
        If ValidDays <= 0 Then Response.Write " disabled"
        Response.Write " onClick=""window.location.href='Admin_User.asp?Action=MinusValid&UserID=" & UserID & "'"">"
    End If
    Response.Write "</td></tr></table>"

    Dim InfoType
    InfoType = PE_CLng(Trim(Request("InfoType")))

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'><td"
    If InfoType = 0 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=0'"""
    End If
    Response.Write ">会员订单</td><td"
    If InfoType = 1 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=1'"""
    End If
    Response.Write ">资金明细</td><td"
    If InfoType = 2 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=2'"""
    End If
    Response.Write ">" & PointName & "明细</td><td"
    If InfoType = 3 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=3'"""
    End If
    Response.Write ">有效期明细</td><td"
    If InfoType = 4 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=4'"""
    End If
    Response.Write ">在线支付</td><td"
    If InfoType = 5 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=5'"""
    End If
    Response.Write ">相关留言</td>"
    If iGroupType = 4 Then
        Response.Write "<td"
        If InfoType = 6 Then
            Response.Write " class='title6'"
        Else
            Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=6'"""
        End If
        Response.Write ">代理订单</td>"
        Response.Write "<td"
        If InfoType = 7 Then
            Response.Write " class='title6'"
        Else
            Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=7'"""
        End If
        Response.Write ">对账单</td>"
        Response.Write "<td"
        If InfoType = 8 Then
            Response.Write " class='title6'"
        Else
            Response.Write " class='title5' onclick=""window.location.href='Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=8'"""
        End If
        Response.Write ">被投诉记录</td>"
    End If
    Response.Write "<td>&nbsp;</td></tr></table>"
    
    strFileName = "Admin_User.asp?Action=Show&UserID=" & UserID & "&InfoType=" & InfoType
    
    Select Case InfoType
    Case 0
        Call ShowOrderList(0, rsUser("UserName"))
    Case 1
        Call ShowBankrollList(rsUser("UserName"))
    Case 2
        Call ShowConsumeLog(rsUser("UserName"))
    Case 3
        Call ShowRechargeLog(rsUser("UserName"))
    Case 4
        Call ShowPayOnline(rsUser("UserName"))
    Case 5
        Call ShowGuestBook(rsUser("UserName"))
    Case 6
        Call ShowOrderList(1, rsUser("UserName"))
    Case 7
        Call ShowMyBill(rsUser("UserName"))
    Case 8
        Call ShowComplain(rsUser("UserName"))
    End Select

    rsUser.Close
    Set rsUser = Nothing
    Response.Write "<br><br>"
End Sub

Sub ShowComplain(UserName)
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>投诉时间</td>" & vbCrLf
    Response.Write "    <td width='60' align='center'>客户名称</td>" & vbCrLf
    Response.Write "    <td width='80' align='center'>投诉类型</td>" & vbCrLf
    Response.Write "    <td>主题</td>" & vbCrLf
    Response.Write "    <td width='60' align='center'>紧急程度</td>" & vbCrLf
    Response.Write "    <td width='60' align='center'>记录状态</td>" & vbCrLf
'    Response.Write "    <td width='60' align='center'>处理人</td>" & vbCrLf
'    Response.Write "    <td width='60' align='center'>处理时间</td>" & vbCrLf
'    Response.Write "    <td width='60' align='center'>处理结果</td>" & vbCrLf
'    Response.Write "    <td width='60' align='center'>回访人</td>" & vbCrLf
'    Response.Write "    <td width='60' align='center'>回访时间</td>" & vbCrLf
'    Response.Write "    <td width='60' align='center'>客户评分</td>" & vbCrLf
    Response.Write "  </tr>"
    
    Dim rsComplain, sqlComplain, TotalIncome, TotalPayout, arrComplainType, arrMagnitudeOfExigence, arrStatus
    arrComplainType = GetArrFromDictionary("PE_ComplainItem", "ComplainType")
    arrMagnitudeOfExigence = GetArrFromDictionary("PE_ComplainItem", "MagnitudeOfExigence")
    arrStatus = Array("未处理", "处理中", "已处理", "已回访")
    sqlComplain = "select S.*,C.ShortedForm from PE_ComplainItem S inner join PE_Client C on S.ClientID=C.ClientID where S.Defendant='" & UserName & "' order by S.ItemID desc"
    Set rsComplain = Server.CreateObject("Adodb.RecordSet")
    rsComplain.Open sqlComplain, Conn, 1, 1
    If rsComplain.BOF And rsComplain.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何被投诉记录！</td></tr>"
    Else
        totalPut = rsComplain.RecordCount
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
                rsComplain.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim i
        i = 0
        Do While Not rsComplain.EOF

            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td width='120'>" & rsComplain("DateAndTime") & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'><a href='Admin_Client.asp?Action=Show&InfoType=6&ClientID=" & rsComplain("ClientID") & "'>" & rsComplain("ShortedForm") & "</a></td>" & vbCrLf
            Response.Write "    <td width='80' align='center'>" & GetArrItem(arrComplainType, rsComplain("ComplainType")) & "</td>" & vbCrLf
            Response.Write "    <td><a href='Admin_Complain.asp?Action=Show&ItemID=" & rsComplain("ItemID") & "'>" & rsComplain("Title") & "</a></td>" & vbCrLf
            Response.Write "    <td width='60' align='center'>" & GetArrItem(arrMagnitudeOfExigence, rsComplain("MagnitudeOfExigence")) & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'>" & GetArrItem(arrStatus, rsComplain("Status")) & "</td>" & vbCrLf
'            Response.Write "    <td width='60' align='center'>" & rsComplain("Processor") & "</td>" & vbCrLf
'            Response.Write "    <td width='60' align='center'>" & rsComplain("EndTime") & "</td>" & vbCrLf
'            Response.Write "    <td width='60' align='center'>" & rsComplain("Result") & "</td>" & vbCrLf
'            Response.Write "    <td width='60' align='center'>" & rsComplain("ConfirmCaller") & "</td>" & vbCrLf
'            Response.Write "    <td width='60' align='center'>" & rsComplain("EndTime") & "</td>" & vbCrLf
'            Response.Write "    <td width='60' align='center'>" & String(PE_Clng(rsComplain("ConfirmStar")),"★") & "</td>" & vbCrLf
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsComplain.MoveNext
        Loop
    End If
    rsComplain.Close
    Set rsComplain = Nothing

    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条投诉记录", True)

End Sub

Sub ShowGuestBook(UName)
    Dim sqlGuest, rsGuest
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_GuestBook.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "  <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "     <tr class='title'>"
    Response.Write "    <td width='30' height='22' align='center'>属性</td>"
    Response.Write "    <td width='30' height='22' align='center'>选中</td>"
    Response.Write "    <td width='85' height='22' align='center'>留言人</td>"
    Response.Write "    <td height='22' align='center'>留言主题</td>"
    Response.Write "    <td width='120' height='22' align='center'>留言时间</td>"
    Response.Write "    <td width='30' height='22' align='center'>审核</td>"
    Response.Write "    <td width='228' height='22' align='center'>操作</td>"
    Response.Write "  </tr>"

    sqlGuest = " select G.*,K.KindName from PE_GuestBook G left join PE_GuestKind K on G.KindID=K.KindID where GuestName='" & UName & "'"
    If Keyword <> "" Then
        Select Case strField
        Case "GuestTitle"
            sqlGuest = sqlGuest & " and GuestTitle like '%" & Keyword & "%' "
        Case "GuestContent"
            sqlGuest = sqlGuest & " and GuestContent like '%" & Keyword & "%' "
        Case "GuestReply"
            sqlGuest = sqlGuest & " and GuestReply like '%" & Keyword & "%' "
        Case "GuestName"
            sqlGuest = sqlGuest & " and GuestName like '%" & Keyword & "%' "
        Case Else
            sqlGuest = sqlGuest & " and GuestTitle like '%" & Keyword & "%' "
        End Select
    End If
    sqlGuest = sqlGuest & " order by G.TopicID desc,G.GuestId asc"
    Set rsGuest = Server.CreateObject("adodb.recordset")
    rsGuest.Open sqlGuest, Conn, 1, 1
    If rsGuest.BOF And rsGuest.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何留言！<br><br></td></tr>"
    Else
        totalPut = rsGuest.RecordCount
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
                rsGuest.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim GuestNum
        GuestNum = 0

        Do While Not rsGuest.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If rsGuest("TopicID") = rsGuest("GuestID") Then
                Response.Write "      <td width='30' align='center'>主题</td>"
            Else
                Response.Write "      <td width='30' align='center' bgcolor='#ffffff'></td>"
            End If
            Response.Write "      <td width='30' align='center'><input name='GuestID' type='checkbox' onclick='unselectall()' value='" & rsGuest("GuestID") & "'></td>"
            Response.Write "      <td width='85' align='center'><div style='cursor:hand' "
            If rsGuest("GuestType") = 1 Then
                Response.Write " title='类型：注册用户" & vbCrLf & "  IP：" & rsGuest("GuestIP") & "'"
            Else
                Response.Write " title='类型：游客" & vbCrLf
                Response.Write "性别："
                If rsGuest("GuestSex") = "0" Then
                    Response.Write "女"
                ElseIf rsGuest("GuestSex") = "1" Then
                    Response.Write "男"
                Else
                    Response.Write "保密"
                End If
                Response.Write vbCrLf & "信箱：" & rsGuest("GuestEmail") & vbCrLf & "OICQ：" & rsGuest("GuestOicq") & vbCrLf & " ICQ：" & rsGuest("GuestIcq") & vbCrLf & " MSN：" & rsGuest("GuestMsn") & vbCrLf & "主页：" & rsGuest("GuestHomepage") & vbCrLf & "  IP：" & rsGuest("GuestIP") & "'"
            End If

            Response.Write " >" & rsGuest("GuestName") & "</div></td>"
            Response.Write "      <td><a href='Admin_GuestBook.asp?Action=Show&GuestID=" & rsGuest("GuestID") & "'"
            Response.Write "      title='留言内容：" & nohtml(rsGuest("GuestContent"))
            If rsGuest("GuestReply") <> "" Then
                Response.Write vbCrLf & rsGuest("GuestReplyAdmin") & "回复：" & nohtml(rsGuest("GuestReply"))
            End If
            Response.Write "'>"
            If rsGuest("GuestIsPrivate") = True Then
                Response.Write "<font color=green>【隐藏】</font>" & vbCrLf
            End If
            Dim Title
            Title = rsGuest("GuestTitle")
            If Len(Title) > 18 Then
                Title = Left(Title, 18) & "..."
            End If
            If rsGuest("KindName") <> "" Then
                Response.Write "[" & rsGuest("KindName") & "]" & Title & "</a></td>"
            Else
                Response.Write "[不指定类别]" & Title & "</a></td>"
            End If
            Response.Write "      <td width='120' align='center'>"
            If rsGuest("GuestDatetime") <> "" Then
                Response.Write FormatDateTime(rsGuest("GuestDatetime"), 0)
            End If
            Response.Write "</td>"
            Response.Write "      <td width='30' align='center'>"
            If rsGuest("GuestIsPassed") = True Then
                Response.Write "√"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='228' align='center'>"
            Response.Write "      <a href='Admin_GuestBook.asp?Action=Modify&GuestID=" & rsGuest("GuestID") & "'>修改</a>"
            If rsGuest("TopicID") <> rsGuest("GuestID") Then
                Response.Write "      <a href='Admin_GuestBook.asp?Action=Del&GuestID=" & rsGuest("GuestID") & "' onClick=""return confirm('确定要删除此回复吗？');"">删除</a>"
            Else
                Response.Write "      <a href='Admin_GuestBook.asp?Action=Del&GuestID=" & rsGuest("GuestID") & "' onClick=""return confirm('删除此主题将删除其下所有回复，确定要删除此主题吗？');"">删除</a>"
            End If
            Response.Write "      <a href='Admin_GuestBook.asp?Action=AdminReply&GuestID=" & rsGuest("GuestID") & "'>回复</a>"
            If rsGuest("GuestReply") <> "" Then
                Response.Write "      <a href='Admin_GuestBook.asp?Action=DelReply&GuestID=" & rsGuest("GuestID") & "'>清除回复</a>"
            End If
            If rsGuest("GuestIsPassed") = False Then
                Response.Write "      <a href='Admin_GuestBook.asp?Action=SetPassed&GuestID=" & rsGuest("GuestID") & "'>通过审核</a>"
            Else
                Response.Write "      <a href='Admin_GuestBook.asp?Action=CancelPassed&GuestID=" & rsGuest("GuestID") & "'>取消审核</a>"
            End If
            If rsGuest("TopicID") = rsGuest("GuestID") Then
                If rsGuest("Quintessence") = 0 Then
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=Quintessence&GuestID=" & rsGuest("GuestID") & "'>推荐精华</a>"
                Else
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=Cquintessence&GuestID=" & rsGuest("GuestID") & "'>取消精华</a>"
                End If
                If rsGuest("OnTop") = 0 Then
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=SetOnTop&GuestID=" & rsGuest("GuestID") & "'>固顶</a>"
                Else
                    Response.Write "      <a href='Admin_GuestBook.asp?Action=CancelOnTop&GuestID=" & rsGuest("GuestID") & "'>解固</a>"
                End If
            End If
            Response.Write "      </td>"
            Response.Write "    </tr>"

            GuestNum = GuestNum + 1
            If GuestNum >= MaxPerPage Then Exit Do
            rsGuest.MoveNext
        Loop
    End If
    rsGuest.Close
    Set rsGuest = Nothing
    Response.Write "</table>"
End Sub
      

Sub ShowOrderList(UserType, UserName)
    Response.Write "<table width='100%'  border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='110'>订单编号</td>"
    Response.Write "    <td>客户名称</td>"
    Response.Write "    <td>用户名</td>"
    Response.Write "    <td width='120'>下单时间</td>"
    Response.Write "    <td width='70'>订单金额</td>"
    Response.Write "    <td width='70'>收款金额</td>"
    Response.Write "    <td width='30'>需要<br>发票</td>"
    Response.Write "    <td width='30'>已开<br>发票</td>"
    Response.Write "    <td width='50'>订单状态</td>"
    Response.Write "    <td width='50'>付款状态</td>"
    Response.Write "    <td width='50'>物流状态</td>"
    Response.Write "  </tr>"

    Dim rsOrderList, sqlOrderList, dblMoneyTotal1, dblMoneyTotal2
    dblMoneyTotal1 = 0
    dblMoneyTotal2 = 0
    sqlOrderList = "select  O.OrderFormID,O.OrderFormNum,O.InputTime,O.UserName,O.ClientID,C.ShortedForm as ClientName,O.MoneyTotal,O.MoneyReceipt,O.NeedInvoice,O.Invoiced,O.Status,O.DeliverStatus from PE_OrderForm O left join PE_Client C on O.ClientID=C.ClientID "
    If UserType = 1 Then
        sqlOrderList = sqlOrderList & " where O.AgentName='" & UserName & "' order by OrderFormID desc"
    Else
        sqlOrderList = sqlOrderList & " where O.UserName='" & UserName & "' order by OrderFormID desc"
    End If
    Set rsOrderList = Server.CreateObject("adodb.recordset")
    rsOrderList.Open sqlOrderList, Conn, 1, 1
    If rsOrderList.BOF And rsOrderList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何订单！</td></tr>"
    Else
        totalPut = rsOrderList.RecordCount
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
                rsOrderList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim OrderNum
        Do While Not rsOrderList.EOF
            dblMoneyTotal1 = dblMoneyTotal1 + rsOrderList("MoneyTotal")
            dblMoneyTotal2 = dblMoneyTotal2 + rsOrderList("MoneyReceipt")
    
            Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='110'><a href='Admin_Order.asp?Action=ShowOrder&OrderFormID=" & rsOrderList("OrderFormID") & "'>" & rsOrderList("OrderFormNum") & "</a></td>"
            Response.Write "    <td><a href='Admin_Client.asp?Action=Show&InfoType=2&ClientID=" & rsOrderList("ClientID") & "'>" & rsOrderList("ClientName") & "</a></td>"
            Response.Write "    <td><a href='Admin_User.asp?Action=Show&UserName=" & rsOrderList("UserName") & "'>" & rsOrderList("UserName") & "</a></td>"
            Response.Write "    <td width='120'>" & rsOrderList("InputTime") & "</td>"
            Response.Write "    <td width='70' align='right'>" & FormatNumber(rsOrderList("MoneyTotal"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='70' align='right'>"
            If rsOrderList("MoneyReceipt") < rsOrderList("MoneyTotal") Then
                Response.Write "<font color='red'>" & FormatNumber(rsOrderList("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue) & "</font>"
            Else
                Response.Write FormatNumber(rsOrderList("MoneyReceipt"), 2, vbTrue, vbFalse, vbTrue)
            End If
            Response.Write "</td>"
            Response.Write "    <td width='30'>"
            If rsOrderList("NeedInvoice") = True Then
                Response.Write "<font color='red'>√</font>"
            End If
            Response.Write "</td>"
            Response.Write "    <td width='30'>"
            If rsOrderList("NeedInvoice") = True Then
                If rsOrderList("Invoiced") = True Then
                    Response.Write "√"
                Else
                    Response.Write "<font color='red'>×</font>"
                End If
            End If
            Response.Write "</td>"
            Response.Write "           <td width='50'>"
            Select Case rsOrderList("Status")
                Case 0, 1
                    Response.Write "<font color='red'>等待确认</font>"
                Case 2, 3
                    Response.Write "<font color='blue'>已经确认</font>"
                Case 4
                    Response.Write "<font color='gray'>已结清</font>"
            End Select
            Response.Write "</td>"
            Response.Write "           <td width='50'>"
            If rsOrderList("MoneyTotal") > rsOrderList("MoneyReceipt") Then
                If rsOrderList("MoneyReceipt") > 0 Then
                    Response.Write "<font color='green'>已收定金</font>"
                Else
                    Response.Write "<font color='red'>等待汇款</font>"
                End If
            Else
                Response.Write "<font color='blue'>已经付清</font>"
            End If
            Response.Write "</td>"
            Response.Write "           <td width='50'>"
            Select Case rsOrderList("DeliverStatus")
                Case 0, 1
                    Response.Write "<font color='red'>配送中</font>"
                Case 2
                    Response.Write "<font color='blue'>已发货</font>"
                Case 3
                    Response.Write "<font color='green'>已签收</font>"
            End Select
            Response.Write "</td>"
            Response.Write "  </tr>"
    
            OrderNum = OrderNum + 1
            If OrderNum >= MaxPerPage Then Exit Do
            rsOrderList.MoveNext
        Loop
    End If
    rsOrderList.Close
    Set rsOrderList = Nothing

    Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='4' align='right'>合计：</td>"
    Response.Write "    <td width='70' align='right'>" & FormatNumber(dblMoneyTotal1, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td width='70' align='right'>" & FormatNumber(dblMoneyTotal2, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='5'> </td>"
    Response.Write "  </tr>"
    
    Dim dblCount, rsCount, dblCount2
    If UserType = 1 Then
        Set rsCount = Conn.Execute("select sum(MoneyTotal) from PE_OrderForm where AgentName='" & UserName & "'")
    Else
        Set rsCount = Conn.Execute("select sum(MoneyTotal) from PE_OrderForm where UserName='" & UserName & "'")
    End If
    If IsNull(rsCount(0)) Then
        dblCount = 0
    Else
        dblCount = rsCount(0)
    End If
    Set rsCount = Nothing
    If UserType = 1 Then
        Set rsCount = Conn.Execute("select sum(MoneyReceipt) from PE_OrderForm where AgentName='" & UserName & "'")
    Else
        Set rsCount = Conn.Execute("select sum(MoneyReceipt) from PE_OrderForm where UserName='" & UserName & "'")
    End If
    If IsNull(rsCount(0)) Then
        dblCount2 = 0
    Else
        dblCount2 = rsCount(0)
    End If
    Set rsCount = Nothing
    
    Response.Write "         <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "           <td colspan='4' align='right'>总计金额：</td>"
    Response.Write "           <td align='right'>" & FormatNumber(dblCount, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "           <td align='right'>" & FormatNumber(dblCount2, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "           <td colspan='5'> </td>"
    Response.Write "         </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个订单", True)
End Sub

Sub ShowMyBill(UserName)
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='200'>交易时间</td>"
    Response.Write "    <td width='110'>订单号</td>"
    'Response.Write "    <td width='50'>币种</td>"
    Response.Write "    <td width='80'>收入金额</td>"
    Response.Write "    <td width='80'>支出金额</td>"
    'Response.Write "    <td width='40'>摘要</td>"
    'Response.Write "    <td width='60'>银行名称</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "  </tr>"
    
    Dim rsBankroll, sqlBankroll, TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    sqlBankroll = "select Max(O.OrderFormNum) as tOrderFormNum,sum(B.Money) as tMoney,Max(B.DateAndTime) as tDateAndTime,Max(B.Remark) as tRemark From PE_BankrollItem B Left join PE_OrderForm O On B.OrderFormID=O.OrderFormID where B.UserName='" & UserName & "' Group By B.OrderFormID Order by Max(B.DateAndTime) desc"
    'sqlBankroll = "select Max(O.OrderFormNum) as tOrderFormNum,sum(B.Money) as tMoney,Max(B.DateAndTime) as tDateAndTime,Max(B.Remark) as tRemark From PE_BankrollItem B Left join PE_OrderForm O On B.OrderFormID=O.OrderFormID Group By B.OrderFormID"
    Set rsBankroll = Server.CreateObject("Adodb.RecordSet")
    rsBankroll.Open sqlBankroll, Conn, 1, 1
    If rsBankroll.BOF And rsBankroll.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何资金明细记录！</td></tr>"
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
            If rsBankroll("tMoney") > 0 Then
                TotalIncome = TotalIncome + rsBankroll("tMoney")
            Else
                TotalPayout = TotalPayout + rsBankroll("tMoney")
            End If
    
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='200' align='center'>" & rsBankroll("tDateAndTime") & "</td>"
            Response.Write "    <td width='110' align='center'>"
            Response.Write rsBankroll("tOrderFormNum")
            Response.Write "</td>"
            Response.Write "    <td width='80' align='right'>"
            If rsBankroll("tMoney") > 0 Then Response.Write FormatNumber(rsBankroll("tMoney"), 2, vbTrue, vbFalse, vbTrue)
            Response.Write "</td>"
            Response.Write "    <td width='80' align='right'>"
            If rsBankroll("tMoney") <= 0 Then Response.Write FormatNumber(Abs(rsBankroll("tMoney")), 2, vbTrue, vbFalse, vbTrue)
            Response.Write "</td>"
            Response.Write "    <td align='center'>" & rsBankroll("tRemark") & "</td>"
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsBankroll.MoveNext
        Loop
    End If
    rsBankroll.Close
    Set rsBankroll = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='2' align='right'>本页合计：</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncome, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayout), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='center'> </td>"
    
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
    'Response.Write "    <td colspan='2' align='right'>总计金额：</td>"
    'Response.Write "    <td align='right'>" & FormatNumber(TotalIncomeAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    'Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayoutAll), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='5' align='right'><strong>资金余额：" & FormatNumber(TotalIncomeAll + TotalPayoutAll, 2, vbTrue, vbFalse, vbTrue) & "</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条对账记录", True)
End Sub
Sub ShowBankrollList(UserName)
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>交易时间</td>"
    Response.Write "    <td width='60'>交易方式</td>"
    Response.Write "    <td width='50'>币种</td>"
    Response.Write "    <td width='80'>收入金额</td>"
    Response.Write "    <td width='80'>支出金额</td>"
    Response.Write "    <td width='60'>银行名称</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "  </tr>"
    
    Dim rsBankroll, sqlBankroll, TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    sqlBankroll = "select * from PE_BankrollItem where UserName='" & UserName & "' order by ItemID desc"
    Set rsBankroll = Server.CreateObject("Adodb.RecordSet")
    rsBankroll.Open sqlBankroll, Conn, 1, 1
    If rsBankroll.BOF And rsBankroll.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何资金明细记录！</td></tr>"
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
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsBankroll.MoveNext
        Loop
    End If
    rsBankroll.Close
    Set rsBankroll = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='3' align='right'>本页合计：</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncome, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayout), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='3' align='center'> </td>"
    
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
    Response.Write "    <td colspan='3' align='right'>总计金额：</td>"
    Response.Write "    <td align='right'>" & FormatNumber(TotalIncomeAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td align='right'>" & FormatNumber(Abs(TotalPayoutAll), 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='3' align='center'>资金余额：" & FormatNumber(TotalIncomeAll + TotalPayoutAll, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条资金明细记录", True)
End Sub

Sub ShowConsumeLog(UserName)
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>消费时间</td>"
    Response.Write "    <td width='120'>IP地址</td>"
    Response.Write "    <td width='50'>收入" & PointName & "数</td>"
    Response.Write "    <td width='50'>支出" & PointName & "数</td>"
    Response.Write "    <td width='40'>摘要</td>"
    Response.Write "    <td width='60'>重复次数</td>"
    Response.Write "    <td width='60'>操作员</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "  </tr>"
    
    Dim rsConsumeLog, sqlConsumeLog
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    
    sqlConsumeLog = "select * from PE_ConsumeLog where UserName='" & UserName & "' order by LogID desc"
    Set rsConsumeLog = Server.CreateObject("Adodb.RecordSet")
    rsConsumeLog.Open sqlConsumeLog, Conn, 1, 1
    If rsConsumeLog.BOF And rsConsumeLog.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何消费记录！</td></tr>"
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
            Response.Write "    <td width='120' align='center'>" & rsConsumeLog("IP") & "</td>"
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
    Response.Write "    <td colspan='2' align='right'>本页合计：</td>"
    Response.Write "    <td align='right'>" & TotalIncome & "</td>"
    Response.Write "    <td align='right'>" & TotalPayout & "</td>"
    Response.Write "    <td colspan='4'>&nbsp;</td>"
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
    Response.Write "    <td colspan='2' align='right'>总计" & PointName & "数：</td>"
    Response.Write "    <td align='right'>" & TotalIncomeAll & "</td>"
    Response.Write "    <td align='right'>" & TotalPayoutAll & "</td>"
    Response.Write "    <td colspan='4' align='center'>" & PointName & "数余额：" & TotalIncomeAll - TotalPayoutAll & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条" & PointName & "明细记录", True)
End Sub

Sub ShowRechargeLog(UserName)
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>时间</td>"
    Response.Write "    <td width='120'>IP地址</td>"
    Response.Write "    <td width='60'>增加有效期</td>"
    Response.Write "    <td width='60'>减去有效期</td>"
    Response.Write "    <td width='40'>摘要</td>"
    Response.Write "    <td width='60'>操作员</td>"
    Response.Write "    <td>备注/说明</td>"
    Response.Write "  </tr>"
    
    Dim rsRechargeLog, sqlRechargeLog
    
    sqlRechargeLog = "select * from PE_RechargeLog where UserName='" & UserName & "' order by LogID desc"
    Set rsRechargeLog = Server.CreateObject("Adodb.RecordSet")
    rsRechargeLog.Open sqlRechargeLog, Conn, 1, 1
    If rsRechargeLog.BOF And rsRechargeLog.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何有效期更改记录！</td></tr>"
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
            Response.Write "    <td width='120' align='center'>" & rsRechargeLog("IP") & "</td>"
            Response.Write "    <td width='60' align='center'>"
            If rsRechargeLog("Income_Payout") = 1 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>"
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
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条有效期更改明细记录", True)
End Sub

Sub ShowPayOnline(UserName)
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='80'>支付序号</td>"
    Response.Write "    <td width='70'>支付平台</td>"
    Response.Write "    <td width='120'>交易时间</td>"
    Response.Write "    <td width='70'>汇款金额</td>"
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
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>没有任何在线支付记录！</td></tr>"
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
            Response.Write "    <td width='70' align='right'>" & FormatNumber(rsPaymentList("MoneyPay"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='80' align='right'>" & FormatNumber(rsPaymentList("MoneyTrue"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "    <td width='60' align='center'>"
            Select Case rsPaymentList("Status")
            Case 1
                Response.Write "未提交"
            Case 2
                Response.Write "已经提交，但未成功"
            Case 3
                Response.Write "支付成功"
            End Select
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
    Response.Write "    <td colspan='3' align='right'>合计金额：</td>"
    Response.Write "    <td width='80' align='right'>" & FormatNumber(TotalMoneyPay, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td width='80' align='right'>" & FormatNumber(TotalMoneyTrue, 2, vbTrue, vbFalse, vbTrue) & "</td>"
    Response.Write "    <td colspan='4' align='center'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条在线支付记录", True)
End Sub

Sub BatchMove()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    
    Dim toUserGroupID, BatchUserGroupID, UserID, BatchUserName, uUserID, lUserID
    UserID = Trim(Request("UserID"))
      
    Response.Write "<form method='POST' name='myform' action='Admin_User.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='4' align='center'>批量移动会员</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td vlign='top' width='300'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='UserType' value='1' checked>指定会员ID：<input type='text' name='BatchUserID' value='" & UserID & "' size='30'><br>"
    Response.Write "              <input type='radio' name='UserType' value='2'>指定用户名：<input type='text' name='BatchUserName' size='30'><br>"
    Response.Write "              <input type='radio' name='UserType' value='3'>指定会员ID的范围：<input type='text' name='uUserID' size='6'> 至 <input type='text' name='lUserID' size='6'><br>"
    Response.Write "              <input type='radio' name='UserType' value='4'>指定要移动的会员组：<br><select name='BatchUserGroupID' size='2' multiple style='height:360px;width:300px;'>" & GetUserGroup_Option(0) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td>移动到&gt;&gt;</td>"
    Response.Write "      <td valign='bottom'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td><br><br><br>"
    Response.Write "              请指定目标会员组：<br><select name='toUserGroupID' size='2' style='height:360px;width:300px;'>" & GetUserGroup_Option(0) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='BatchMove'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 执行批处理 ' style='cursor:hand;' onClick=""document.myform.Action.value='DoBatchMove';"">&nbsp; "
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_User.asp';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub DoBatchMove()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim toUserGroupID, BatchUserGroupID, BatchUserID, UserType, BatchUserName, uUserID, lUserID, sqlstr
    UserType = PE_CLng(Trim(Request("UserType")))
    BatchUserID = Trim(Request.Form("BatchUserID"))
    BatchUserGroupID = Trim(Request.Form("BatchUserGroupID"))
    toUserGroupID = PE_CLng(Request("toUserGroupID"))
    BatchUserName = Trim(Request.Form("BatchUserName"))
    uUserID = PE_CLng(Trim(Request.Form("uUserID")))
    lUserID = PE_CLng(Trim(Request.Form("lUserID")))
         
    Select Case UserType
    Case 1
        If IsValidID(BatchUserID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定批量移动用户的ID</li>"
        Else
            BatchUserID = ReplaceBadChar(BatchUserID)
        End If
    Case 2
        If BatchUserName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定批量移动的用户名</li>"
        Else
            BatchUserName = Replace(ReplaceBadChar(BatchUserName), ",", "','")
        End If
    Case 3
        If uUserID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定批量移动用户的ID的上限</li>"
        End If
        If lUserID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定批量移动用户的ID的下限</li>"
        End If
    Case 4
        If IsValidID(BatchUserGroupID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定批量移动用户组</li>"
        End If
    End Select
        
    If toUserGroupID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定目标用户组</li>"
    End If
    If FoundErr = True Then Exit Sub
    
    sqlstr = "Update PE_User set GroupID= " & toUserGroupID & ""
    Select Case UserType
    Case 1
        sqlstr = sqlstr & " where UserID in (" & BatchUserID & ")"
        Conn.Execute sqlstr
    Case 2
        sqlstr = sqlstr & " where UserName in ('" & BatchUserName & "')"
        Conn.Execute sqlstr
    Case 3
        sqlstr = sqlstr & " where UserId between " & uUserID & " and " & lUserID & ""
        Conn.Execute sqlstr
    Case 4
        sqlstr = sqlstr & " where GroupID in (" & BatchUserGroupID & ")"
        Conn.Execute sqlstr
    End Select
    ComeUrl = "Admin_User.asp"
    Call WriteSuccessMsg("成功将选定的用户移动到目标用户组中！", ComeUrl)
    Call ClearSiteCache(0)
End Sub

Sub AddIncome(IncomeType)
    If AdminPurview > 1 And arrPurview(6) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, rsUser
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！"
    End If
    If FoundErr = True Then Exit Sub

    Set rsUser = Conn.Execute("select UserName,Balance from PE_User where UserID=" & UserID & "")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的ID！</li>"
        Set rsUser = Nothing
        Exit Sub
    End If

    Response.Write "<form name='form4' method='post' action='Admin_User.asp' onsubmit=""return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')"">"
    Response.Write "  <table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    If IncomeType = 1 Then
        Response.Write "      <td height='22' colspan='2'>添 加 会 员 汇 款 信 息</td>"
    Else
        Response.Write "      <td height='22' colspan='2'>添 加 会 员 其 他 收 入</td>"
    End If
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>会 员 名：</td>"
    Response.Write "      <td align='left'><input name='UserName' type='text' id='UserName' value='" & rsUser("UserName") & "' size='30' maxlength='50' disabled></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>资金余额：</td>"
    Response.Write "      <td align='left'><input name='Balance' type='text' id='Balance' value='" & rsUser("Balance") & "元' size='30' maxlength='50' disabled></td>"
    Response.Write "    </tr>"
    If IncomeType = 1 Then '银行汇款
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>到款日期：</td>"
        Response.Write "      <td align='left'><input name='ReceiptDate' type='text' id='ReceiptDate' value='" & Date & "' size='15' maxlength='30'></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>汇款金额：</td>"
        Response.Write "      <td align='left'><input name='Money' type='text' id='Money' value='' size='10' maxlength='20'> 元</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>汇入银行：</td>"
        Response.Write "      <td align='left'>"
        Response.Write "<STYLE TYPE=""text/css"" TITLE="""">" & vbCrLf
        Response.Write "DIV { FONT: 12px 宋体 }" & vbCrLf
        Response.Write "LABEL { PADDING-RIGHT: 0px; PADDING-LEFT: 4px; PADDING-BOTTOM: 0px; PADDING-TOP: 3px; HEIGHT: 19px }" & vbCrLf
        Response.Write ".link_box { CURSOR: default; TEXT-ALIGN: left }" & vbCrLf
        Response.Write ".link_head { BORDER-RIGHT: 2px inset; BORDER-TOP: 2px inset; BORDER-LEFT: 2px inset; WIDTH: 100%; BORDER-BOTTOM: 2px inset; HEIGHT: 23px }" & vbCrLf
        Response.Write ".link_text { PADDING-LEFT: 2px; BACKGROUND: #fff }" & vbCrLf
        Response.Write ".link_arrow0 { BORDER-RIGHT: 2px outset; BORDER-TOP: 2px outset; BACKGROUND: buttonface; FONT: 14px marlett; BORDER-LEFT: 2px outset; WIDTH: 18px; BORDER-BOTTOM: 2px outset; HEIGHT: 100%; TEXT-ALIGN: center }" & vbCrLf
        Response.Write ".link_arrow1 { BORDER-RIGHT: buttonshadow 1px solid; PADDING-RIGHT: 0px; BORDER-TOP: buttonshadow 1px solid; PADDING-LEFT: 2px; BACKGROUND: buttonface; PADDING-BOTTOM: 0px; FONT: 14px marlett; BORDER-LEFT: buttonshadow 1px solid; WIDTH: 18px; PADDING-TOP: 2px; BORDER-BOTTOM: buttonshadow 1px solid; HEIGHT: 100%; TEXT-ALIGN: center }" & vbCrLf
        Response.Write ".link_value { BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; FILTER: alpha(opacity:0); VISIBILITY: hidden; OVERFLOW-X: hidden; OVERFLOW: auto; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; POSITION: absolute }" & vbCrLf
        Response.Write ".link_record0 { BORDER-TOP: #eee 1px solid; PADDING-LEFT: 2px; BACKGROUND: #fff; WIDTH: 100%; COLOR: #000; HEIGHT: 20px }" & vbCrLf
        Response.Write ".link_record1 { BORDER-TOP: #047 1px solid; PADDING-LEFT: 2px; BACKGROUND: #058; WIDTH: 100%; COLOR: #fe0; HEIGHT: 20px }" & vbCrLf
        Response.Write "</style>" & vbCrLf
        Response.Write "<script language=""JavaScript"">" & vbCrLf
        Response.Write "<!--" & vbCrLf
        Response.Write "var dropShow=false" & vbCrLf
        Response.Write "var currentID" & vbCrLf
        Response.Write "function dropdown(el){" & vbCrLf
        Response.Write "    if(dropShow){" & vbCrLf
        Response.Write "        dropFadeOut()" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        currentID=el" & vbCrLf
        Response.Write "        el.style.visibility=""visible""" & vbCrLf
        Response.Write "        dropFadeIn()" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "function dropFadeIn(){  //选单淡入的效果" & vbCrLf
        Response.Write "    if(currentID.filters.alpha.opacity<100){" & vbCrLf
        Response.Write "        currentID.filters.alpha.opacity+=20" & vbCrLf
        Response.Write "        fadeTimer=setTimeout(""dropFadeIn()"",36)" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        dropShow=true" & vbCrLf
        Response.Write "        clearTimeout(fadeTimer)" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "function dropFadeOut(){//选单淡出的效果" & vbCrLf
        Response.Write "    if(currentID.filters.alpha.opacity>0){" & vbCrLf
        Response.Write "        clearTimeout(fadeTimer)" & vbCrLf
        Response.Write "        currentID.filters.alpha.opacity-=20" & vbCrLf
        Response.Write "        fadeTimer=setTimeout(""dropFadeOut()"",36)" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        dropShow=false" & vbCrLf
        Response.Write "        currentID.style.visibility=""hidden""" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "function dropdownHide(){" & vbCrLf
        Response.Write "    if(dropShow){" & vbCrLf
        Response.Write "        dropFadeOut()" & vbCrLf
        Response.Write "        dropShow=false" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "function hiLight(el){//高亮度显示指标位置" & vbCrLf
        Response.Write "    if(dropShow){" & vbCrLf
        Response.Write "        for(i=0;i<el.parentElement.childNodes.length;i++){" & vbCrLf
        Response.Write "            el.parentElement.childNodes(i).className=""link_record0""" & vbCrLf
        Response.Write "        }" & vbCrLf
        Response.Write "        el.className=""link_record1""" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "function CheckMe(el){//替换显示内容" & vbCrLf
        Response.Write "    document.all.text1.innerText=el.innerText" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "document.onclick=dropdownHide" & vbCrLf
        Response.Write "-->" & vbCrLf
        Response.Write "</script>" & vbCrLf
        Response.Write "      <div class=""link_box"" onselectstart=""return false"" style=""WIDTH: 100px"">" & vbCrLf
        Response.Write "      <div class=""link_head"" onclick=""dropdown(value1)"">" & vbCrLf
        Response.Write "        <table height=""100%"" cellSpacing=""0"" cellPadding=""0"" width=""100%"" border=""0"">" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td>" & vbCrLf
        Dim rsBank, DefaultBank
        DefaultBank = ""
        Set rsBank = Conn.Execute("select BankShortName from PE_Bank where IsDisabled=" & PE_False & " and IsDefault =" & PE_True & "")
        If Not (rsBank.EOF And rsBank.BOF) Then
            DefaultBank = rsBank("BankShortName")
        End If
        Set rsBank = Nothing
        Response.Write "            <div class=""link_text""><nobr><label id=""text1"">" & DefaultBank & "</label></nobr>" & vbCrLf
        Response.Write "            </div>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "            <td align=""right"" width=""20"">" & vbCrLf
        Response.Write "            <div onmouseup=""this.className='link_arrow0'"" class=""link_arrow0"" onmousedown=""this.className='link_arrow1'""" & vbCrLf
        Response.Write "            onmouseout=""this.className='link_arrow0'"">6" & vbCrLf
        Response.Write "            </div>" & vbCrLf
        Response.Write "            </td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "        </table>" & vbCrLf
        Response.Write "        </div>" & vbCrLf

       
        Set rsBank = Conn.Execute("select BankShortName,BankPic,IsDefault from PE_Bank where IsDisabled=" & PE_False & " order by OrderID asc")
        Response.Write "        <div class=""link_value"" id=""value1"" style=""WIDTH: 400px; HEIGHT: 150px;background: #ffffff;"" >" & vbCrLf
        Response.Write "        <table cellspacing=""0"" rules=""all"" oldValue=""oldlace"" singleValue=""#FFFFFF"" border=""1"" id=""DataGrid1"" style=""font-size:12px;width:100%;border-collapse:collapse;"">" & vbCrLf
        Do While Not rsBank.EOF
            If Trim(rsBank("BankPic")) <> "" Then
                Response.Write "          <tr onmouseover=""this.bgColor='#C1D2EE'"" onclick=""document.all.text1.innerText=this.cells[0].innerText;document.all.form4.Bank.value=this.cells[0].innerText;"" bgcolor=""#FFFFFF"" onmouseout=""this.bgColor=document.getElementById('DataGrid1').getAttribute('singleValue')"">" & vbCrLf
                Response.Write "            <td style=""width:300px;""><img src='" & Trim(rsBank("BankPic")) & "' align='absmiddle'>" & Trim(rsBank("BankShortName")) & "</td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            Else
                Response.Write "          <tr onmouseover=""this.bgColor='#C1D2EE'"" onclick=""document.all.text1.innerText=this.cells[0].innerText;document.all.form4.Bank.value=this.cells[0].innerText;"" bgcolor=""#FFFFFF"" onmouseout=""this.bgColor=document.getElementById('DataGrid1').getAttribute('singleValue')"">" & vbCrLf
                Response.Write "            <td style=""width:300px;"">" & Trim(rsBank("BankShortName")) & "</td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            rsBank.MoveNext
        Loop
        Set rsBank = Nothing
        If DefaultBank = "" Then
            DefaultBank = "您尚未选择银行汇款"
        End If
        Response.Write "          </table>" & vbCrLf
        Response.Write "        </div>" & vbCrLf
        Response.Write "        </div>" & vbCrLf
        Response.Write "        <Input type=""hidden"" value='" & DefaultBank & "' name=""Bank"">" & vbCrLf
        Response.Write "      </td>"
        Response.Write "    </tr>"
    Else
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>收入金额：</td>"
        Response.Write "      <td align='left'><input name='Money' type='text' id='Money' size='10' maxlength='10'> 元</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>关联订单号：</td>"
    Response.Write "      <td align='left'><input name='OrderFormNum' type='text' id='OrderFormNum' size='30' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>备注：</td>"
    Response.Write "      <td align='left'><input name='Remark' type='text' id='Remark' size='50' maxlength='200'></td>"
    Response.Write "    </tr>"
    If EnableSMS = True Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'></td>"
        Response.Write "      <td align='left'><input type='checkbox' name='SendSMSToUser' value='Yes'>同时发送手机短信通知会员</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='30' colspan='2'><font color='#FF0000'>注意：汇款/收入信息一旦录入，就不能再修改或删除！所以在保存之前确认输入无误！</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td height='30' colspan='2'>"
    Response.Write "      <input name='UserID' type='hidden' id='UserID' value='" & UserID & "'>"
    If IncomeType = 1 Then
        Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveRemit'>"
        Response.Write "      <input type='submit' name='Submit' value=' 保存汇款信息 '></td>"
    Else
        Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveIncome'>"
        Response.Write "      <input type='submit' name='Submit' value=' 保存收入信息 '></td>"
    End If
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    Set rsUser = Nothing
End Sub


Sub AddPayment()
    If AdminPurview > 1 And arrPurview(6) = False And arrPurview(15) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, Balance, OrderFormID, Money, trs, Remark
    Dim rsOrderForm, sqlOrderForm
    UserID = PE_CLng(Trim(Request("UserID")))
    UserName = ReplaceBadChar(Trim(Request("UserName")))
    OrderFormID = PE_CLng(Trim(Request("OrderFormID")))
    If UserID <= 0 And UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员！"
    End If
    
    If FoundErr = True Then Exit Sub
    
    If UserID > 0 Then
        Set trs = Conn.Execute("select UserID,UserName,Balance from PE_User where UserID=" & UserID & "")
    Else
        Set trs = Conn.Execute("select UserID,UserName,Balance from PE_User where UserName='" & UserName & "'")
    End If
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserID = trs(0)
        UserName = trs(1)
        Balance = trs(2)
    End If
    Set trs = Nothing

    If Balance <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>会员的资金余额不足！</li>"
    End If
    
    If FoundErr = True Then Exit Sub
    
    If OrderFormID > 0 Then
        sqlOrderForm = "select OrderFormNum,Status,MoneyTotal,MoneyReceipt from PE_OrderForm where OrderFormID=" & OrderFormID
        Set rsOrderForm = Conn.Execute(sqlOrderForm)
        If rsOrderForm.BOF And rsOrderForm.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的订单！</li>"
            rsOrderForm.Close
            Set rsOrderForm = Nothing
            Exit Sub
        End If
        If rsOrderForm("MoneyTotal") <= rsOrderForm("MoneyReceipt") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此订单已经付清，无需再支付！</li>"
            rsOrderForm.Close
            Set rsOrderForm = Nothing
            Exit Sub
        End If
        Money = rsOrderForm("MoneyTotal") - rsOrderForm("MoneyReceipt")
        If Balance <= Money Then
            Money = Balance
        End If
        Remark = "支付订单费用。订单编号：" & rsOrderForm("OrderFormNum")
    End If

    Response.Write "<form name='form4' method='post' action='Admin_User.asp' onsubmit=""return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')"">"
    Response.Write "  <table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    Response.Write "      <td height='22' colspan='2'>添 加 会 员 支 出 信 息</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>会 员 名：</td>"
    Response.Write "      <td align='left'><input name='UserName' type='text' id='UserName' value='" & UserName & "' size='30' maxlength='50' disabled></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>资金余额：</td>"
    Response.Write "      <td align='left'><input name='Balance' type='text' id='Balance' value='" & Balance & "' size='30' maxlength='50' disabled></td>"
    Response.Write "    </tr>"
    If OrderFormID > 0 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>支付内容：</td>"
        Response.Write "      <td align='left'><table  border='0' cellspacing='2' cellpadding='0'>"
        Response.Write "        <tr>"
        Response.Write "          <td align='right' width='15%' class='tdbg5'>订单编号：</td>"
        Response.Write "          <td align='left'><input name='OrderFormNum' type='text' id='OrderFormNum' value='" & rsOrderForm("OrderFormNum") & "' size='30' maxlength='30' disabled></td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td align='right' width='15%' class='tdbg5'>订单金额：</td>"
        Response.Write "          <td align='left'><input name='MoneyTotal' type='text' id='MoneyTotal' value='" & rsOrderForm("MoneyTotal") & "元' size='30' maxlength='30' disabled></td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td align='right' width='15%' class='tdbg5'>已 付 款：</td>"
        Response.Write "          <td align='left'><input name='MoneyTotal' type='text' id='MoneyTotal' value='" & rsOrderForm("MoneyReceipt") & "元' size='30' maxlength='30' disabled></td>"
        Response.Write "        </tr>"
        Response.Write "      </table></td>"
        Response.Write "    </tr>"

        rsOrderForm.Close
        Set rsOrderForm = Nothing
    Else
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>关联订单号：</td>"
        Response.Write "      <td align='left'><input name='OrderFormNum' type='text' id='OrderFormNum' size='30' maxlength='50'></td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>支出金额：</td>"
    Response.Write "      <td align='left'><input name='Money' type='text' id='Money' value='" & Money & "' size='10' maxlength='10'> 元</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>备注：</td>"
    Response.Write "      <td align='left'><input name='Remark' type='text' id='Remark' size='50' maxlength='200' value='" & Remark & "'></td>"
    Response.Write "    </tr>"
    If EnableSMS = True Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'></td>"
        Response.Write "      <td align='left'><input type='checkbox' name='SendSMSToUser' value='Yes'>同时发送手机短信通知会员</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='30' colspan='2'><font color='#FF0000'>注意：支出信息一旦录入，就不能再修改或修改！所以在保存之前确认输入无误！</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td height='30' colspan='2'><input name='Action' type='hidden' id='Action' value='SavePayment'>"
    Response.Write "      <input name='OrderFormID' type='hidden' id='OrderFormID' value='" & OrderFormID & "'>"
    Response.Write "      <input name='UserID' type='hidden' id='UserID' value='" & UserID & "'>"
    Response.Write "      <input type='submit' name='Submit' value=' 保存支付信息 '></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Exchange()
    If AdminPurview > 1 Then
        If (arrPurview(7) = False And InStr(Action, "Point") = 0) Or (arrPurview(8) = False And InStr(Action, "Valid") = 0) Then
            Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
            Call WriteEntry(6, AdminName, "越权操作")
            Exit Sub
        End If
    End If
    Dim UserID, strTitle, strCommond, strDisabled, strAction
    Dim rsUser, sqlUser
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    End If
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 3
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If
    Select Case Action
    Case "AddPoint"
        strTitle = "奖励" & PointName & ""
        strAction = "DoAddPoint"
        strCommond = "执行奖励"
    Case "MinusPoint"
        strTitle = "扣除" & PointName & ""
        strAction = "DoMinusPoint"
        strCommond = "执行扣除"
    Case "ExchangePoint"
        strTitle = "会员" & PointName & "兑换"
        strAction = "DoExchangePoint"
        strCommond = "执行兑换"
    Case "AddValid"
        strTitle = "奖励有效期"
        strAction = "DoAddValid"
        strCommond = "执行奖励"
    Case "MinusValid"
        strTitle = "扣除有效期"
        strAction = "DoMinusValid"
        strCommond = "执行扣除"
    Case "ExchangeValid"
        strTitle = "兑换有效期"
        strAction = "DoExchangeValid"
        strCommond = "执行兑换"
    End Select

    Response.Write "<form name='myform' action='Admin_User.asp' method='post' onSubmit=""return confirm('确定所有数据都录入正确吗？一旦录入就不能再修改哦！');"">" & vbCrLf
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height=22 colSpan=2 align='center'>" & strTitle & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>用户名：</td>" & vbCrLf
    Response.Write "      <td>" & rsUser("UserName") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>所属会员组：</td>" & vbCrLf
    Response.Write "      <td>" & GetGroupName(rsUser("GroupID")) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>资金余额：</td>" & vbCrLf
    Response.Write "      <td>" & rsUser("Balance") & " 元</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>可用" & PointName & "数：</td>" & vbCrLf
    Response.Write "      <td>" & rsUser("UserPoint") & " " & PointUnit & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>有效期限信息：</td>" & vbCrLf
    Response.Write "      <td>开始计算日期：" & FormatDateTime(rsUser("BeginTime"), 2) & "&nbsp;&nbsp;&nbsp;&nbsp;有效期："
    If rsUser("ValidNum") = -1 Then
        Response.Write "无限期"
    Else
        Response.Write rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "&nbsp;&nbsp;&nbsp;&nbsp;"
        ValidDays = ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime"))
        If ValidDays >= 0 Then
            Response.Write "尚有 <font color=blue>" & ValidDays & "</font> 天到期"
        Else
            Response.Write "已经过期 <font color=red>" & Abs(ValidDays) & "</font> 天"
        End If
    End If
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Select Case Action
    Case "AddPoint"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>奖励" & PointName & "数：</td>" & vbCrLf
        Response.Write "      <td><input name='Point' type='text' id='Point' value='100' size='6' maxlength='8' style='text-align:center'> " & PointUnit & "" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>奖励原因：</td>" & vbCrLf
        Response.Write "      <td><input name='Reason' type='text' id='Reason' value='' size='50' maxlength='100'>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
    Case "MinusPoint"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>扣除" & PointName & "数：</td>" & vbCrLf
        Response.Write "      <td><input name='Point' type='text' id='Point' value='100' size='6' maxlength='8' style='text-align:center'> " & PointUnit & "" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>扣除" & PointName & "原因：</td>" & vbCrLf
        Response.Write "      <td><input name='Reason' type='text' id='Reason' value='' size='50' maxlength='100'>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
    Case "ExchangePoint"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>增加" & PointName & "：</td>" & vbCrLf
        Response.Write "      <td><input name='Point' type='text' id='Point' value='100' size='6' maxlength='8' style='text-align:center'> " & PointUnit & "" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>同时减去：</td>" & vbCrLf
        Response.Write "      <td><input name='Money' type='text' value='" & (100 * MoneyExchangePoint) & "' size='6' maxlength='8' style='text-align:center'> 元&nbsp;&nbsp;&nbsp;&nbsp;资金与" & PointName & "的默认比率：" & FormatNumber(MoneyExchangePoint, 2, vbTrue, vbFalse, vbTrue) & "元:1" & PointName & "</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf

    Case "ExchangeValid"
        If rsUser("ValidNum") = -1 Then
            strDisabled = " disabled"
        Else
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='15%' class='tdbg5' align='right'>追加有效期：</td>" & vbCrLf
            Response.Write "      <td><input type='radio' name='ValidType' value='1' checked> 指定期限：<input name='ValidNum' type='text' id='ValidNum' value='10' size='6' maxlength='8' style='text-align:center'>"
            Response.Write "      <select name='ValidUnit' id='ValidUnit'><option value='1' "
            If rsUser("ValidUnit") = 1 Then Response.Write " selected"
            Response.Write ">天</option><option value='2' "
            If rsUser("ValidUnit") = 2 Then Response.Write " selected"
            Response.Write ">月</option><option value='3' "
            If rsUser("ValidUnit") = 3 Then Response.Write " selected"
            Response.Write ">年</option></select><br>&nbsp;&nbsp;&nbsp;&nbsp;若目前会员尚未到期，则追加相应天数<br>&nbsp;&nbsp;&nbsp;&nbsp;若目前会员已经过了有效期，则有效期从续费之日起重新计数。<br>" & vbCrLf
            Response.Write "<input type='radio' name='ValidType' value='2'"
            If rsUser("ValidNum") = -1 Then Response.Write " disabled"
            Response.Write "> 改为无限期</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='15%' class='tdbg5' align='right'>同时减去：</td>" & vbCrLf
            Response.Write "      <td><input name='Money' type='text' value='10' size='6' maxlength='8' style='text-align:center'> 元</td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf
        End If
    Case "AddValid"
        If rsUser("ValidNum") = -1 Then
            strDisabled = " disabled"
        Else
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='15%' class='tdbg5' align='right'>追加有效期：</td>" & vbCrLf
            Response.Write "      <td><input type='radio' name='ValidType' value='1' checked> 指定期限：<input name='ValidNum' type='text' id='ValidNum' value='10' size='6' maxlength='8' style='text-align:center'>"
            Response.Write "      <select name='ValidUnit' id='ValidUnit'><option value='1' "
            If rsUser("ValidUnit") = 1 Then Response.Write " selected"
            Response.Write ">天</option><option value='2' "
            If rsUser("ValidUnit") = 2 Then Response.Write " selected"
            Response.Write ">月</option><option value='3' "
            If rsUser("ValidUnit") = 3 Then Response.Write " selected"
            Response.Write ">年</option></select><br>&nbsp;&nbsp;&nbsp;&nbsp;若目前会员尚未到期，则追加相应天数<br>&nbsp;&nbsp;&nbsp;&nbsp;若目前会员已经过了有效期，则有效期从续费之日起重新计数。<br>" & vbCrLf
            Response.Write "<input type='radio' name='ValidType' value='2'"
            If rsUser("ValidNum") = -1 Then Response.Write " disabled"
            Response.Write "> 改为无限期</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='15%' class='tdbg5' align='right'>请输入原因：</td>" & vbCrLf
            Response.Write "      <td><input name='Reason' type='text' id='Reason' value='' size='50' maxlength='100'>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf
        End If

    Case "MinusValid"
        If rsUser("ValidNum") = -1 Then
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='15%' class='tdbg5' align='right'>减去有效期：</td>" & vbCrLf
            Response.Write "      <td><input type='radio' name='ValidType' value='1' disabled> 指定时间：<input name='ValidNum' type='text' id='ValidNum' value='10' size='6' maxlength='8' style='text-align:center' disabled>"
            Response.Write "      <select name='ValidUnit' id='ValidUnit' disabled><option value='1'>天</option></select><br>" & vbCrLf
            Response.Write "<input type='radio' name='ValidType' value='2' checked> 有效期归零（当某个会员的有效期是“无限期”时，可以通过归零重设其有效期期限）</td>"
            Response.Write "    </tr>" & vbCrLf
        Else
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='15%' class='tdbg5' align='right'>减去有效期：</td>" & vbCrLf
            Response.Write "      <td><input type='radio' name='ValidType' value='1' checked> 指定时间：<input name='ValidNum' type='text' id='ValidNum' value='10' size='6' maxlength='8' style='text-align:center'>"
            Response.Write "      <select name='ValidUnit' id='ValidUnit'><option value='1' "
            If rsUser("ValidUnit") = 1 Then Response.Write " selected"
            Response.Write ">天</option><option value='2' "
            If rsUser("ValidUnit") = 2 Then Response.Write " selected"
            Response.Write ">月</option><option value='3' "
            If rsUser("ValidUnit") = 3 Then Response.Write " selected"
            Response.Write ">年</option></select><br>" & vbCrLf
            Response.Write "<input type='radio' name='ValidType' value='2'> 有效期归零（当某个会员的有效期是“无限期”时，可以通过归零重设其有效期期限）</td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>请输入原因：</td>" & vbCrLf
        Response.Write "      <td><input name='Reason' type='text' id='Reason' value='' size='50' maxlength='100'>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
    End Select
    If EnableSMS = True Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'></td>"
        Response.Write "      <td align='left'><input type='checkbox' name='SendSMSToUser' value='Yes'>同时发送手机短信通知会员</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='" & strAction & "'>" & vbCrLf
    Response.Write "        <input name=Submit   type=submit id='Submit' value='" & strCommond & "'" & strDisabled & "> <input name='UserID' type='hidden' id='UserID' value='" & rsUser("UserID") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </TABLE>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    rsUser.Close
    Set rsUser = Nothing
End Sub


Sub UpdateUser()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Response.Write "<form name='myform' action='Admin_User.asp' method='post'>"
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td height='22' colspan='2'>更 新 会 员 数 据</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "      <td colspan='2'><p>说明：<br>"
    Response.Write "          1、本操作将重新计算会员的发表文章数。<br>"
    Response.Write "          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。"
    If API_Enable Then
        Response.Write "<br>          3、本操作将不会对整合的其他系统的数据进行同步！"
    End If
    Response.Write "</p>"
    Response.Write "      </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='25'>开始会员ID：</td>"
    Response.Write "    <td height='25'><input name='BeginID' type='text' id='BeginID' value='1' size='10' maxlength='10'>"
    Response.Write "      会员ID，可以填写您想从哪一个ID号开始进行修复</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='25'>结束会员ID：</td>"
    Response.Write "    <td height='25'><input name='EndID' type='text' id='EndID' value='1000' size='10' maxlength='10'>"
    Response.Write "      将更新开始到结束ID之间的会员数据，之间的数值最好不要选择过大</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='25'>&nbsp;</td>"
    Response.Write "    <td height='25'><input name='Submit' type='submit' id='Submit' value='更新会员数据'> <input name='Action' type='hidden' id='Action' value='DoUpdate'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub Batch()
    If AdminPurview > 1 Then
        If (arrPurview(4) = False And InStr(Action, "Del") > 0) Or (arrPurview(6) = False And InStr(Action, "Money") > 0) Or (arrPurview(7) = False And InStr(Action, "Point") > 0) Or (arrPurview(8) = False And InStr(Action, "Valid") > 0) Then
            Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
            Call WriteEntry(6, AdminName, "越权操作")
            Exit Sub
        End If
    End If

    Response.Write "<form method='POST' name='myform' action='Admin_User.asp' onsubmit=""return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')"">"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'>"
    Dim strActionName
    Select Case Action
    Case "BatchAddMoney"
        Response.Write "批量添加奖金"
        strActionName = "奖励"
    Case "BatchMinusMoney"
        Response.Write "批量扣除奖金"
        strActionName = "扣除"
    Case "BatchAddPoint"
        Response.Write "批量奖励" & PointName & ""
        strActionName = "奖励"
    Case "BatchMinusPoint"
        Response.Write "批量扣除" & PointName & ""
        strActionName = "扣除"
    Case "BatchAddValid"
        Response.Write "批量奖励有效期"
        strActionName = "奖励"
    Case "BatchMinusValid"
        Response.Write "批量扣除有效期"
        strActionName = "扣除"
    Case "BatchDel"
        Response.Write "删除会员"
        strActionName = "删除"
    End Select
    Response.Write "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' width='15%' class='tdbg5'>选择会员：</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='UserType' value='0'> 所有会员</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='UserType' value='1'> 指定会员组</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='UserType' value='2' checked> 指定会员ID</td><td><input type='text' name='UserID' size='80' value='" & Replace(Trim(Request("UserID")), " ", "") & "'></td></tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Select Case Action
    Case "BatchAddMoney", "BatchMinusMoney"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='right' width='15%' class='tdbg5'>" & strActionName & "金额：</td>"
        Response.Write "      <td align='left'><input name='Money' type='text' id='Money' value='10' size='10' maxlength='20'> 元"
        If Action = "BatchMinusMoney" Then
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;如果输入的扣除金额大于会员的资金余额，则扣除后，会员资金余额为0"
        End If
        Response.Write "</td>"
        Response.Write "    </tr>"
    Case "BatchAddPoint", "BatchMinusPoint"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right' width='15%' class='tdbg5'>" & strActionName & PointName & "：</td>" & vbCrLf
        Response.Write "      <td align='left'><input name='Point' type='text' id='Point' value='100' size='6' maxlength='8' style='text-align:center'> " & PointUnit & "" & vbCrLf
        If Action = "BatchMinusMoney" Then
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;如果输入的扣除" & PointName & "数大于会员的现有" & PointName & "数，则扣除后，会员" & PointName & "数为0"
        End If
        Response.Write "</td>"
        Response.Write "    </tr>" & vbCrLf
    Case "BatchAddValid"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right' width='15%' class='tdbg5'>追加有效期：</td>" & vbCrLf
        Response.Write "      <td align='left'><input type='radio' name='ValidType' value='1' checked> 指定期限：<input name='ValidNum' type='text' id='ValidNum' value='10' size='6' maxlength='8' style='text-align:center'>"
        Response.Write "      <select name='ValidUnit' id='ValidUnit'><option value='1'>天</option><option value='2'>月</option><option value='3'>年</option></select><br>" & vbCrLf
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;若目前会员尚未到期，则追加相应天数<br>&nbsp;&nbsp;&nbsp;&nbsp;若目前会员已经过了有效期，则有效期从续费之日起重新计数。<br>"
        Response.Write "<input type='radio' name='ValidType' value='2'> 改为无限期</td>"
        Response.Write "    </tr>" & vbCrLf
    Case "BatchMinusValid"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right' width='15%' class='tdbg5'>减去有效期：</td>" & vbCrLf
        Response.Write "      <td align='left'><input type='radio' name='ValidType' value='1' checked> 指定时间：<input name='ValidNum' type='text' id='ValidNum' value='10' size='6' maxlength='8' style='text-align:center'>"
        Response.Write "      <select name='ValidUnit' id='ValidUnit'><option value='1'>天</option><option value='2'>月</option><option value='3'>年</option></select><br>" & vbCrLf
        Response.Write "<input type='radio' name='ValidType' value='2'> 有效期归零（当某个会员的有效期是“无限期”时，可以通过归零重设其有效期期限）</td>"
        Response.Write "    </tr>" & vbCrLf
    Case "BatchDel"
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right' width='15%' class='tdbg5'>排除条件：</td>" & vbCrLf
        Response.Write "      <td align='left'><table width='100%'><tr><td><input name='ExcludingBalance' type='checkbox' value='Yes' checked>不包括资金余额大于0的会员</td><td><input name='ExcludingPoint' type='checkbox' value='Yes' checked>不包括可用" & PointName & "数大于0的会员</td><td><input name='ExcludingExp' type='checkbox' value='Yes' checked>不包括可用积分大于0的会员</td></tr>" & vbCrLf
        Response.Write "<tr><td><input name='ExcludingValid' type='checkbox' value='Yes' checked>不包括有效期未到期的会员</td><td><input name='ExcludingOrder' type='checkbox' value='Yes' checked>不包括已下订单的会员</td><td><input name='ExcludingBankroll' type='checkbox' value='Yes' checked>不包括有资金明细记录的会员</td></tr>" & vbCrLf
        Response.Write "<tr><td><input name='ExcludingConsume' type='checkbox' value='Yes' checked>不包括有消费明细记录的会员</td><td><input name='ExcludingRecharge' type='checkbox' value='Yes' checked>不包括有有效期充值记录的会员</td><td><input name='ExcludingPayment' type='checkbox' value='Yes' checked>不包括有在线支付记录的会员</td></tr></table></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right' width='15%' class='tdbg5'>删除相关信息：</td>" & vbCrLf
        Response.Write "      <td align='left'><table width='100%'><tr><td><input name='DelOrder' type='checkbox' value='Yes'>相关订单</td><td><input name='DelPayment' type='checkbox' value='Yes'>相关在线支付记录</td><td><input name='DelBankroll' type='checkbox' value='Yes'>相关资金明细记录</td><td><input name='DelConsumeLog' type='checkbox' value='Yes'>" & PointName & "消费明细</td><td><input name='DelRechargeLog' type='checkbox' value='Yes'>有效期明细</td><td><input name='DelMessage' type='checkbox' value='Yes'>短消息</td></tr>" & vbCrLf
        Response.Write "<tr><td><input name='DelArticle' type='checkbox' value='Yes'>添加的文章</td><td><input name='DelSoft' type='checkbox' value='Yes'>添加的软件</td><td><input name='DelPhoto' type='checkbox' value='Yes'>添加的图片</td><td><input name='DelComment' type='checkbox' value='Yes'>发表的评论</td><td><input name='DelGuestbook' type='checkbox' value='Yes'>发表的留言</td><td><input name='DelFavorite' type='checkbox' value='Yes'>收藏的信息</td></tr></table></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
    End Select
    If Action <> "BatchDel" Then
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right' width='15%' class='tdbg5'>" & strActionName & "原因：</td>" & vbCrLf
        Response.Write "      <td align='left'><input name='Reason' type='text' id='Reason' value='' size='50' maxlength='100'>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td align='right'></td>" & vbCrLf
        Response.Write "      <td align='left'><input name='SaveItem' type='checkbox' value='Yes' checked>为每个会员记录明细记录（推荐）</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Do" & Action & "'>"
    Response.Write "        <input type='submit' name='Submit' value='执行批量操作'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    Response.Write "<font color='red'>注：本操作可能将非常消耗服务器资源，而且运行时间很长，请尽量少用，并尽量在网站访问量少时进行。</font>"
    If API_Enable Then
        Response.Write "<br><font color='red'>　　批量操作将不会同步其他整合系统的数据。</font>"
    End If
End Sub

Sub DoBatch()
    If AdminPurview > 1 Then
        If (arrPurview(4) = False And InStr(Action, "Del") > 0) Or (arrPurview(6) = False And InStr(Action, "Money") > 0) Or (arrPurview(7) = False And InStr(Action, "Point") > 0) Or (arrPurview(8) = False And InStr(Action, "Valid") > 0) Then
            Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
            Call WriteEntry(6, AdminName, "越权操作")
            Exit Sub
        End If
    End If

    Dim UserType, GroupID, UserID
    Dim Money, Point, ValidType, ValidNum, ValidUnit, Reason, SaveItem
    Dim ExcludingBalance, ExcludingPoint, ExcludingExp, ExcludingValid, ExcludingOrder, ExcludingBankroll, ExcludingConsume, ExcludingRecharge, ExcludingPayment
    Dim DelOrder, DelPayment, DelBankroll, DelConsumeLog, DelRechargeLog, DelArticle, DelSoft, DelPhoto, DelComment, DelGuestbook, DelFavorite, DelMessage
    UserType = PE_CLng(Trim(Request("UserType")))
    GroupID = Trim(Request("GroupID"))
    UserID = Trim(Request("UserID"))
    Money = PE_CDbl(Trim(Request("Money")))
    Point = PE_CLng(Trim(Request("Point")))
    ValidType = PE_CLng(Trim(Request("ValidType")))
    ValidNum = PE_CLng(Trim(Request("ValidNum")))
    ValidUnit = PE_CLng(Trim(Request("ValidUnit")))
    
    Reason = ReplaceBadChar(Trim(Request("Reason")))
    SaveItem = Trim(Request("SaveItem"))

    ExcludingBalance = Trim(Request("ExcludingBalance"))
    ExcludingPoint = Trim(Request("ExcludingPoint"))
    ExcludingExp = Trim(Request("ExcludingExp"))
    ExcludingValid = Trim(Request("ExcludingValid"))
    ExcludingOrder = Trim(Request("ExcludingOrder"))
    ExcludingBankroll = Trim(Request("ExcludingBankroll"))
    ExcludingConsume = Trim(Request("ExcludingConsume"))
    ExcludingRecharge = Trim(Request("ExcludingRecharge"))
    ExcludingPayment = Trim(Request("ExcludingPayment"))
    DelOrder = Trim(Request("DelOrder"))
    DelPayment = Trim(Request("DelPayment"))
    DelBankroll = Trim(Request("DelBankroll"))
    DelConsumeLog = Trim(Request("DelConsumeLog"))
    DelRechargeLog = Trim(Request("DelRechargeLog"))
    DelArticle = Trim(Request("DelArticle"))
    DelSoft = Trim(Request("DelSoft"))
    DelPhoto = Trim(Request("DelPhoto"))
    DelComment = Trim(Request("DelComment"))
    DelGuestbook = Trim(Request("DelGuestbook"))
    DelFavorite = Trim(Request("DelFavorite"))
    DelMessage = Trim(Request("DelMessage"))

    Select Case UserType
    Case 0
    
    Case 1
        If GroupID = "" Or IsValidID(GroupID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择会员组！"
        End If
    Case 2
        If UserID = "" Or IsValidID(UserID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定会员ID！"
        End If
    End Select
    Select Case Action
    Case "DoAddMoney", "DoMinusMoney"
        If Money <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入金额！</li>"
        End If
    Case "DoAddPoint", "DoMinusPoint"
        If Point <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入" & PointName & "！</li>"
        End If
    Case "DoAddValid", "DoMinusValid"
        If ValidType = 1 Then
            If ValidNum <= 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请输入金额！</li>"
            End If
        End If
    End Select
    
    If Action <> "DoBatchDel" And Reason = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入原因！</li>"
    End If
    
    If FoundErr = True Then Exit Sub
    
    Dim rsBatch, sqlBatch, iTemp, strMsg, tUserID, tUserName
    iTemp = 0
    Select Case UserType
    Case 0
        sqlBatch = "select * from PE_User order by UserID"
    Case 1
        sqlBatch = "select * from PE_User where GroupID in (" & GroupID & ")  order by UserID"
    Case 2
        sqlBatch = "select * from PE_User where UserID in (" & UserID & ") order by UserID"
    End Select
    FoundErr = False
    
    Set rsBatch = Server.CreateObject("adodb.recordset")
    rsBatch.Open sqlBatch, Conn, 1, 3
    Dim tempDelName
    Do While Not rsBatch.EOF
        Select Case Action
        Case "DoBatchDel"
            tUserID = rsBatch("UserID")
            tUserName = rsBatch("UserName")
            If ExcludingBalance = "Yes" And rsBatch("Balance") > 0 Then
                FoundErr = True
            End If
            If ExcludingPoint = "Yes" And rsBatch("UserPoint") > 0 Then
                FoundErr = True
            End If
            If ExcludingExp = "Yes" And rsBatch("UserExp") > 0 Then
                FoundErr = True
            End If
            If ExcludingValid = "Yes" And ChkValidDays(rsBatch("ValidNum"), rsBatch("ValidUnit"), rsBatch("BeginTime")) > 0 Then
                FoundErr = True
            End If
            If ExcludingOrder = "Yes" And PE_CLng(Conn.Execute("select count(0) from PE_OrderForm where UserName='" & tUserName & "'")(0)) > 0 Then
                FoundErr = True
            End If
            If ExcludingBankroll = "Yes" And PE_CLng(Conn.Execute("select count(0) from PE_BankrollItem where UserName='" & tUserName & "'")(0)) > 0 Then
                FoundErr = True
            End If
            If ExcludingConsume = "Yes" And PE_CLng(Conn.Execute("select count(0) from PE_ConsumeLog where UserName='" & tUserName & "'")(0)) > 0 Then
                FoundErr = True
            End If
            If ExcludingRecharge = "Yes" And PE_CLng(Conn.Execute("select count(0) from PE_RechargeLog where UserName='" & tUserName & "'")(0)) > 0 Then
                FoundErr = True
            End If
            If ExcludingPayment = "Yes" And PE_CLng(Conn.Execute("select count(0) from PE_Payment where UserName='" & tUserName & "'")(0)) > 0 Then
                FoundErr = True
            End If

            If FoundErr = False Then
                If DelOrder = "Yes" Then
                    Conn.Execute ("delete from PE_OrderForm where UserName='" & tUserName & "'")
                End If
                If DelPayment = "Yes" Then
                    Conn.Execute ("delete from PE_Payment where UserName='" & tUserName & "'")
                End If
                If DelBankroll = "Yes" Then
                    Conn.Execute ("delete from PE_BankrollItem where UserName='" & tUserName & "'")
                End If
                If DelConsumeLog = "Yes" Then
                    Conn.Execute ("delete from PE_ConsumeLog where UserName='" & tUserName & "'")
                End If
                If DelRechargeLog = "Yes" Then
                    Conn.Execute ("delete from PE_RechargeLog where UserName='" & tUserName & "'")
                End If
                If DelArticle = "Yes" Then
                    Conn.Execute ("delete from PE_Article where Inputer='" & tUserName & "'")
                End If
                If DelSoft = "Yes" Then
                    Conn.Execute ("delete from PE_Soft where Inputer='" & tUserName & "'")
                End If
                If DelPhoto = "Yes" Then
                    Conn.Execute ("delete from PE_Photo where Inputer='" & tUserName & "'")
                End If
                If DelComment = "Yes" Then
                    Conn.Execute ("delete from PE_Comment where UserType=1 and UserName='" & tUserName & "'")
                End If
                If DelGuestbook = "Yes" Then
                    Conn.Execute ("delete from PE_GuestBook where GuestType=1 and GuestName='" & tUserName & "'")
                End If
                If DelFavorite = "Yes" Then
                    Conn.Execute ("delete from PE_Favorite where UserID=" & tUserID & "")
                End If
                If DelMessage = "Yes" Then
                    Conn.Execute ("delete from PE_Message where sender='" & tUserName & "' or incept='" & tUserName & "'")		
                End If										
		Dim rsMail,arrUserID,tempUserID,newarr,NeedUpdate
		NeedUpdate = False
                Set rsMail = Server.CreateObject("adodb.recordset")
                rsMail.Open "select * from PE_MailChannel", Conn, 1, 3
	  	Do While Not rsMail.Eof
		    arrUserID = split(rsMail("UserID"),",")
		    For tempUserID=0 to Ubound(arrUserID)
		        IF PE_Clng(arrUserID(tempUserID)) = tUserID Then
			    arrUserID(tempuserID) = ""
                            NeedUpdate = True
			End If		
		    Next
			newarr = ""
		    For tempUserID=0 to Ubound(arrUserID)
		        IF arrUserID(tempuserID)="" Then    
                        ElseIF newarr="" then
                            newarr = arrUserID(tempuserID)
                        Else
                            newarr = newarr&","&arrUserID(tempuserID)
			End If		
		    Next
                    IF NeedUpdate = True then
		        rsMail("UserID")=newarr
		        rsMail.update
		    End IF	
                    rsMail.movenext		    
		Loop
                rsBatch.Delete
                rsBatch.Update
                iTemp = iTemp + 1
                tempDelName = tempDelName & "," & tUserName
            End If
            FoundErr = False
        Case "DoBatchAddMoney"
            rsBatch("Balance") = rsBatch("Balance") + Money
            rsBatch.Update
            If SaveItem = "Yes" Then
                Call AddBankrollItem(AdminName, rsBatch("UserName"), rsBatch("ClientID"), Money, 4, "", 0, 1, 0, 0, Reason, Now())
            End If
        Case "DoBatchMinusMoney"
            If rsBatch("Balance") > Money Then
                iTemp = Money
                rsBatch("Balance") = rsBatch("Balance") - Money
            Else
                iTemp = rsBatch("Balance")
                rsBatch("Balance") = 0
            End If
            rsBatch.Update
            If SaveItem = "Yes" Then
                Call AddBankrollItem(AdminName, rsBatch("UserName"), rsBatch("ClientID"), iTemp, 4, "", 0, 2, 0, 0, Reason, Now())
            End If
        Case "DoBatchAddPoint"
            rsBatch("UserPoint") = rsBatch("UserPoint") + Point
            rsBatch.Update
            If SaveItem = "Yes" Then
                Call AddConsumeLog(AdminName, 0, rsBatch("UserName"), 0, Point, 1, Reason)
            End If
        Case "DoBatchMinusPoint"
            If rsBatch("UserPoint") > Point Then
                iTemp = Point
                rsBatch("UserPoint") = rsBatch("UserPoint") - Point
            Else
                iTemp = rsBatch("UserPoint")
                rsBatch("UserPoint") = 0
            End If
            rsBatch.Update
            If SaveItem = "Yes" And iTemp > 0 Then
                Call AddConsumeLog(AdminName, 0, rsBatch("UserName"), 0, iTemp, 2, Reason)
            End If
        Case "DoBatchAddValid"
            If rsBatch("ValidNum") = -1 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>会员“" & rsBatch("UserName") & "”的有效期已经是“无限期”，无需奖励。"
            Else
                If ValidType = 2 Then
                    rsBatch("ValidNum") = -1
                    rsBatch.Update
                    If SaveItem = "Yes" Then
                        Call AddRechargeLog(AdminName, rsBatch("UserName"), 0, 0, 0, "（奖励）将有效期变为“无限期”，原因：" & Reason & "")
                    End If
                Else
                    ValidDays = ChkValidDays(rsBatch("ValidNum"), rsBatch("ValidUnit"), rsBatch("BeginTime"))
                    If ValidDays > 0 Then
                        If rsBatch("ValidUnit") = ValidUnit Then
                            rsBatch("ValidNum") = rsBatch("ValidNum") + ValidNum
                        ElseIf rsBatch("ValidUnit") < ValidUnit Then
                            If rsBatch("ValidUnit") = 1 Then
                                If ValidUnit = 2 Then
                                    rsBatch("ValidNum") = rsBatch("ValidNum") + ValidNum * 30
                                Else
                                    rsBatch("ValidNum") = rsBatch("ValidNum") + ValidNum * 365
                                End If
                            Else
                                rsBatch("ValidNum") = rsBatch("ValidNum") + ValidNum * 12
                            End If
                        Else
                            If ValidUnit = 1 Then
                                If rsBatch("ValidUnit") = 2 Then
                                    rsBatch("ValidNum") = ValidNum + rsBatch("ValidNum") * 30
                                Else
                                    rsBatch("ValidNum") = ValidNum + rsBatch("ValidNum") * 365
                                End If
                            Else
                                rsBatch("ValidNum") = ValidNum + rsBatch("ValidNum") * 12
                            End If
                            rsBatch("ValidUnit") = ValidUnit
                            If SaveItem = "Yes" Then
                                Call AddRechargeLog(AdminName, rsBatch("UserName"), 0, 0, 0, "添加有效期时更改有效期计费单位")
                            End If
                        End If
                    Else
                        rsBatch("BeginTime") = FormatDateTime(Now(), 2)
                        rsBatch("ValidNum") = ValidNum
                        rsBatch("ValidUnit") = ValidUnit
                        If SaveItem = "Yes" Then
                            Call AddRechargeLog(AdminName, rsBatch("UserName"), 0, 0, 0, "添加有效期时将原来过期的有效期重新计算")
                        End If
                    End If
                    rsBatch.Update
                End If
                If SaveItem = "Yes" Then
                    Call AddRechargeLog(AdminName, rsBatch("UserName"), ValidNum, ValidUnit, 1, Reason)
                End If
            End If
        Case "DoBatchMinusValid"
            If ValidType = 2 Then
                rsBatch("BeginTime") = FormatDateTime(Now(), 2)
                rsBatch("ValidNum") = 0
                rsBatch("ValidUnit") = 1
                rsBatch.Update
                If SaveItem = "Yes" Then
                    Call AddRechargeLog(AdminName, rsBatch("UserName"), 0, 0, 2, "有效期归零，原因：" & Reason & "")
                End If
            Else
                If rsBatch("ValidNum") = -1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & rsBatch("UserName") & "的有效期是“无限期”，只能执行“有效期归零”的操作。</li>"
                Else
                    ValidDays = ChkValidDays(rsBatch("ValidNum"), rsBatch("ValidUnit"), rsBatch("BeginTime"))
                    If ValidDays <= 0 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>" & rsBatch("UserName") & "的有效期已经过期，跳过。</li>"
                    Else
                        If rsBatch("ValidUnit") = ValidUnit Then
                            rsBatch("ValidNum") = rsBatch("ValidNum") - ValidNum
                        ElseIf rsBatch("ValidUnit") < ValidUnit Then
                            If rsBatch("ValidUnit") = 1 Then
                                If ValidUnit = 2 Then
                                    rsBatch("ValidNum") = rsBatch("ValidNum") - ValidNum * 30
                                Else
                                    rsBatch("ValidNum") = rsBatch("ValidNum") - ValidNum * 365
                                End If
                            Else
                                rsBatch("ValidNum") = rsBatch("ValidNum") - ValidNum * 12
                            End If
                        Else
                            rsBatch("ValidUnit") = ValidUnit
                            If ValidUnit = 1 Then
                                rsBatch("ValidNum") = ValidDays - ValidNum
                            Else
                                rsBatch("ValidNum") = rsBatch("ValidNum") * 12 - ValidNum
                            End If
                            If SaveItem = "Yes" Then
                                Call AddRechargeLog(AdminName, rsBatch("UserName"), 0, 0, 0, "扣除有效期时更改有效期计费单位")
                            End If
                        End If
                        rsBatch.Update
                        If rsBatch("ValidNum") < 0 Then
                            rsBatch("ValidNum") = 0
                            rsBatch.Update
                        End If
                        If SaveItem = "Yes" Then
                            Call AddRechargeLog(AdminName, rsBatch("UserName"), ValidNum, ValidUnit, 2, Reason)
                        End If
                    End If
                End If
            End If
        End Select	
		
        rsBatch.Update 	        		
		rsBatch.MoveNext
    Loop
    rsBatch.Close
    Set rsBatch = Nothing
    Select Case Action
    Case "DoBatchAddMoney"
        strMsg = "批量发奖金成功！"
    Case "DoBatchMinusMoney"
        strMsg = "批量扣奖金成功！"
    Case "DoBatchAddPoint"
        strMsg = "批量奖励" & PointName & "成功！"
    Case "DoBatchMinusPoint"
        strMsg = "批量扣除" & PointName & "成功！"
    Case "DoBatchAddValid"
        strMsg = "批量奖励有效期成功！"
    Case "DoBatchMinusValid"
        strMsg = "批量扣除有效期成功！"
    Case "DoBatchDel"
        strMsg = "批量删除会员成功！共删除了 " & iTemp & " 名会员。"
        '如果启用了整合，则发送一个不需判断返回的批量删除请求
        If API_Enable Then
            Call API_DelUser(tempDelName)
        End If
    End Select
    Call WriteSuccessMsg(strMsg, ComeUrl)

End Sub

Sub RegCompany()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, rsUser
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
        Exit Sub
    End If
    Set rsUser = Conn.Execute("select * from PE_User where UserID=" & UserID & "")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

    If rsUser("UserType") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>此会员已经是企业会员！</li>"
    End If
    UserName = rsUser("UserName")
    rsUser.Close
    Set rsUser = Nothing

    If FoundErr = True Then Exit Sub

    Response.Write "<br>您现在的位置：<a href='Admin_User.asp'>会员管理</a> >> 将个人会员 <font color='red'>" & UserName & "</font> 升级为企业会员"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectCompany(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=CompanyList','','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        var ss=arr.split('$$$');" & vbCrLf
    Response.Write "        document.myform.CompanyName.value=ss[0];" & vbCrLf
    Response.Write "        document.myform.CompanyID.value=ss[1];" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "    document.myform2.Country.value=frm1.document.regionform.Country.value;" & vbCrLf
    Response.Write "    document.myform2.Province.value=frm1.document.regionform.Province.value;" & vbCrLf
    Response.Write "    document.myform2.City.value=frm1.document.regionform.City.value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    
    Response.Write "<form name='myform' action='Admin_User.asp' method='post'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='title'>" & vbCrLf
    Response.Write "      <td height='22' colSpan='10'><b>方式一：将会员加入已有企业</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='250' align='right' class='tdbg5'>请选择要加入的企业：</td>" & vbCrLf
    Response.Write "      <td><input name='CompanyName' type='text' maxLength='200' size='50' value='请点击右边的按钮选择要加入的企业' readonly><input type='button' name='Show' value='...' onclick='SelectCompany()'><input type='hidden' name='CompanyID' value='0'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='250' align='right' class='tdbg5'>加入后的成员级别：</td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='UserType' value='2'>管理员&nbsp;&nbsp;<input type='radio' name='UserType' value='3' checked>普通成员&nbsp;&nbsp;<input type='radio' name='UserType' value='4'>待审核成员</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' height='40' align='center'><td colspan='10'><input type='submit' name='Join' value='加入此企业'><input type='hidden' name='Action' value='Join'><input type='hidden' name='UserID' value='" & UserID & "'></td></tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    
    Dim arrStatusInField, arrCompanySize, arrManagementForms
    arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
    arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
    arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")
    
    Response.Write "<form name='myform2' action='Admin_User.asp' method='post' onsubmit='CheckSubmit()'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='title'>" & vbCrLf
    Response.Write "      <td height='22' colSpan='10'><b>方式二：创建新企业并加入</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' width='12%'>单位名称：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='CompanyName' type='text' size='35' maxlength='200' value=''></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>单位简称：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='ShortedForm' type='text' size='35' maxlength='30' value=''></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>通讯地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <iframe name='frm' id='frm1' src='../Region.asp?Action=Modify&Country=&Province=&City=' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "                            <input name='Country' type='hidden'> <input name='Province' type='hidden'> <input name='City' type='hidden'>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
    Response.Write "                                    <td><input name='Address' type='text' size='60' maxlength='255' value=''></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td align='right' class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
    Response.Write "                                    <td><input name='ZipCode' type='text' size='35' maxlength='10' value=''></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                            </table>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>联系电话：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Phone' type='text' size='35' maxlength='30' value=''></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>传真号码：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Fax' type='text' size='35' maxlength='30' value=''></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>开户银行：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='BankOfDeposit' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>银行帐号：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='BankAccount' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >税号：</td>" & vbCrLf
    Response.Write "                        <td><input name='TaxNum' type='text' size='35' maxlength='20' value=''></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >网址：</td>" & vbCrLf
    Response.Write "                        <td><input name='Homepage' type='text' size='35' maxlength='100' value=''></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >行业地位：</td>" & vbCrLf
    Response.Write "                        <td><select name='StatusInField'>" & Array2Option(arrStatusInField, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >公司规模：</td>" & vbCrLf
    Response.Write "                        <td><select name='CompanySize'>" & Array2Option(arrCompanySize, -1) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >业务范围：</td>" & vbCrLf
    Response.Write "                        <td><input name='BusinessScope' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >年销售额：</td>" & vbCrLf
    Response.Write "                        <td><input name='AnnualSales' type='text' size='15' maxlength='20' value=''> 万元</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >经营状态：</td>" & vbCrLf
    Response.Write "                        <td><select name='ManagementForms'>" & Array2Option(arrManagementForms, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >注册资本：</td>" & vbCrLf
    Response.Write "                        <td><input name='RegisteredCapital' type='text' size='15' maxlength='20' value=''> 万元</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >公司照片：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'><input name='CompamyPic' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >公司简介：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'><textarea name='CompanyIntro' cols='75' rows='5' id='CompanyIntro'></textarea></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "<tr class='tdbg' height='50'><td colspan='10' align='center'><input type='submit' name='Join' value='创建此企业并加入'><input type='hidden' name='Action' value='SaveRegCompany'><input type='hidden' name='UserID' value='" & UserID & "'></td></tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub Up2Client()
    If AdminPurview > 1 And arrPurview(5) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, rsUser, ShortedForm, ClientType
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
        Exit Sub
    End If
    Set rsUser = Conn.Execute("select * from PE_User where UserID=" & UserID & "")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If
    If rsUser("ClientID") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>此会员已经是客户！</li>"
    End If
    If rsUser("ContacterID") = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>此会员还没有填写详细的联系资料，不能升级为客户！</li>"
    Else
        If rsUser("UserType") = 0 Then
            ClientType = 1
            Dim rsContacter
            Set rsContacter = Conn.Execute("select TrueName From PE_Contacter where ContacterID=" & rsUser("ContacterID") & "")
            If rsContacter.BOF And rsContacter.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>此会员还没有填写详细的联系资料，不能升级为客户！</li>"
            Else
                ShortedForm = rsContacter(0)
            End If
            rsContacter.Close
            Set rsContacter = Nothing
        Else
            ClientType = 0
            Dim rsCompany
            Set rsCompany = Conn.Execute("select CompanyName From PE_Company where CompanyID=" & rsUser("CompanyID") & "")
            If rsCompany.BOF And rsCompany.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到对应的企业信息，不能升级为企业客户！</li>"
            Else
                ShortedForm = Left(rsCompany(0), 8)
            End If
            rsCompany.Close
            Set rsCompany = Nothing
        End If
    End If
    UserName = rsUser("UserName")
    rsUser.Close
    Set rsUser = Nothing

    If FoundErr = True Then Exit Sub

    Dim arrArea, arrClientField, arrValueLevel, arrCreditLevel, arrImportance
    Dim arrConnectionLevel, arrGroupID, arrSourceType, arrPhaseType
    Dim arrClientType
    arrArea = GetArrFromDictionary("PE_Client", "Area")
    arrClientField = GetArrFromDictionary("PE_Client", "ClientField")
    arrValueLevel = GetArrFromDictionary("PE_Client", "ValueLevel")
    arrCreditLevel = GetArrFromDictionary("PE_Client", "CreditLevel")
    arrImportance = GetArrFromDictionary("PE_Client", "Importance")
    arrConnectionLevel = GetArrFromDictionary("PE_Client", "ConnectionLevel")
    arrGroupID = GetArrFromDictionary("PE_Client", "GroupID")
    arrSourceType = GetArrFromDictionary("PE_Client", "SourceType")
    arrPhaseType = GetArrFromDictionary("PE_Client", "PhaseType")
    arrClientType = Array("企业客户", "个人客户")

    Response.Write "<br>您现在的位置：<a href='Admin_User.asp'>会员管理</a> >> 将会员 <font color='red'>" & UserName & "</font> 升级为客户"
    
    Response.Write "<form name='myform' action='Admin_User.asp' method='post'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='title'>" & vbCrLf
    Response.Write "      <td height='22' colSpan='10'><b>请输入客户信息</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >助记名称：</td>" & vbCrLf
    Response.Write "                        <td><input name='ShortedForm' type='text' id='ShortedForm' size='35' maxlength='20' value='" & ShortedForm & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>客户编号：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='ClientNum' type='text' id='ClientNum' size='35' maxlength='30' value='" & GetNumString() & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >区域：</td>" & vbCrLf
    Response.Write "                        <td><select name='Area'>" & Array2Option(arrArea, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >行业：</td>" & vbCrLf
    Response.Write "                        <td><select name='ClientField'>" & Array2Option(arrClientField, -1) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >价值评估：</td>" & vbCrLf
    Response.Write "                        <td><select name='ValueLevel'>" & Array2Option(arrValueLevel, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >信用等级：</td>" & vbCrLf
    Response.Write "                        <td><select name='CreditLevel'>" & Array2Option(arrCreditLevel, -1) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >重要程度：</td>" & vbCrLf
    Response.Write "                        <td><select name='Importance'>" & Array2Option(arrImportance, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >关系等级：</td>" & vbCrLf
    Response.Write "                        <td><select name='ConnectionLevel'>" & Array2Option(arrConnectionLevel, -1) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >客户来源：</td>" & vbCrLf
    Response.Write "                        <td><select name='SourceType'>" & Array2Option(arrSourceType, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >阶段：</td>" & vbCrLf
    Response.Write "                        <td><select name='PhaseType'>" & Array2Option(arrPhaseType, -1) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >客户组别：</td>" & vbCrLf
    Response.Write "                        <td><select name='GroupID'>" & Array2Option(arrGroupID, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >客户类别：</td>" & vbCrLf
    Response.Write "                        <td>" & arrClientType(ClientType) & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "<tr class='tdbg' height='50'><td colspan='10' align='center'><input type='submit' name='Up' value='保存客户信息'><input type='hidden' name='Action' value='SaveClient'><input type='hidden' name='UserID' value='" & UserID & "'></td></tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub JoinCompany()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, rsUser, UserType
    Dim CompanyID, CompanyName, rsCompany

    UserID = PE_CLng(Trim(Request("UserID")))
    CompanyID = PE_CLng(Trim(Request("CompanyID")))
    UserType = PE_CLng(Trim(Request("UserType")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
    End If
    If CompanyID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要加入的企业！</li>"
    End If
    If FoundErr = True Then Exit Sub

    Set rsUser = Conn.Execute("select UserType,UserName from PE_User where UserID=" & UserID & "")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserName = rsUser("UserName")
        If rsUser("UserType") > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此会员已经是企业会员！</li>"
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing

    Set rsCompany = Conn.Execute("select CompanyName from PE_Company where CompanyID=" & CompanyID & "")
    If rsCompany.BOF And rsCompany.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的企业！</li>"
    Else
        CompanyName = rsCompany(0)
    End If
    rsCompany.Close
    Set rsCompany = Nothing

    If FoundErr = True Then Exit Sub

    Conn.Execute ("update PE_User set UserType=" & UserType & ",CompanyID=" & CompanyID & " where UserID=" & UserID & "")
    
    Call WriteSuccessMsg("成功将 " & UserName & " 加入到企业 " & CompanyName & " 中！", "Admin_User.asp?Action=Show&UserID=" & UserID & "")
End Sub

Sub SaveRegCompany()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, rsUser, ClientID

    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
    End If
    If FoundErr = True Then Exit Sub

    Set rsUser = Conn.Execute("select UserType,UserName,ClientID from PE_User where UserID=" & UserID & "")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserName = rsUser("UserName")
        ClientID = rsUser("ClientID")
        If rsUser("UserType") > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此会员已经是企业会员！</li>"
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing
    If FoundErr = True Then Exit Sub

    Dim CompanyName, ShortedForm, Country, Province, City, Address, ZipCode, HomePage, Phone, Fax
    Dim BankOfDeposit, BankAccount, TaxNum, StatusInField, CompanySize, BusinessScope, AnnualSales, ManagementForms, RegisteredCapital
    Dim CompanyIntro, CompamyPic
    CompanyName = ReplaceBadChar(Trim(Request("CompanyName")))
    ShortedForm = ReplaceBadChar(Trim(Request("ShortedForm")))
    Country = Trim(Request.Form("Country"))
    Province = Trim(Request.Form("Province"))
    City = Trim(Request.Form("City"))
    Address = Trim(Request.Form("Address"))
    ZipCode = Trim(Request.Form("ZipCode"))
    Phone = Trim(Request.Form("Phone"))
    Fax = Trim(Request.Form("Fax"))
    HomePage = Trim(Request.Form("Homepage"))
    BankOfDeposit = Trim(Request.Form("BankOfDeposit"))
    BankAccount = Trim(Request.Form("BankAccount"))
    TaxNum = Trim(Request.Form("TaxNum"))
    StatusInField = PE_CLng(Trim(Request.Form("StatusInField")))
    CompanySize = PE_CLng(Trim(Request.Form("CompanySize")))
    BusinessScope = Trim(Request.Form("BusinessScope"))
    AnnualSales = Trim(Request.Form("AnnualSales"))
    ManagementForms = PE_CLng(Trim(Request.Form("ManagementForms")))
    RegisteredCapital = Trim(Request.Form("RegisteredCapital"))
    CompanyIntro = Trim(Request.Form("CompanyIntro"))
    CompamyPic = Trim(Request.Form("CompamyPic"))
    If CompanyName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入企业名称！</li>"
    End If
    If Address = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入单位的联系地址！</li>"
    End If
    If ZipCode = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入单位的邮政编码！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim sqlCompany, rsCompany, CompanyID
    CompanyID = GetNewID("PE_Company", "CompanyID")
    Set rsCompany = Server.CreateObject("adodb.recordset")
    sqlCompany = "select top 1 * From PE_Company"
    rsCompany.Open sqlCompany, Conn, 1, 3
    rsCompany.addnew
    rsCompany("CompanyID") = CompanyID
    rsCompany("ClientID") = ClientID
    rsCompany("CompanyName") = CompanyName
    rsCompany("Country") = Country
    rsCompany("Province") = Province
    rsCompany("City") = City
    rsCompany("Address") = Address
    rsCompany("ZipCode") = ZipCode
    rsCompany("Phone") = Phone
    rsCompany("Fax") = Fax
    rsCompany("HomePage") = HomePage
    rsCompany("BankOfDeposit") = BankOfDeposit
    rsCompany("BankAccount") = BankAccount
    rsCompany("TaxNum") = TaxNum
    rsCompany("StatusInField") = StatusInField
    rsCompany("CompanySize") = CompanySize
    rsCompany("BusinessScope") = BusinessScope
    rsCompany("AnnualSales") = AnnualSales
    rsCompany("ManagementForms") = ManagementForms
    rsCompany("RegisteredCapital") = RegisteredCapital
    rsCompany("CompanyIntro") = PE_HTMLEncode(CompanyIntro)
    rsCompany("CompamyPic") = PE_HTMLEncode(CompamyPic)
    rsCompany.Update
    rsCompany.Close
    Set rsCompany = Nothing
    Conn.Execute ("update PE_User set UserType=1,CompanyID=" & CompanyID & " where UserID=" & UserID & "")
    If ClientID > 0 Then
        Conn.Execute ("update PE_Client set ClientName='" & CompanyName & "',ShortedForm='" & ShortedForm & "',ClientType=0 where ClientID=" & ClientID & "")
    End If
    Call WriteSuccessMsg("成功创建了新企业：" & CompanyName & "<br>并将会员 " & UserName & " 设为这家企业的创建人，拥有这家企业的管理权限（如审核批准其他人的申请）。", "Admin_User.asp?Action=Show&UserID=" & UserID & "")
End Sub

Sub SaveClient()
    If AdminPurview > 1 And arrPurview(5) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, rsUser
    Dim ClientID, ClientType, ClientName
    Dim CompanyID, ContacterID, ContacterUserType

    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
    End If
    If FoundErr = True Then Exit Sub

    Set rsUser = Conn.Execute("select UserType,UserName,ClientID,CompanyID,ContacterID from PE_User where UserID=" & UserID & "")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserName = rsUser("UserName")
        ClientID = rsUser("ClientID")
        CompanyID = rsUser("CompanyID")
        ContacterID = rsUser("ContacterID")
        If ClientID > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此会员已经是客户！</li>"
        End If
        If ContacterID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此会员还没有填写详细的联系资料，不能升级为客户！</li>"
        Else
            If rsUser("UserType") = 0 Then
                ClientType = 1
                ContacterUserType = 0
                Dim rsContacter
                Set rsContacter = Conn.Execute("select TrueName From PE_Contacter where ContacterID=" & ContacterID & "")
                If rsContacter.BOF And rsContacter.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>此会员还没有填写详细的联系资料，不能升级为客户！</li>"
                Else
                    ClientName = rsContacter(0)
                End If
                rsContacter.Close
                Set rsContacter = Nothing
            Else
                ClientType = 0
                If rsUser("UserType") = 1 Then
                    ContacterUserType = 1
                Else
                    ContacterUserType = 2
                End If
                Dim rsCompany
                Set rsCompany = Conn.Execute("select CompanyName From PE_Company where CompanyID=" & CompanyID & "")
                If rsCompany.BOF And rsCompany.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>找不到对应的企业信息，不能升级为企业客户！</li>"
                Else
                    ClientName = rsCompany(0)
                End If
                rsCompany.Close
                Set rsCompany = Nothing
            End If
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing

    If FoundErr = True Then Exit Sub

    Dim ClientNum, ShortedForm
    Dim Area, ClientField, ValueLevel, CreditLevel, Importance, ConnectionLevel, SourceType, GroupID, ParentID, PhaseType, Remark

    ClientNum = Trim(Request.Form("ClientNum"))
    ShortedForm = Trim(Request.Form("ShortedForm"))
    Area = PE_CLng(Trim(Request.Form("Area")))
    ClientField = PE_CLng(Trim(Request.Form("ClientField")))
    ValueLevel = PE_CLng(Trim(Request.Form("ValueLevel")))
    CreditLevel = PE_CLng(Trim(Request.Form("CreditLevel")))
    Importance = PE_CLng(Trim(Request.Form("Importance")))
    ConnectionLevel = PE_CLng(Trim(Request.Form("ConnectionLevel")))
    SourceType = PE_CLng(Trim(Request.Form("SourceType")))
    GroupID = PE_CLng(Trim(Request.Form("GroupID")))
    ParentID = PE_CLng(Trim(Request.Form("ParentID")))
    GroupID = PE_CLng(Trim(Request.Form("GroupID")))
    PhaseType = PE_CLng(Trim(Request.Form("PhaseType")))

    If ShortedForm = "" Then
        FoundErr = True
        ErrMsg = "客户简称（助记码）不能为空！"
    End If

    If FoundErr = True Then Exit Sub

    ClientID = GetNewID("PE_Client", "ClientID")

    Dim sqlClient, rsClient
    sqlClient = "SELECT top 1 * FROM PE_Client"
    Set rsClient = Server.CreateObject("adodb.recordset")
    rsClient.Open sqlClient, Conn, 1, 3
    rsClient.addnew
    rsClient("ClientID") = ClientID
    rsClient("ClientName") = ClientName
    rsClient("ClientNum") = ClientNum
    rsClient("ClientType") = ClientType
    rsClient("ShortedForm") = ShortedForm
    rsClient("Area") = Area
    rsClient("ClientField") = ClientField
    rsClient("ValueLevel") = ValueLevel
    rsClient("CreditLevel") = CreditLevel
    rsClient("Importance") = Importance
    rsClient("ConnectionLevel") = ConnectionLevel
    rsClient("SourceType") = SourceType
    rsClient("GroupID") = GroupID
    rsClient("ParentID") = ParentID
    rsClient("PhaseType") = PhaseType
    rsClient("Remark") = Remark
    rsClient("UpdateTime") = Now()
    rsClient("CreateTime") = Now()
    rsClient("Owner") = AdminName
    rsClient.Update
    rsClient.Close
    Set rsClient = Nothing

    Conn.Execute ("update PE_User set ClientID=" & ClientID & " where UserID=" & UserID & "")
    Conn.Execute ("update PE_Company set ClientID=" & ClientID & " where CompanyID=" & CompanyID & "")
    Conn.Execute ("update PE_Contacter set ClientID=" & ClientID & ",UserType=" & ContacterUserType & " where ContacterID=" & ContacterID & "")

    Call WriteSuccessMsg("成功将会员 " & UserName & " 升级为客户！", "Admin_User.asp?Action=Show&UserID=" & UserID & "")
End Sub

Sub SaveUser()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim GroupID, GroupName, GroupType, ClientID, ContacterID, CompanyID, UserType
    Dim UserID, UserName, UserPassword, LastPassword, Question, Answer, Email
    Dim UserFace, FaceWidth, FaceHeight, Sign, Privacy
    
    UserID = PE_CLng(Trim(Request.Form("UserID")))
    GroupID = PE_CLng(Trim(Request.Form("GroupID")))
    UserName = UserNamefilter(Trim(Request.Form("UserName")))
    UserPassword = ReplaceBadChar(Trim(Request.Form("UserPassword")))
    Question = ReplaceBadChar(Trim(Request.Form("Question")))
    Answer = ReplaceBadChar(Trim(Request.Form("Answer")))
    Email = ReplaceBadChar(Trim(Request.Form("Email")))
    UserFace = Trim(Request.Form("UserFace"))
    FaceWidth = PE_CLng(Trim(Request.Form("FaceWidth")))
    FaceHeight = PE_CLng(Trim(Request.Form("FaceHeight")))
    Sign = Trim(Request.Form("Sign"))
    Privacy = PE_CLng(Trim(Request.Form("Privacy")))

    If Action = "SaveAdd" Then
        If UserName = "" Then
            FoundErr = True
            ErrMsg = "用户名不能为空！"
        End If
        If UserPassword = "" Then
            FoundErr = True
            ErrMsg = "密码不能为空！"
        End If
        If Answer = "" Then
            FoundErr = True
            ErrMsg = "提示答案不能为空！"
        End If
    End If
    
    If Question = "" Then
        FoundErr = True
        ErrMsg = "提示问题不能为空！"
    End If
    If Email = "" Then
        FoundErr = True
        ErrMsg = "Email不能为空！"
    End If

    If FoundErr Then
        Exit Sub
    End If
    
    '整合操作
    If Action <> "SaveAddUser" Then
        Dim tempUser, tempName
        Set tempUser = Conn.Execute("SELECT UserName FROM PE_User WHERE UserID=" & UserID)
        tempName = tempUser(0)
        tempUser.Close
        Set tempUser = Nothing
    End If
    If Action = "SaveAddUser" Then
        If Not API_RegUser Then Exit Sub
    Else
        If Not API_UpdateUser(tempName) Then Exit Sub
    End If
    '完毕
    Dim sqlUser, rsUser
    If Action = "SaveAddUser" Then
        sqlUser = "SELECT * FROM PE_User Where UserName='" & UserName & "'"
        Set rsUser = Server.CreateObject("adodb.recordset")
        rsUser.Open sqlUser, Conn, 1, 3
        If rsUser.BOF And rsUser.EOF Then
            UserID = GetNewID("PE_User", "UserID")
            ClientID = 0
            ContacterID = 0
            CompanyID = 0
            UserType = 0
            rsUser.addnew
            rsUser("UserID") = UserID
            rsUser("UserName") = UserName
            rsUser("ClientID") = 0
            rsUser("ContacterID") = 0
            rsUser("RegTime") = Now()
            rsUser("JoinTime") = Now()
            rsUser("LoginTimes") = 0
            rsUser("Balance") = 0
            rsUser("UserExp") = 0
            rsUser("PostItems") = 0
            rsUser("PassedItems") = 0
            rsUser("DelItems") = 0
            rsUser("IsLocked") = False
            rsUser("UnsignedItems") = ""
            rsUser("UnreadMsg") = 0
            rsUser("UserPoint") = 0
            rsUser("ValidNum") = 0
            rsUser("ValidUnit") = 1
            rsUser("UserFriendGroup") = "黑名单$我的好友"
            rsUser("BeginTime") = FormatDateTime(Now(), 2)
        Else
            FoundErr = True
            ErrMsg = "该用户名已被他人占用，请输入不同的用户名！"
        End If
    Else
        sqlUser = "SELECT * FROM PE_User Where UserID=" & UserID & ""
        Set rsUser = Server.CreateObject("adodb.recordset")
        rsUser.Open sqlUser, Conn, 1, 3
        If rsUser.BOF And rsUser.EOF Then
            FoundErr = True
            ErrMsg = "找不到指定的会员！"
        Else
            UserName = rsUser("UserName")
            ClientID = rsUser("ClientID")
            ContacterID = rsUser("ContacterID")
            CompanyID = rsUser("CompanyID")
            UserType = rsUser("UserType")
        End If
    End If
    If FoundErr Then
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

    If UserPassword <> "" Then
        rsUser("UserPassword") = MD5(UserPassword, 16)
    End If
    rsUser("Question") = Question
    If Answer <> "" Then
        rsUser("Answer") = MD5(Answer, 16)
    End If
    rsUser("GroupID") = GroupID
    rsUser("Email") = Email
    rsUser("UserFace") = UserFace
    rsUser("FaceWidth") = FaceWidth
    rsUser("FaceHeight") = FaceHeight
    rsUser("Sign") = Sign
    rsUser("Privacy") = Privacy
    
    rsUser.Update
    rsUser.Close
    Set rsUser = Nothing

    Dim TrueName, Title
    Dim Country1, Province1, City1, Address1, ZipCode1
    Dim Mobile, OfficePhone, Homephone, Fax1, PHS
    Dim HomePage, Email1, QQ, ICQ, MSN, Yahoo, UC, Aim
    Dim IDCard, Birthday, NativePlace, Nation, Sex, Marriage, Income, Education, GraduateFrom
    Dim InterestsOfLife, InterestsOfCulture, InterestsOfAmusement, InterestsOfSport, InterestsOfOther
    Dim Company, Department, Position, Operation, CompanyAddress

    Dim Company2, Country2, Province2, City2, Address2, ZipCode2, HomePage2, Phone, Fax2
    Dim BankOfDeposit, BankAccount, TaxNum, StatusInField, CompanySize, BusinessScope, AnnualSales, ManagementForms, RegisteredCapital
    Dim CompanyIntro, CompamyPic
   
    TrueName = Trim(Request.Form("TrueName"))
    Title = Trim(Request.Form("Title"))
    Country1 = Trim(Request.Form("Country1"))
    Province1 = Trim(Request.Form("Province1"))
    City1 = Trim(Request.Form("City1"))
    Address1 = Trim(Request.Form("Address1"))
    ZipCode1 = Trim(Request.Form("ZipCode1"))
    
    Mobile = Trim(Request.Form("Mobile"))
    OfficePhone = Trim(Request.Form("OfficePhone"))
    Homephone = Trim(Request.Form("HomePhone"))
    PHS = Trim(Request.Form("PHS"))
    Fax1 = Trim(Request.Form("Fax1"))

    HomePage = Trim(Request.Form("Homepage1"))
    Email1 = Trim(Request.Form("Email1"))
    QQ = Trim(Request.Form("QQ"))
    MSN = Trim(Request.Form("MSN"))
    ICQ = Trim(Request.Form("ICQ"))
    Yahoo = Trim(Request.Form("Yahoo"))
    UC = Trim(Request.Form("UC"))
    Aim = Trim(Request.Form("Aim"))
    
    IDCard = Trim(Request.Form("IDCard"))
    Birthday = PE_CDate(Trim(Request.Form("Birthday")))
    NativePlace = Trim(Request.Form("NativePlace"))
    Nation = Trim(Request.Form("Nation"))
    Sex = PE_CLng(Trim(Request.Form("Sex")))
    Marriage = PE_CLng(Trim(Request.Form("Marriage")))
    Education = PE_CLng(Trim(Request.Form("Education")))
    GraduateFrom = Trim(Request.Form("GraduateFrom"))
    Income = PE_CLng(Trim(Request.Form("Income")))
    InterestsOfLife = Trim(Request.Form("InterestsOfLife"))
    InterestsOfCulture = Trim(Request.Form("InterestsOfCulture"))
    InterestsOfAmusement = Trim(Request.Form("InterestsOfAmusement"))
    InterestsOfSport = Trim(Request.Form("InterestsOfSport"))
    InterestsOfOther = Trim(Request.Form("InterestsOfOther"))

    Company = Trim(Request.Form("Company"))
    Department = Trim(Request.Form("Department"))
    Position = Trim(Request.Form("Position"))
    Operation = Trim(Request.Form("Operation"))
    CompanyAddress = Trim(Request.Form("CompanyAddress"))

    Company2 = Trim(Request.Form("Company2"))
    Country2 = Trim(Request.Form("Country2"))
    Province2 = Trim(Request.Form("Province2"))
    City2 = Trim(Request.Form("City2"))
    Address2 = Trim(Request.Form("Address2"))
    ZipCode2 = Trim(Request.Form("ZipCode2"))
    Phone = Trim(Request.Form("Phone"))
    Fax2 = Trim(Request.Form("Fax2"))
    HomePage2 = Trim(Request.Form("Homepage2"))
    BankOfDeposit = Trim(Request.Form("BankOfDeposit"))
    BankAccount = Trim(Request.Form("BankAccount"))
    TaxNum = Trim(Request.Form("TaxNum"))
    StatusInField = PE_CLng(Trim(Request.Form("StatusInField")))
    CompanySize = PE_CLng(Trim(Request.Form("CompanySize")))
    BusinessScope = Trim(Request.Form("BusinessScope"))
    AnnualSales = Trim(Request.Form("AnnualSales"))
    ManagementForms = PE_CLng(Trim(Request.Form("ManagementForms")))
    RegisteredCapital = Trim(Request.Form("RegisteredCapital"))
    CompanyIntro = Trim(Request.Form("CompanyIntro"))
    CompamyPic = Trim(Request.Form("CompamyPic"))

    If FoundInArr(RegFields_MustFill, "TrueName", ",") = True And TrueName = "" Then
        FoundErr = True
        ErrMsg = "真实姓名不能为空！"
        Exit Sub
    End If
    Dim sqlContacter, rsContacter
    Set rsContacter = Server.CreateObject("adodb.recordset")
    sqlContacter = "select * From PE_Contacter where ContacterID=" & ContacterID & ""
    rsContacter.Open sqlContacter, Conn, 1, 3
    If rsContacter.BOF And rsContacter.EOF Then
        ContacterID = GetNewID("PE_Contacter", "ContacterID")
        Conn.Execute ("update PE_User set ContacterID=" & ContacterID & " where UserID=" & UserID & "")
        rsContacter.addnew
        rsContacter("ContacterID") = ContacterID
        rsContacter("ClientID") = ClientID
        rsContacter("ParentID") = 0
        rsContacter("Family") = ""
        rsContacter("CreateTime") = Now()
        rsContacter("Owner") = ""
    End If
    rsContacter("UserType") = UserType
    rsContacter("TrueName") = TrueName
    rsContacter("Country") = Country1
    rsContacter("Province") = Province1
    rsContacter("City") = City1
    rsContacter("ZipCode") = ZipCode1
    rsContacter("Address") = Address1
    rsContacter("Mobile") = Mobile
    rsContacter("OfficePhone") = OfficePhone
    rsContacter("HomePhone") = Homephone
    rsContacter("PHS") = PHS
    rsContacter("Fax") = Fax1
    rsContacter("Homepage") = HomePage
    rsContacter("Email") = Email1
    rsContacter("QQ") = QQ
    rsContacter("MSN") = MSN
    rsContacter("ICQ") = ICQ
    rsContacter("Yahoo") = Yahoo
    rsContacter("UC") = UC
    rsContacter("Aim") = Aim
    rsContacter("Company") = Company
    rsContacter("CompanyAddress") = CompanyAddress
    rsContacter("Department") = Department
    rsContacter("Position") = Position
    rsContacter("Title") = Title
    rsContacter("BirthDay") = Birthday
    rsContacter("IDCard") = IDCard
    rsContacter("NativePlace") = NativePlace
    rsContacter("Nation") = Nation
    rsContacter("Sex") = Sex
    rsContacter("Marriage") = Marriage
    rsContacter("Education") = Education
    rsContacter("GraduateFrom") = GraduateFrom
    rsContacter("InterestsOfLife") = InterestsOfLife
    rsContacter("InterestsOfCulture") = InterestsOfCulture
    rsContacter("InterestsOfAmusement") = InterestsOfAmusement
    rsContacter("InterestsOfSport") = InterestsOfSport
    rsContacter("InterestsOfOther") = InterestsOfOther
    rsContacter("Income") = Income
    rsContacter("UpdateTime") = Now()
    rsContacter("Operation") = Operation
    rsContacter.Update
    rsContacter.Close
    Set rsContacter = Nothing
    
    If UserType = 1 Then
        If Company2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入单位名称！</li>"
        End If
        If Address2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入单位的联系地址！</li>"
        End If
        If ZipCode2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入单位的邮政编码！</li>"
        End If
        If FoundErr = True Then
            Exit Sub
        End If

        Dim sqlCompany, rsCompany
        Set rsCompany = Server.CreateObject("adodb.recordset")
        sqlCompany = "select * From PE_Company where CompanyID=" & CompanyID & ""
        rsCompany.Open sqlCompany, Conn, 1, 3
        If rsCompany.BOF And rsCompany.EOF Then
            CompanyID = GetNewID("PE_Company", "CompanyID")
            Conn.Execute ("update PE_User set CompanyID=" & CompanyID & " where UserID=" & UserID & "")
            rsCompany.addnew
            rsCompany("CompanyID") = CompanyID
            rsCompany("ClientID") = 0
        End If
        rsCompany("CompanyName") = Company2
        rsCompany("Country") = Country2
        rsCompany("Province") = Province2
        rsCompany("City") = City2
        rsCompany("Address") = Address2
        rsCompany("ZipCode") = ZipCode2
        rsCompany("Phone") = Phone
        rsCompany("Fax") = Fax2
        rsCompany("HomePage") = HomePage2
        rsCompany("BankOfDeposit") = BankOfDeposit
        rsCompany("BankAccount") = BankAccount
        rsCompany("TaxNum") = TaxNum
        rsCompany("StatusInField") = StatusInField
        rsCompany("CompanySize") = CompanySize
        rsCompany("BusinessScope") = BusinessScope
        rsCompany("AnnualSales") = AnnualSales
        rsCompany("ManagementForms") = ManagementForms
        rsCompany("RegisteredCapital") = RegisteredCapital
        rsCompany("CompanyIntro") = PE_HTMLEncode(CompanyIntro)
        rsCompany("CompamyPic") = PE_HTMLEncode(CompamyPic)
        rsCompany.Update
        rsCompany.Close
        Set rsCompany = Nothing
    End If

    Call WriteSuccessMsg("保存会员信息成功", "Admin_User.asp?Action=Show&UserID=" & UserID)
End Sub

Sub SavePurview()
    If AdminPurview > 1 And arrPurview(2) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim sqlUser, rsUser, UserID, strValue
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定UserID！</li>"
        Exit Sub
    End If

    sqlUser = "SELECT * FROM PE_User Where UserID=" & UserID & ""
    Set rsUser = Server.CreateObject("adodb.recordset")
    rsUser.Open sqlUser, Conn, 1, 3
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = "此用户不存在！"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

    Dim i, SpecialPermission
    SpecialPermission = Trim(Request.Form("SpecialPermission"))
    If SpecialPermission = "1" Then
        SpecialPermission = True
    Else
        SpecialPermission = False
    End If
    rsUser("SpecialPermission") = SpecialPermission
    If SpecialPermission = True Then
        For i = 0 To 40
            strValue = Trim(Request.Form("UserSetting" & i & ""))
            If strValue = "" Or (Not IsNumeric(strValue)) Then
                strValue = "0"
            End If
            If UserSetting = "" Then
                UserSetting = strValue
            Else
                UserSetting = UserSetting & "," & strValue
            End If
        Next
        
        arrClass_Browse = ""
        arrClass_View = ""
        arrClass_Input = ""
        Dim tBrowse, tView, tInput
        Dim rsChannel, ChannelDir
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType<>4 And ModuleType<>5 And Disabled=" & PE_False & " ORDER BY OrderID")
        Do While Not rsChannel.EOF
            ChannelDir = rsChannel(0)
            tBrowse = ReplaceBadChar(Trim(Request.Form("arrClass_Browse_" & ChannelDir)))
            tView = ReplaceBadChar(Trim(Request.Form("arrClass_View_" & ChannelDir)))
            tInput = ReplaceBadChar(Trim(Request.Form("arrClass_Input_" & ChannelDir)))
            If tBrowse = "" And tView = "" And tInput = "" Then
                If arrClass_Browse = "" Then
                    arrClass_Browse = ChannelDir & "none"
                Else
                    arrClass_Browse = arrClass_Browse & "," & ChannelDir & "none"
                End If
                If arrClass_View = "" Then
                    arrClass_View = ChannelDir & "none"
                Else
                    arrClass_View = arrClass_View & "," & ChannelDir & "none"
                End If
                If arrClass_View = "" Then
                    arrClass_View = ChannelDir & "none"
                Else
                    arrClass_View = arrClass_View & "," & ChannelDir & "none"
                End If
           Else
                If tBrowse <> "" Then
                    If arrClass_Browse = "" Then
                        arrClass_Browse = tBrowse
                    Else
                        arrClass_Browse = arrClass_Browse & "," & tBrowse
                    End If
                End If
                If tView <> "" Then
                    If arrClass_View = "" Then
                        arrClass_View = tView
                    Else
                        arrClass_View = arrClass_View & "," & tView
                    End If
                End If
                If tInput <> "" Then
                    If arrClass_Input = "" Then
                        arrClass_Input = tInput
                    Else
                        arrClass_Input = arrClass_Input & "," & tInput
                    End If
                End If
            End If
            rsChannel.MoveNext
        Loop
        Set rsChannel = Nothing
        
        rsUser("UserSetting") = UserSetting

        rsUser("arrClass_Browse") = arrClass_Browse
        rsUser("arrClass_View") = arrClass_View
        rsUser("arrClass_Input") = arrClass_Input
    End If
    rsUser.Update
    rsUser.Close
    Set rsUser = Nothing
    Call CloseConn
    Response.Redirect "Admin_User.asp?Action=Show&UserID=" & UserID
End Sub

Sub SaveRemit()
    If AdminPurview > 1 And arrPurview(6) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, Money, ReceiptDate, Bank, Remark, OrderFormNum, OrderFormID
    UserID = PE_CLng(Trim(Request("UserID")))
    Money = PE_CDbl(Trim(Request("Money")))
    ReceiptDate = Trim(Request("ReceiptDate"))
    Bank = Trim(Request("Bank"))
    Remark = Trim(Request("Remark"))

    OrderFormNum = ReplaceBadChar(Trim(Request("OrderFormNum")))
    If OrderFormNum <> "" Then
    'response.write"1<br>"&OrderFormID
        Dim tOrderFormID
        Set tOrderFormID = Conn.Execute("select OrderFormID from PE_OrderForm where OrderFormNum='" & OrderFormNum & "'")
        If tOrderFormID.BOF And tOrderFormID.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的订单号！</li>"
            'response.write"2<br>"&OrderFormNum
        Else
            OrderFormID = tOrderFormID(0)
        End If
        Set tOrderFormID = Nothing
    Else
        OrderFormID = 0
    End If
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！</li>"
    End If
    If Money <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入汇款金额！</li>"
    End If
    If ReceiptDate = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入汇款日期！</li>"
    Else
        ReceiptDate = CDate(ReceiptDate)
    End If
    
    If FoundErr = True Then Exit Sub
    
    Dim trs, ClientID
    Set trs = Conn.Execute("select UserName,Balance,ClientID from PE_User where UserID=" & UserID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserName = trs(0)
        ClientID = trs(2)
    End If
    Set trs = Nothing
    
    If FoundErr = True Then Exit Sub

    Conn.Execute ("update PE_User set Balance=Balance+" & Money & " where UserID=" & UserID & "")
    
    Dim strMsg
    strMsg = "给会员添加银行汇款记录成功！"
    Call AddBankrollItem(AdminName, UserName, ClientID, Money, 2, Bank, 0, 1, OrderFormID, 0, Remark, ReceiptDate)
    MessageOfAddRemit = Replace(MessageOfAddRemit, "{$Money}", Money)
    MessageOfAddRemit = Replace(MessageOfAddRemit, "{$ReceiptDate}", ReceiptDate)
    MessageOfAddRemit = Replace(MessageOfAddRemit, "{$BankName}", Bank)
    strMsg = strMsg & SendMessageToUser(UserName, MessageOfAddRemit)
    
    If FoundErr = True Then
        Exit Sub
    End If
    Call WriteSuccessMsg(strMsg, "Admin_User.asp?Action=Show&UserID=" & UserID & "")
End Sub

Sub SaveIncome()
    If AdminPurview > 1 And arrPurview(6) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, OrderFormID, OrderFormNum
    Dim Money, Remark
    UserID = PE_CLng(Trim(Request("UserID")))
    Money = PE_CDbl(Trim(Request("Money")))
    Remark = Trim(Request("Remark"))

    OrderFormNum = ReplaceBadChar(Trim(Request("OrderFormNum")))
    If OrderFormNum <> "" Then
    'response.write"1<br>"&OrderFormID
        Dim tOrderFormID
        Set tOrderFormID = Conn.Execute("select OrderFormID from PE_OrderForm where OrderFormNum='" & OrderFormNum & "'")
        If tOrderFormID.BOF And tOrderFormID.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的订单号！</li>"
            'response.write"2<br>"&OrderFormNum
        Else
            OrderFormID = tOrderFormID(0)
        End If
        Set tOrderFormID = Nothing
    Else
        OrderFormID = 0
    End If

    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！"
    End If
    If Money <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入收入金额！</li>"
    End If
    If Remark = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入收入项目！</li>"
    End If
    
    If FoundErr = True Then Exit Sub
    
    Dim trs, ClientID
    Set trs = Conn.Execute("select UserName,Balance,ClientID from PE_User where UserID=" & UserID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserName = trs(0)
        ClientID = trs(2)
    End If
    Set trs = Nothing
    
    If FoundErr = True Then Exit Sub
        
    '给会员的资金余额中加上相应款项
    Conn.Execute ("update PE_User set Balance=Balance+" & Money & " where UserID=" & UserID & "")
                
    '向资金明细表中添加收入记录
    Call AddBankrollItem(AdminName, UserName, ClientID, Money, 4, "", 0, 1, OrderFormID, 0, Remark, Now())

    Dim strMsg
    strMsg = "给会员添加收入成功！"
    MessageOfAddIncome = Replace(MessageOfAddIncome, "{$Money}", Money)
    MessageOfAddIncome = Replace(MessageOfAddIncome, "{$Reason}", Remark)
    strMsg = strMsg & SendMessageToUser(UserName, MessageOfAddIncome)
    
    If FoundErr = True Then
        Exit Sub
    End If
    Call WriteSuccessMsg(strMsg, "Admin_User.asp?Action=Show&UserID=" & UserID & "")
End Sub

Sub SavePayment()
    If AdminPurview > 1 And arrPurview(6) = False And arrPurview(15) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, UserName, Balance, OrderFormID, OrderFormNum
    Dim rsOrderForm, sqlOrderForm
    Dim Money, Remark
    UserID = PE_CLng(Trim(Request("UserID")))
    OrderFormID = PE_CLng(Trim(Request("OrderFormID")))
    OrderFormNum = ReplaceBadChar(Trim(Request("OrderFormNum")))
    If OrderFormID = 0 And OrderFormNum <> "" Then
    'response.write"1<br>"&OrderFormID
        Dim tOrderFormID
        Set tOrderFormID = Conn.Execute("select OrderFormID from PE_OrderForm where OrderFormNum='" & OrderFormNum & "'")
        If tOrderFormID.BOF And tOrderFormID.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的订单号！</li>"
            'response.write"2<br>"&OrderFormNum
        Else
            OrderFormID = tOrderFormID(0)
        End If
        Set tOrderFormID = Nothing
    End If
    Money = PE_CDbl(Trim(Request("Money")))
    Remark = Trim(Request("Remark"))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定会员ID！"
    Else
        UserID = CLng(UserID)
    End If
    If Money <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入支付金额！</li>"
    End If
    If Remark = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入支出原因！</li>"
    End If
    
    If FoundErr = True Then Exit Sub
    'response.write"1<br>"&OrderFormID &"2<br>"&OrderFormNum
    'response.end
    Dim trs, ClientID
    Set trs = Conn.Execute("select UserName,Balance,ClientID from PE_User where UserID=" & UserID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
    Else
        UserName = trs(0)
        Balance = trs(1)
        ClientID = trs(2)
    End If
    Set trs = Nothing
    If Balance < Money Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>会员资金余额小于支出金额！</li>"
    End If
    
    If FoundErr = True Then Exit Sub

    
    If OrderFormID > 0 Then
        Set rsOrderForm = Server.CreateObject("Adodb.RecordSet")
        sqlOrderForm = "select OrderFormNum,Status,MoneyTotal,MoneyReceipt from PE_OrderForm where OrderFormID=" & OrderFormID
        rsOrderForm.Open sqlOrderForm, Conn, 1, 3
        If rsOrderForm.BOF And rsOrderForm.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的订单！</li>"
            rsOrderForm.Close
            Set rsOrderForm = Nothing
            Exit Sub
        End If
        If rsOrderForm("MoneyTotal") <= rsOrderForm("MoneyReceipt") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此订单已经付清，无需再支付！</li>"
            rsOrderForm.Close
            Set rsOrderForm = Nothing
            Exit Sub
        End If
        rsOrderForm("MoneyReceipt") = rsOrderForm("MoneyReceipt") + Money
        If rsOrderForm("Status") = 2 Then
            rsOrderForm("Status") = 3
        End If
        rsOrderForm.Update
        rsOrderForm.Close
        Set rsOrderForm = Nothing
    End If

    '给会员的资金余额中减去相应款项
    Conn.Execute ("update PE_User set Balance=Balance-" & Money & " where UserID=" & UserID & "")
                
    '向资金明细表中添加支出记录
    Call AddBankrollItem(AdminName, UserName, ClientID, Money, 4, "", 0, 2, OrderFormID, 0, Remark, Now())

    Dim strMsg
    If OrderFormID > 0 Then
        strMsg = "支付订单金额成功！"
    Else
        strMsg = "给会员扣款成功！"
    End If
    MessageOfAddPayment = Replace(MessageOfAddPayment, "{$Money}", Money)
    MessageOfAddPayment = Replace(MessageOfAddPayment, "{$Reason}", Remark)
    strMsg = strMsg & SendMessageToUser(UserName, MessageOfAddPayment)
    
    If FoundErr = True Then
        Exit Sub
    End If
    If OrderFormID > 0 Then
        Call WriteSuccessMsg(strMsg, "Admin_Order.asp?Action=ShowOrder&OrderFormID=" & OrderFormID)
    Else
        Call WriteSuccessMsg(strMsg, "Admin_User.asp?Action=Show&UserID=" & UserID)
    End If
End Sub

Sub LockUser()
    If AdminPurview > 1 And arrPurview(3) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim UserID, sql
    UserID = Trim(Request("UserID"))
    If IsValidID(UserID) = False Then
        UserID = ""
    End If
    If UserID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要锁定的会员</li>"
        Exit Sub
    End If
    If Action = "Lock" Then
        sql = "Update PE_User set IsLocked=" & PE_True & " where UserID in (" & UserID & ")"
    Else
        sql = "Update PE_User set IsLocked=" & PE_False & " where UserID in (" & UserID & ")"
    End If
    Conn.Execute sql

    Call CloseConn
    If InStr(UserID, ",") > 0 Then
        Response.Redirect "Admin_User.asp"
    Else
        Response.Redirect "Admin_User.asp?Action=Show&UserID=" & UserID
    End If
End Sub


Sub DoUpdate()
    If AdminPurview > 1 And arrPurview(1) = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Exit Sub
    End If
    Dim BeginID, EndID, sqlUser, rsUser, trs
    BeginID = PE_CLng(Trim(Request("BeginID")))
    EndID = PE_CLng(Trim(Request("EndID")))
    If BeginID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定开始ID</li>"
    End If
    If EndID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定结束ID</li>"
    End If
    
    If FoundErr = True Then Exit Sub
    Call UpdateUserData(1, "", BeginID, EndID)
    Call WriteSuccessMsg("已经成功将会员数据进行了更新！", ComeUrl)
End Sub

Sub SaveExchange()
    If AdminPurview > 1 Then
        If (arrPurview(7) = False And InStr(Action, "Point") = 0) Or (arrPurview(8) = False And InStr(Action, "Valid") = 0) Then
            Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
            Call WriteEntry(6, AdminName, "越权操作")
            Exit Sub
        End If
    End If

    Dim UserID, Point, ValidNum, ValidUnit, Money, UseExp, ValidType, Reason
    Dim rsUser, sqlUser, strMsg, iTemp
    UserID = PE_CLng(Trim(Request("UserID")))
    If UserID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定UserID！</li>"
        Exit Sub
    End If
    ValidNum = Abs(PE_CLng(Trim(Request("ValidNum"))))
    ValidUnit = Abs(PE_CLng(Trim(Request("ValidUnit"))))
    Money = Abs(PE_CDbl(Trim(Request("Money"))))
    UseExp = Abs(PE_CLng(Trim(Request("UseExp"))))
    Point = Abs(PE_CLng(Trim(Request("Point"))))
    ValidType = Abs(PE_CLng(Trim(Request("ValidType"))))
    Reason = ReplaceBadChar(Trim(Request("Reason")))
    
    
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 3
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If
    If Action = "DoExchangePoint" Or Action = "DoExchangeValid" Then
        If Money = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入要减去的资金数！</li>"
        Else
            If Money > rsUser("Balance") Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>输入的资金数大于会员的资金余额！</li>"
            End If
        End If
    Else  '奖励或扣除，需要输入原因
        If Reason = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入原因！</li>"
        End If
    End If
    If InStr(Action, "Point") > 0 Then
        If Point = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入要追加/扣除的会员" & PointName & "数！</li>"
        End If
    End If
    If FoundErr = True Then
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If
    
    Select Case Action
    Case "DoAddPoint"
        rsUser("UserPoint") = rsUser("UserPoint") + Point
        rsUser.Update
        strMsg = "奖励" & PointName & "成功！"
        Call AddConsumeLog(AdminName, 0, rsUser("UserName"), 0, Point, 1, Reason)
        
        MessageOfAddPoint = Replace(MessageOfAddPoint, "{$Point}", Point)
        MessageOfAddPoint = Replace(MessageOfAddPoint, "{$Reason}", Reason)
        strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfAddPoint)
    Case "DoMinusPoint"
        If rsUser("UserPoint") > Point Then
            iTemp = Point
            rsUser("UserPoint") = rsUser("UserPoint") - Point
        Else
            iTemp = rsUser("UserPoint")
            rsUser("UserPoint") = 0
        End If
        rsUser.Update
        strMsg = "扣除" & PointName & "成功！"
        Call AddConsumeLog(AdminName, 0, rsUser("UserName"), 0, iTemp, 2, Reason)
        
        MessageOfMinusPoint = Replace(MessageOfMinusPoint, "{$Point}", Point)
        MessageOfMinusPoint = Replace(MessageOfMinusPoint, "{$Reason}", Reason)
        strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfMinusPoint)
    Case "DoExchangePoint"
        rsUser("UserPoint") = rsUser("UserPoint") + Point
        rsUser("Balance") = rsUser("Balance") - Money
        rsUser.Update
        strMsg = "兑换" & PointName & "成功！"
        
        Call AddBankrollItem(AdminName, rsUser("UserName"), rsUser("ClientID"), Money, 4, "", 0, 2, 0, 0, "用于兑换 " & Point & " " & PointUnit & "" & PointName, Now())
        Call AddConsumeLog(AdminName, 0, rsUser("UserName"), 0, Point, 1, "将 " & Money & "元资金兑换成 " & Point & " " & PointUnit & "" & PointName & "")
        
        MessageOfExchangePoint = Replace(MessageOfExchangePoint, "{$Money}", Money)
        MessageOfExchangePoint = Replace(MessageOfExchangePoint, "{$Point}", Point)
        strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfExchangePoint)

    Case "DoExchangeValid", "DoAddValid"
        If rsUser("ValidNum") = -1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & rsUser("UserName") & "的有效期是“无限期”，无需追加有效期。</li>"
        Else
            If ValidType = 2 Then
                rsUser("ValidNum") = -1
                rsUser.Update
                If Action = "DoExchangeValid" Then
                    rsUser("Balance") = rsUser("Balance") - Money
                    rsUser.Update
                    Call AddBankrollItem(AdminName, rsUser("UserName"), rsUser("ClientID"), Money, 4, "", 0, 2, 0, 0, "用于将有效期变为“无限期”", Now())
                    Call AddRechargeLog(AdminName, rsUser("UserName"), 0, 0, 0, "有效期设为“无限期”")
                    strMsg = "兑换有效期成功！"
                    MessageOfExchangeValid = Replace(MessageOfExchangeValid, "{$Money}", Money)
                    MessageOfExchangeValid = Replace(MessageOfExchangeValid, "{$Valid}", "无限期")
                    strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfExchangeValid)
                Else
                    Call AddRechargeLog(AdminName, rsUser("UserName"), 0, 0, 0, "有效期设为“无限期”")
                    strMsg = "奖励有效期成功！"
                    MessageOfAddValid = Replace(MessageOfAddValid, "{$Valid}", "无限期")
                    MessageOfAddValid = Replace(MessageOfAddValid, "{$Reason}", Reason)
                    strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfAddValid)
                End If
            Else
                If ValidNum = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>请指定要追加的有效期！</li>"
                Else
                    ValidDays = ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime"))
                    If ValidDays > 0 Then
                        If rsUser("ValidUnit") = ValidUnit Then
                            rsUser("ValidNum") = rsUser("ValidNum") + ValidNum
                        ElseIf rsUser("ValidUnit") < ValidUnit Then
                            If rsUser("ValidUnit") = 1 Then
                                If ValidUnit = 2 Then
                                    rsUser("ValidNum") = rsUser("ValidNum") + ValidNum * 30
                                Else
                                    rsUser("ValidNum") = rsUser("ValidNum") + ValidNum * 365
                                End If
                            Else
                                rsUser("ValidNum") = rsUser("ValidNum") + ValidNum * 12
                            End If
                        Else
                            If ValidUnit = 1 Then
                                If rsUser("ValidUnit") = 2 Then
                                    rsUser("ValidNum") = ValidNum + rsUser("ValidNum") * 30
                                Else
                                    rsUser("ValidNum") = ValidNum + rsUser("ValidNum") * 365
                                End If
                            Else
                                rsUser("ValidNum") = ValidNum + rsUser("ValidNum") * 12
                            End If
                            rsUser("ValidUnit") = ValidUnit
                            Call AddRechargeLog(AdminName, rsUser("UserName"), 0, 0, 0, "添加有效期时更改有效期计费单位")
                        End If
                    Else
                        rsUser("BeginTime") = FormatDateTime(Now(), 2)
                        rsUser("ValidNum") = ValidNum
                        rsUser("ValidUnit") = ValidUnit
                        Call AddRechargeLog(AdminName, rsUser("UserName"), 0, 0, 0, "添加有效期时将原来过期的有效期重新计算")
                    End If
                    rsUser.Update
                    If Action = "DoExchangeValid" Then
                        rsUser("Balance") = rsUser("Balance") - Money
                        rsUser.Update
                        Call AddBankrollItem(AdminName, rsUser("UserName"), rsUser("ClientID"), Money, 4, "", 0, 2, 0, 0, "用于增加 " & ValidNum & arrCardUnit(ValidUnit) & " 有效期", Now())
                        Call AddRechargeLog(AdminName, rsUser("UserName"), ValidNum, ValidUnit, 1, "将 " & Money & "元资金兑换成有效期")
                        strMsg = "兑换有效期成功！"
                        MessageOfExchangeValid = Replace(MessageOfExchangeValid, "{$Money}", Money)
                        MessageOfExchangeValid = Replace(MessageOfExchangeValid, "{$Valid}", ValidNum & arrCardUnit(ValidUnit))
                        strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfExchangeValid)
                    Else
                        Call AddRechargeLog(AdminName, rsUser("UserName"), ValidNum, ValidUnit, 1, Reason)
                        strMsg = "奖励有效期成功！"
                        MessageOfAddValid = Replace(MessageOfAddValid, "{$Valid}", ValidNum & arrCardUnit(ValidUnit))
                        MessageOfAddValid = Replace(MessageOfAddValid, "{$Reason}", Reason)
                        strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfAddValid)
                    End If
                End If
            End If
        End If
    Case "DoMinusValid"
        If ValidType = 2 Then
            rsUser("BeginTime") = FormatDateTime(Now(), 2)
            rsUser("ValidNum") = 0
            rsUser("ValidUnit") = 1
            rsUser.Update
            Call AddRechargeLog(AdminName, rsUser("UserName"), 0, 0, 2, "有效期归零，原因：" & Reason & "")
            strMsg = "有效期归零成功！"
            MessageOfMinusValid = Replace(MessageOfMinusValid, "{$Valid}", "归零")
            MessageOfMinusValid = Replace(MessageOfMinusValid, "{$Reason}", Reason)
            strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfMinusValid)
        Else
            If rsUser("ValidNum") = -1 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>" & rsUser("UserName") & "的有效期是“无限期”，只能执行“有效期归零”的操作。</li>"
            Else
                If ValidNum = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>请指定要扣除的有效期！</li>"
                Else
                    ValidDays = ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime"))
                    If ValidDays > 0 Then
                        If rsUser("ValidUnit") = ValidUnit Then
                            rsUser("ValidNum") = rsUser("ValidNum") - ValidNum
                        ElseIf rsUser("ValidUnit") < ValidUnit Then
                            If rsUser("ValidUnit") = 1 Then
                                If ValidUnit = 2 Then
                                    rsUser("ValidNum") = rsUser("ValidNum") - ValidNum * 30
                                Else
                                    rsUser("ValidNum") = rsUser("ValidNum") - ValidNum * 365
                                End If
                            Else
                                rsUser("ValidNum") = rsUser("ValidNum") - ValidNum * 12
                            End If
                        Else
                            If ValidUnit = 1 Then
                                rsUser("ValidNum") = ValidDays - ValidNum
                            Else
                                rsUser("ValidNum") = rsUser("ValidNum") * 12 - ValidNum
                            End If
                            rsUser("ValidUnit") = ValidUnit
                            Call AddRechargeLog(AdminName, rsUser("UserName"), 0, 0, 0, "扣除有效期时更改有效期计费单位。")
                        End If
                        rsUser.Update
                        If rsUser("ValidNum") < 0 Then
                            rsUser("ValidNum") = 0
                            rsUser.Update
                        End If
                        Call AddRechargeLog(AdminName, rsUser("UserName"), ValidNum, ValidUnit, 2, Reason)
                        strMsg = "扣除有效期成功！"
                        MessageOfMinusValid = Replace(MessageOfMinusValid, "{$Valid}", ValidNum & arrCardUnit(ValidUnit))
                        MessageOfMinusValid = Replace(MessageOfMinusValid, "{$Reason}", Reason)
                        strMsg = strMsg & SendMessageToUser(rsUser("UserName"), MessageOfMinusValid)
                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>有效期已经过期</li>"
                    End If
                End If
            End If
        End If
    End Select
    

    rsUser.Close
    Set rsUser = Nothing
    
    If FoundErr = True Then
        Exit Sub
    End If
    Call WriteSuccessMsg(strMsg, "Admin_User.asp?Action=Show&UserID=" & UserID & "")
    
End Sub

Function SendMessageToUser(UserName, Content)
    If Content = "" Then
        Exit Function
    End If
    
    Dim SendSMSToUser
    SendSMSToUser = Trim(Request("SendSMSToUser"))
    If SendSMSToUser <> "Yes" Then
        Exit Function
    End If

    Dim strContent, strMsg
    Dim rsUser
    Dim strResult
    Set rsUser = Conn.Execute("select U.UserID,U.Balance,U.UserPoint,U.BeginTime,U.ValidNum,U.ValidUnit,C.Mobile,C.PHS from PE_User U inner join PE_Contacter C on U.ContacterID=C.ContacterID where U.UserName='" & UserName & "'")
    If rsUser.BOF And rsUser.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>会员没有填写联系方式，所以没有发送手机短信！</li>"
    Else
        Dim SendNum
        SendNum = rsUser("Mobile") & ""
        If SendNum = "" Then
            SendNum = rsUser("PHS") & ""
        End If
        If SendNum = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>会员没有填写手机号，所以没有发送手机短信！</li>"
        Else
            strContent = Replace(Content, "{$SiteName}", SiteName)
            strContent = Replace(strContent, "{$UserName}", UserName)
            strContent = Replace(strContent, "{$Balance}", rsUser("Balance"))
            strContent = Replace(strContent, "{$UserPoint}", rsUser("UserPoint"))
            strContent = Replace(strContent, "{$ValidDays}", ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime")))

            strResult = PostSMS(SendNum, strContent, "管理员ID：" & AdminID)
            If strResult = "短信已经提交到发送队列" Then
                strMsg = strMsg & "<br><br>已经向会员发送了一条手机短信通知他！"
            Else
                strMsg = strMsg & "<br><br>向会员发送了手机短信失败，失败原因：" & strResult & ""
            End If
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing
    SendMessageToUser = strMsg
End Function


Sub MemberManage()
    Dim UserID, MemberID
    UserID = PE_CLng(Trim(Request("UserID")))
    MemberID = PE_CLng(Trim(Request("MemberID")))
    If UserID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定UserID！</li>"
        Exit Sub
    End If
    If MemberID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定MemberID！</li>"
        Exit Sub
    End If
    Select Case Action
    Case "Agree"
        Dim rsUser
        Set rsUser = Conn.Execute("select ClientID from PE_User where UserID=" & UserID & "")
        If rsUser.BOF And rsUser.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的会员！</li>"
        Else
            Conn.Execute ("update PE_User set UserType=3,ClientID=" & rsUser(0) & " where UserType=4 and UserID=" & MemberID & "")
        End If
        rsUser.Close
        Set rsUser = Nothing
    Case "Reject", "RemoveFromCompany"
        Conn.Execute ("update PE_User set UserType=0,CompanyID=0,ClientID=0 where UserID=" & MemberID & "")
    Case "AddToAdmin"
        Conn.Execute ("update PE_User set UserType=2 where UserType>2 and UserID=" & MemberID & "")
    Case "RemoveFromAdmin"
        Conn.Execute ("update PE_User set UserType=3 where UserType=2 and UserID=" & MemberID & "")
    End Select

    Response.Redirect "Admin_User.asp?Action=Show&UserID=" & UserID
End Sub

Function GetUserGroup_Option(CurrentGroupID)
    Dim strGroup, rsGroup
    Set rsGroup = Conn.Execute("select GroupID,GroupName,arrClass_Browse,arrClass_View,arrClass_Input,GroupSetting from PE_UserGroup where GroupID not in (-1) order by GroupType asc,GroupID asc")
    Do While Not rsGroup.EOF
        strGroup = strGroup & "<option value='" & rsGroup(0) & "'"
        If rsGroup(0) = CurrentGroupID Then
            strGroup = strGroup & " selected"
            arrClass_Browse = rsGroup(2)
            arrClass_View = rsGroup(3)
            arrClass_Input = rsGroup(4)
            UserSetting = Split(rsGroup(5) & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        End If
        strGroup = strGroup & ">" & rsGroup(1) & "</option>"
        rsGroup.MoveNext
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
    
    GetUserGroup_Option = strGroup
End Function

'**************************************************
'函数名：UserNamefilter(
'作  用：过滤用户名(增强过滤,用户名现用于建立个人空间目录)
'**************************************************
Function UserNamefilter(strChar)
    If strChar = "" Or IsNull(strChar) Then
        UserNamefilter = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",*,|,"""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    UserNamefilter = tempChar
End Function


'向整合的系统添加用户
'返回值：True=添加成功；False=添加失败
Function API_RegUser()
    If Not API_Enable Then
        API_RegUser = True
        Exit Function
    Else
        API_RegUser = False
    End If
    'On Error Resume Next
    If createXmlHttp And createXmlDom Then
        XMLDOM.Load (Server.MapPath(InstallDir & "API/Request.xml"))
        setXmlNode "username", Trim(Request.Form("UserName"))
        setXmlNode "password", Trim(Request.Form("UserPassword"))
        setXmlNode "email", Trim(Request.Form("Email"))
        setXmlNode "question", Trim(Request.Form("Question"))
        setXmlNode "answer", Trim(Request.Form("Answer"))
        setXmlNode "truename", Trim(Request.Form("TrueName"))
        If PE_CLng(Trim(Request.Form("Sex"))) = 2 Then
            setXmlNode "gender", "1"
        ElseIf PE_CLng(Trim(Request.Form("Sex"))) = 1 Then
            setXmlNode "gender", "0"
        Else
            setXmlNode "gender", "2"
        End If
        setXmlNode "birthday", PE_CDate(Trim(Request.Form("Birthday")))
        setXmlNode "qq", Trim(Request.Form("QQ"))
        setXmlNode "msn", Trim(Request.Form("MSN"))
        setXmlNode "mobile", Trim(Request.Form("Mobile"))
        setXmlNode "telephone", Trim(Request.Form("OfficePhone"))
        setXmlNode "address", Trim(Request.Form("Address1"))
        setXmlNode "zipcode", Trim(Request.Form("ZipCode1"))
        setXmlNode "homepage", Trim(Request.Form("Homepage1"))
        If PE_CLng(Trim(Request.Form("GroupID"))) = 8 Then
            setXmlNode "userstatus", "4"
        End If
        setXmlNode "syskey", LCase(MD5(getXmlNode("username") & API_Key, 16))
        setXmlNode "action", "reguser"
        xmlHttp.setTimeouts API_Timeout, API_Timeout, API_Timeout * 6, API_Timeout * 6
        Dim intIndex, arrAPIUrls
        arrAPIUrls = Split(API_Urls, "|")
        For intIndex = 0 To UBound(arrAPIUrls)
            API_RegUser = False
            Dim arrRemoteSys
            arrRemoteSys = Split(arrAPIUrls(intIndex), "@@")
            xmlHttp.Open "POST", arrRemoteSys(1), False
            xmlHttp.setRequestHeader "Content-Type", "text/xml"
            xmlHttp.Send XMLDOM
            If Err Then
                Err.Clear
                FoundErr = True
                ErrMsg = "向[" & arrRemoteSys(0) & "]发送请求失败，请检查[" & arrRemoteSys(0) & "]的接口配置！"
                Exit Function
            End If
            If xmlHttp.readyState = 4 And xmlHttp.Status = 200 Then
                Dim vXmlDoc
                Set vXmlDoc = Server.CreateObject("MSXML.FreeThreadedDOMDocument")
                vXmlDoc.Async = False
                vXmlDoc.Load (xmlHttp.responseXML)
                If Err Then
                    Err.Clear
                    FoundErr = True
                    ErrMsg = "加载[" & arrRemoteSys(0) & "]返回的XML数据错误！"
                    Exit Function
                End If
                If vXmlDoc.parseError.errorCode <> 0 Then
                    FoundErr = True
                    ErrMsg = "校验[" & arrRemoteSys(0) & "]返回的XML数据失败！可能[" & arrRemoteSys(0) & "]未能返回符合规范的数据。"
                    Exit Function
                End If
                If vXmlDoc.selectSingleNode("//status").text <> "0" Then
                    FoundErr = True
                    ErrMsg = "[" & arrRemoteSys(0) & "]返回提示：" & vXmlDoc.documentElement.selectSingleNode("//message").text
                Else
                    API_RegUser = True
                End If
            Else
                FoundErr = True
                ErrMsg = "无法与[" & arrRemoteSys(0) & "]完成同步，通常是由于网络超时或[" & arrRemoteSys(0) & "]接口程序出错导致！"
                Exit Function
            End If
        Next
        Set xmlHttp = Nothing
    Else
        FoundErr = True
        ErrMsg = "服务器不支持MSXML组件，请禁用整合接口！"
        Exit Function
    End If
End Function

'更新整合系统的用户资料
'返回值：True=更新成功；False=更新失败
Function API_UpdateUser(vUserName)
    If Not API_Enable Then
        API_UpdateUser = True
        Exit Function
    Else
        API_UpdateUser = False
    End If
    'On Error Resume Next
    If createXmlHttp And createXmlDom Then
        XMLDOM.Load (Server.MapPath(InstallDir & "API/Request.xml"))
        setXmlNode "username", vUserName
        setXmlNode "password", Trim(Request.Form("UserPassword"))
        setXmlNode "email", Trim(Request.Form("Email"))
        setXmlNode "question", Trim(Request.Form("Question"))
        setXmlNode "answer", Trim(Request.Form("Answer"))
        setXmlNode "truename", Trim(Request.Form("TrueName"))
        If PE_CLng(Trim(Request.Form("Sex"))) = 2 Then
            setXmlNode "gender", "1"
        ElseIf PE_CLng(Trim(Request.Form("Sex"))) = 1 Then
            setXmlNode "gender", "0"
        Else
            setXmlNode "gender", "2"
        End If
        setXmlNode "birthday", PE_CDate(Trim(Request.Form("Birthday")))
        setXmlNode "qq", Trim(Request.Form("QQ"))
        setXmlNode "msn", Trim(Request.Form("MSN"))
        setXmlNode "mobile", Trim(Request.Form("Mobile"))
        setXmlNode "telephone", Trim(Request.Form("OfficePhone"))
        setXmlNode "address", Trim(Request.Form("Address1"))
        setXmlNode "zipcode", Trim(Request.Form("ZipCode1"))
        setXmlNode "homepage", Trim(Request.Form("Homepage1"))
        If PE_CLng(Trim(Request.Form("GroupID"))) = 8 Then
            setXmlNode "userstatus", "4"
        End If
        setXmlNode "syskey", LCase(MD5(getXmlNode("username") & API_Key, 16))
        setXmlNode "action", "update"
        xmlHttp.setTimeouts API_Timeout, API_Timeout, API_Timeout * 6, API_Timeout * 6
        Dim intIndex, arrAPIUrls
        arrAPIUrls = Split(API_Urls, "|")
        For intIndex = 0 To UBound(arrAPIUrls)
            API_UpdateUser = False
            Dim arrRemoteSys
            arrRemoteSys = Split(arrAPIUrls(intIndex), "@@")
            xmlHttp.Open "POST", arrRemoteSys(1), False
            xmlHttp.setRequestHeader "Content-Type", "text/xml"
            xmlHttp.Send XMLDOM
            If Err Then
                Err.Clear
                FoundErr = True
                ErrMsg = "向[" & arrRemoteSys(0) & "]发送请求失败，请检查[" & arrRemoteSys(0) & "]的接口配置！"
                Exit Function
            End If
            If xmlHttp.readyState = 4 And xmlHttp.Status = 200 Then
                Dim vXmlDoc
                Set vXmlDoc = Server.CreateObject("MSXML.FreeThreadedDOMDocument")
                vXmlDoc.Async = False
                vXmlDoc.Load (xmlHttp.responseXML)
                If Err Then
                    Err.Clear
                    FoundErr = True
                    ErrMsg = "加载[" & arrRemoteSys(0) & "]返回的XML数据错误！"
                    Exit Function
                End If
                If vXmlDoc.parseError.errorCode <> 0 Then
                    FoundErr = True
                    ErrMsg = "校验[" & arrRemoteSys(0) & "]返回的XML数据失败！可能[" & arrRemoteSys(0) & "]未能返回符合规范的数据。"
                    Exit Function
                End If
                If vXmlDoc.selectSingleNode("//status").text <> "0" Then
                    FoundErr = True
                    ErrMsg = "[" & arrRemoteSys(0) & "]返回提示：" & vXmlDoc.documentElement.selectSingleNode("//message").text
                Else
                    API_UpdateUser = True
                End If
            Else
                FoundErr = True
                ErrMsg = "无法与[" & arrRemoteSys(0) & "]完成同步，通常是由于网络超时或[" & arrRemoteSys(0) & "]接口程序出错导致！"
                Exit Function
            End If
        Next
        Set xmlHttp = Nothing
    Else
        FoundErr = True
        ErrMsg = "服务器不支持MSXML组件，请禁用整合接口！"
        Exit Function
    End If
End Function

'删除整合系统的用户
Function API_DelUser(vDelName)
    If Left(vDelName, 1) = "," Then
        vDelName = Right(vDelName, Len(vDelName) - 1)
    End If
    API_DelUser = False
    'On Error Resume Next
    If createXmlHttp And createXmlDom Then
        XMLDOM.Load (Server.MapPath(InstallDir & "API/Request.xml"))
        setXmlNode "username", vDelName
        setXmlNode "syskey", LCase(MD5(getXmlNode("username") & API_Key, 16))
        setXmlNode "action", "delete"
        xmlHttp.setTimeouts API_Timeout, API_Timeout, API_Timeout * 6, API_Timeout * 6
        Dim intIndex, arrAPIUrls
        arrAPIUrls = Split(API_Urls, "|")
        For intIndex = 0 To UBound(arrAPIUrls)
            Dim arrRemoteSys
            arrRemoteSys = Split(arrAPIUrls(intIndex), "@@")
            xmlHttp.Open "POST", arrRemoteSys(1), False
            xmlHttp.setRequestHeader "Content-Type", "text/xml"
            xmlHttp.Send XMLDOM
            If Err Then
                Err.Clear
                FoundErr = True
                ErrMsg = "向[" & arrRemoteSys(0) & "]发送请求失败，请检查[" & arrRemoteSys(0) & "]的接口配置！并到[" & arrRemoteSys(0) & "]的后台删除相关用户。"
                Exit Function
            End If
            If xmlHttp.readyState = 4 And xmlHttp.Status = 200 Then
                Dim vXmlDoc
                Set vXmlDoc = Server.CreateObject("MSXML.FreeThreadedDOMDocument")
                vXmlDoc.Async = False
                vXmlDoc.Load (xmlHttp.responseXML)
                If Err Then
                    Err.Clear
                    FoundErr = True
                    ErrMsg = "加载[" & arrRemoteSys(0) & "]返回的XML数据错误！请到[" & arrRemoteSys(0) & "]的后台删除相关用户。"
                    Exit Function
                End If
                If vXmlDoc.parseError.errorCode <> 0 Then
                    FoundErr = True
                    ErrMsg = "校验[" & arrRemoteSys(0) & "]返回的XML数据失败！可能[" & arrRemoteSys(0) & "]未能返回符合规范的数据。请到[" & arrRemoteSys(0) & "]的后台删除相关用户。"
                    Exit Function
                End If
                If vXmlDoc.selectSingleNode("//status").text <> "0" Then
                    FoundErr = True
                    ErrMsg = "[" & arrRemoteSys(0) & "]返回提示：" & vXmlDoc.documentElement.selectSingleNode("//message").text & "请到[" & arrRemoteSys(0) & "]的后台删除相关用户。"
                End If
            Else
                FoundErr = True
                ErrMsg = "无法与[" & arrRemoteSys(0) & "]完成同步，通常是由于网络超时或[" & arrRemoteSys(0) & "]接口程序出错导致！请到[" & arrRemoteSys(0) & "]的后台删除相关用户。"
                Exit Function
            End If
        Next
        Set xmlHttp = Nothing
    End If
    If Err Then
        Err.Clear
    End If
End Function

'设置XmlDoc对象中的Node的Text值
Sub setXmlNode(vNodeName, vNodeValue)
    If IsNull(vNodeValue) Or vNodeValue = "" Then Exit Sub
    'On Error Resume Next
    XMLDOM.documentElement.selectSingleNode(vNodeName).text = vNodeValue
End Sub

'从vXmlDoc对象中解析vNodeName的Text值
Function getXmlNode(vNodeName)
    'On Error Resume Next
    getXmlNode = XMLDOM.documentElement.selectSingleNode(vNodeName).text
    If IsNull(getXmlNode) Then getXmlNode = ""
End Function

Function createXmlHttp()
    'On Error Resume Next
    Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    If Err Then
        createXmlHttp = False
    Else
        createXmlHttp = True
    End If
End Function

Function createXmlDom()
    'On Error Resume Next
    Set XMLDOM = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
    If Err Then
        createXmlDom = False
    Else
        createXmlDom = True
    End If
End Function
Function GetGroupName(iGroupID)
    If Not IsNumeric(iGroupID) Then Exit Function
    Dim rsGroup
    Set rsGroup = Conn.Execute("select GroupName,GroupSetting,GroupType from PE_UserGroup where GroupID=" & iGroupID & "")
    If rsGroup.BOF And rsGroup.EOF Then
        GetGroupName = "未知会员组"
    Else
        GetGroupName = rsGroup(0)
        UserSetting = Split(rsGroup(1) & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        iGroupType = rsGroup(2)
    End If
    Set rsGroup = Nothing
End Function
%>
