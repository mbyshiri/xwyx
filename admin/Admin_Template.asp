<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
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

Dim TemplateType, downright, TempType, ProjectName
Dim NavigationCss '导航风格
Dim IsOnlinePayment '解决通用模板公用在线支付问题
Dim TemplateProjectID, i

'检查管理员操作权限
If AdminPurview > 1 Then
    If ChannelID > 0 And ModuleType <> 4 Then
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir)
    Else
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Template")
    End If
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Response.End
    End If
End If

TemplateType = Trim(Request("TemplateType"))

downright = PE_CLng(Trim(Request("downright")))
ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
IsOnlinePayment = PE_CLng(Trim(Request("IsOnlinePayment")))
NavigationCss = "title"

If ProjectName = "" Then
    Dim rs
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rs.BOF And rs.EOF Then
        Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
        Response.End
    Else
        ProjectName = rs("TemplateProjectName")
    End If
    Set rs = Nothing
End If

If TemplateType = "" Then
    TemplateType = 1
Else
    TemplateType = PE_CLng(TemplateType)
End If

TempType = PE_CLng(Trim(Request("TempType")))

Response.Write "<html><head><title>" & ChannelName & "管理----模板管理</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>"
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

If ChannelID = 0 Then
    If TempType = 0 Then
        Call ShowPageTitle(ProjectName & "方案----通用模板管理", 10006)
    ElseIf TempType = 1 Then
        Call ShowPageTitle(ProjectName & "方案----网站会员模板管理", 10006)
    End If
Else
    Call ShowPageTitle(ProjectName & "方案----" & ChannelName & "模板管理", 10006)
End If

Response.Write "      <tr class='tdbg'>"
Response.Write "        <td width='70' height='30'><strong>管理导航：</strong></td><td>"

If TempType = 1 Then
    Response.Write "<a href='Admin_Template.asp?TemplateType=8&TempType=1&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'>"
Else
    Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'>"
End If

Response.Write "模板管理首页</a> | <a href='Admin_Template.asp?ChannelID="
If ChannelID = 0 And IsOnlinePayment = 1 Then
    Response.Write "1000&Action=Add&TemplateType=" & TemplateType & "&ProjectName=" & Server.UrlEncode(ProjectName)
Else
    Response.Write ChannelID & "&Action=Add&TemplateType=" & TemplateType & "&ProjectName=" & Server.UrlEncode(ProjectName)
End If

If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "''>添加模板</a>"
Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Import&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>导入模板</a>"
Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Export&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>导出模板</a>"
Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=ChannelCopyTemplate&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>频道模板复制</a>"
Response.Write " | <a href='Admin_Template.asp?Action=BatchReplace&ChannelID=" & ChannelID & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>批量替换模板代码</a>"
Response.Write " | <a href='Admin_Template.asp?Action=Main&ChannelID=" & ChannelID & "&downright=1&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>模板回收站管理</a>"
Response.Write " | <a href='Admin_Template.asp?Action=BatchDefault&ProjectName=" & Server.UrlEncode(ProjectName)
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>模板默认批量设置</a>"
Response.Write "        </td>"
Response.Write "      </tr>"
Response.Write "    </table>"
Response.Write "    <br>"

If Action = "" Or Action = "SaveAdd" Or Action = "SaveModify" Or Action = "Main" Or Action = "main" Then
    
    Response.Write "<table width='100%' border='0' align='center' "
    If TemplateProjectID <> 0 Then
        Response.Write "cellpadding='2' cellspacing='1' class='border'"
    End If
    Response.Write ">"

    If TemplateProjectID <> 0 Then
        Response.Write "  <tr class='" & NavigationCss & "'>"
        Response.Write "    <td>"
        Response.Write "    | <a href='Admin_Template.asp?Action=Main&ChannelID=0&ProjectName=" & Server.UrlEncode(ProjectName) & "&TempType=0&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' "
        If ChannelID = 0 And TempType = 0 Then
            Response.Write " color='red'"
        End If
         Response.Write ">网站通用模板</FONT></a>"
        Response.Write "    | <a href='Admin_Template.asp?Action=Main&ChannelID=0&ProjectName=" & Server.UrlEncode(ProjectName) & "&TempType=1&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' "
        If ChannelID = 0 And TempType = 1 Then
            Response.Write " color='red'"
        End If
        Response.Write " >会员模板</FONT></a>"
        i = 0
        Set rs = Conn.Execute("SELECT DISTINCT t.ChannelID,c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID where c.Disabled=" & PE_False)
        If rs.BOF And rs.EOF Then
            Response.Write " 没有模板请先到导入模板"
        Else
            Do While Not rs.EOF
                Response.Write "    | <a href='Admin_Template.asp?Action=Main&ChannelID=" & rs("ChannelID") & "&TemplateType=" & TemplateType & "&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TempType=" & TempType & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(rs("ChannelID"), ChannelID) & ">" & rs("ChannelName") & "频道模板</FONT></a>"
                If i > 3 Then
                    Response.Write " | </td><tr class='" & NavigationCss & "'><td>"
                    i = 0
                Else
                    i = i + 1
                End If

                rs.MoveNext
            Loop
            Response.Write " | "
        End If
        rs.Close
        Set rs = Nothing
        Response.Write "    </td>"
        Response.Write "  </tr>"
        NavigationCss = "tdbg"
    End If
    Response.Write "  <tr class='" & NavigationCss & "'>"
    Response.Write "    <td>"
    Response.Write "      <table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "       <tr class='" & NavigationCss & "'><td>"
    
    If ChannelID > 0 Then
        If ModuleType = 4 Then
            Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">留言板模板</FONT></a> | "
            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">发表留言模板</FONT></a> | "
            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">留言回复模板</FONT></a> | "
            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">留言搜索页模板</FONT></a> | "
        Else
            Select Case ModuleType
            Case 6
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">频道首页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=2&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">栏目模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">专题页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">搜索页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=6&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 6) & ">最新" & ChannelShortName & "页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">推荐" & ChannelShortName & "页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=8&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">热门" & ChannelShortName & "页模板</FONT></a>"
                Response.Write "<tr class='" & NavigationCss & "'><td> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=16&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 16) & ">评论" & ChannelShortName & "页模板</FONT></a> | "
            Case 7  '*********************增加房产模块管理********************
                Response.Write "| <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">频道首页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=2&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">栏目模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">推荐页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=8&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">热门页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=30&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 30) & ">出售内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=31&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 31) & ">出租内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=32&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 32) & ">求购内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=33&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 33) & ">求租内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=34&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 34) & ">合租内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">搜索页模板</FONT></a> | "
            Case 8 '招聘模块模板管理
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">频道首页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">职位搜索页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">职位申请页模板</FONT></a> | "
            Case Else
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">频道首页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=2&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">栏目模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">内容页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">专题页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=22&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 22) & ">专题列表页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">搜索页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=6&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 6) & ">最新" & ChannelShortName & "页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">推荐" & ChannelShortName & "页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=8&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">热门" & ChannelShortName & "页模板</FONT></a>"
                Response.Write "<tr class='" & NavigationCss & "'><td> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=16&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 16) & ">评论" & ChannelShortName & "页模板</FONT></a> | "
            End Select
            If ModuleType = 1 Then
                Response.Write "  <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=17&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 17) & ">打印页模板</FONT></a>"
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=20&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 20) & ">告诉好友页模板</FONT></a> | "
            ElseIf ModuleType = 5 Then
                Response.Write "   <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=9&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 9) & ">购物车模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=10&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 10) & ">收银台模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=11&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 11) & ">订单预览页模板</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=12&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 12) & ">订购成功模板</FONT></a> | "
                'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=13&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 13) & ">在线支付第一步模板</FONT></a> | "
                'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=14&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 14) & ">在线支付第二步模板</FONT></a> | "
                'Response.Write "<tr class='" & NavigationCss & "'><td> | "
                'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=15&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 15) & ">在线支付第三步模板</FONT></a>"
                Response.Write " <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=19&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 19) & ">特价商品页模板</FONT></a> | "
                Response.Write " <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=21&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 21) & ">商城帮助页模板</FONT></a> | "
            End If
        End If
        Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=0&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 0) & ">所有模板</FONT></a> | </td></tr>"
    Else
        If TempType = 0 Then
            Response.Write " | <a href='Admin_Template.asp?TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">网站首页模板　</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">网站搜索页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">网站公告页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=22&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 22) & ">公告列表页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">友情链接页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=6&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 6) & ">网站调查页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">版权声明页模板</FONT></a>"

            Response.Write "</td></tr>"
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=10&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 10) & ">作者显示页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=11&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 11) & ">作者列表页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=12&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 12) & ">来源显示页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=13&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 13, IsOnlinePayment, 0) & ">来源列表页模板</FONT></a>"
            If ShowAnonymous = True Then			
                Response.Write " | <a href='Admin_Template.asp?TemplateType=103&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 13, IsOnlinePayment, 0) & ">匿名投稿模板</FONT></a>"	
            End If					
            Response.Write "</td></tr>"
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=14&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 14, IsOnlinePayment, 0) & ">厂商显示页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=15&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 15, IsOnlinePayment, 0) & ">厂商列表页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=16&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 16) & ">品牌显示页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=17&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 17) & ">品牌列表页模板</FONT></a>"
            'Response.Write " | "

            'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=13&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=1'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 13, IsOnlinePayment, 1) & ">在线支付第一步模板</FONT></a> | "
            'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=14&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=1'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 14, IsOnlinePayment, 1) & ">在线支付第二步模板</FONT></a> | "
            'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=15&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=1'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 15, IsOnlinePayment, 1) & ">在线支付第三步模板</FONT></a></td></tr>"
            
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=29&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 29) & ">全站专题列表页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=30&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 30) & ">全站专题页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=101&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 101) & ">自定义列表模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=0&TempType=" & TempType & "&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 0) & " >所有模板</FONT></a> | </td></tr>"
            Response.Write "</td></tr>"
        Else
            Response.Write " | <a href='Admin_Template.asp?TemplateType=8&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">会员信息页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=9&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 9) & ">会员列表页模板</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=18&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 18) & ">会员注册页模板（许可协议）</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=19&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 19) & ">会员注册页模板（注册表单）</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=21&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 21) & ">会员注册页模板（注册结果）</FONT></a>"
            If ShowUserModel = True Then			
                Response.Write " | <a href='Admin_Template.asp?TemplateType=102&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 102) & ">会员中心通用模板</FONT></a>"			
            End If				
            Response.Write "</td></tr>"
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=0&TempType=" & TempType & "&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 0) & " >所有模板</FONT></a> | </td></tr>"
            Response.Write "</td></tr>"
        End If
    End If
    Response.Write "   </table>"
    Response.Write "  </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
End If

'解决通用模板加载商城在线支付
If IsOnlinePayment > 0 Then
    ChannelID = 1000
End If

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call Save
Case "Set"
    Call SetDefault
Case "Del"
    Call DelTemplate
Case "Export"
    Call Export
Case "DoExport"
    Call DoExport
Case "Import"
    Call Import
Case "Import2"
    Call Import2
Case "DoImport"
    Call DoImport
Case "DoTemplateCopy"
    Call DoTemplateCopy
Case "ChannelCopyTemplate"
    Call ChannelCopyTemplate
Case "DoCopy"
    Call DoCopy
Case "BatchReplace"
    Call BatchReplace
Case "DoBatchReplace"
    Call DoBatchReplace
Case "BatchDefault"
    Call BatchDefault
Case "DoBatchDefault"
    Call DoBatchDefault
Case Else
    Call main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


'=================================================
'过程名：Main
'作  用：模板首页
'=================================================
Sub main()
    Dim iTemplateType, i
    Dim sql, rs, TempType
    Dim TemplateSelect, TemplateSelectContent
    Dim rsProjectName, SysDefault '系统默认
    
    TempType = PE_CLng(Trim(Request.QueryString("TempType")))
    TemplateSelect = PE_CLng(Trim(Request.Form("TemplateSelect")))

    If TemplateSelect = 1 Then
        TemplateSelectContent = Trim(Request.Form("TemplateSelectContent"))

        If TemplateSelectContent = "" Then
            ErrMsg = ErrMsg & "<li>模板查询不能为空！</li>"
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    End If

    '得到系统方案默认名称
    Set rsProjectName = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")

    If rsProjectName.BOF And rsProjectName.EOF Then
        Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
        Exit Sub
    Else
        SysDefault = rsProjectName("TemplateProjectName")
    End If

    Set rsProjectName = Nothing

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true&&e.zzz!=1)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function mysub()" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        esave.style.visibility=""visible"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf

    If TemplateSelect = 1 Then
        If ProjectName = "" Then
            sql = "select * from PE_Template where ProjectName='' or ProjectName is null"
        ElseIf ProjectName = "所有方案" Then
            sql = "select * from PE_Template"
        Else
            sql = "select * from PE_Template where ProjectName='" & ProjectName & "'"
        End If

    Else

        If TemplateType = 0 Then
            If TempType = 1 Then
                sql = "select * from PE_Template where ChannelID=" & ChannelID & " and TemplateType in (8,9,18,19,20,21)"
            Else
                sql = "select * from PE_Template where ChannelID=" & ChannelID
            End If

            If downright = 1 Then
                sql = sql & " and Deleted=" & PE_True
            Else
                sql = sql & " and Deleted=" & PE_False
            End If

        Else
            sql = "select * from PE_Template where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType

            If downright = 1 Then
                sql = sql & " and Deleted=" & PE_True
            Else
                sql = sql & " and Deleted=" & PE_False
            End If
        End If

        If ProjectName = "" Then
            sql = sql & " and ProjectName='' or ProjectName is null order by TemplateType,TemplateID"
        ElseIf ProjectName = "所有方案" Then
            sql = sql & " order by TemplateType,TemplateID"
        Else
            sql = sql & " and ProjectName='" & ProjectName & "' order by TemplateType,TemplateID"
        End If
    End If

    Set rs = Conn.Execute(sql)

    Response.Write "<form name='form1' method='post' action='Admin_Template.asp'>"

    If ChannelName = "" Then
        ChannelName = "通用模板"
    End If

    Response.Write "<IMG SRC='images/img_u.gif' height='12'>您现在的位置：" & ProjectName & "&nbsp;&gt;&gt;&nbsp;"
    If downright = 1 Then
        Response.Write "模板回收站&nbsp;&gt;&gt;" & vbCrLf
    End If
    Response.Write ChannelName & "&nbsp;&gt;&gt;&nbsp;" & GetTemplateTypeName(TemplateType, ChannelID)

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "     <tr class='title' height='22'>"
    Response.Write "      <td width='30' align='center'><strong>选择</strong></td>"
    Response.Write "      <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "      <td width='100' align='center'><strong>方案名称</strong></td>"
    Response.Write "      <td width='120' align='center'><strong>模板类型</strong></td>"
    Response.Write "      <td height='22' align='center'><strong>模板名称</strong></td>"

    If ProjectName = SysDefault Then
        Response.Write "      <td width='60' align='center'><strong>系统默认</strong></td>"
    Else
        Response.Write "      <td width='60' align='center'><strong>方案默认</strong></td>"
    End If

    Response.Write "      <td width='260' align='center'><strong>操作</strong></td>"
    Response.Write "     </tr>"
    iTemplateType = 0
    i = 0

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='10' align='center' height='50'>此模板类型中还没有模板</td></tr>"
    Else

        Do While Not rs.EOF

            If TemplateSelect <> 1 Or (TemplateSelect = 1 And InStr(rs("TemplateContent"), TemplateSelectContent) > 0) Then
                If i > 0 And rs("TemplateType") <> iTemplateType Then
                    Response.Write "<tr height='10'><td colspan='6'></td></tr>"
                End If

                iTemplateType = rs("TemplateType")
                i = i + 1
                Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
                Response.Write "  <td width=""30"" align=""center"" height=""30"">" & vbCrLf
                Response.Write "    <input type=""checkbox"" value=" & rs("TemplateID") & " name=""TemplateID"""

                If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then Response.Write "disabled"
                Response.Write "> " & vbCrLf
                Response.Write "  </td>" & vbCrLf
                Response.Write "      <td width='30' align='center'>" & rs("TemplateID") & "</td>"
                Response.Write "      <td width='100' align='center'>" & rs("ProjectName") & "</td>"
                Response.Write "      <td width='120' align='center'>" & GetTemplateTypeName(rs("TemplateType"), ChannelID) & "</td>"
                Response.Write "      <td align='center'><a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&ProjectName=" & Server.UrlEncode(rs("ProjectName")) & "&TemplateID=" & rs("TemplateID") & "'>" & rs("TemplateName") & "</a></td>"

                If ProjectName = SysDefault Then
                    Response.Write "      <td width='60' align='center'><b>"

                    If rs("IsDefault") = True Then
                        Response.Write "<FONT style='font-size:12px' color='#008000'>√</FONT>"
                    Else
                    End If

                    Response.Write "</td>"
                Else
                    Response.Write "      <td width='60' align='center'><b>"

                    If rs("IsDefaultInProject") = True Then
                        Response.Write "√"
                    Else
                    End If
                End If

                Response.Write "</td>"
                Response.Write "      <td width='260' align='center'>"

                If rs("Deleted") = True Then
                    If rs("IsDefault") = False Then
                        Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Del&TemplateID=" & rs("TemplateID") & "&downright=1" & "&ProjectName=" & Server.UrlEncode(ProjectName)

                        If TempType > 0 Then Response.Write "&TempType=" & TempType
                        Response.Write "' onClick=""return confirm('确定要彻底删除此版面设计模板吗？该模板删除后不可恢复,删除此版面设计模板后原使用此版面设计模板的文章将改为使用系统默认版面设计模板。');"">彻底删除模板</a>&nbsp;&nbsp;"
                    Else
                        Response.Write "<font color='gray'>彻底删除模板&nbsp;&nbsp;</font>"
                    End If

                    Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Del&TemplateID=" & rs("TemplateID") & "&TemplateType=" & rs("TemplateType") & "&downright=3" & "&ProjectName=" & Server.UrlEncode(ProjectName)

                    If TempType > 0 Then Response.Write "&TempType=" & TempType
                    Response.Write "'>还原模板</a><br>"
                Else

                    '设为系统默认
                    If ProjectName = SysDefault Then
                        If rs("IsDefault") = False And ProjectName = SysDefault Then
                            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Set&DefaultType=1&TemplateType=" & rs("TemplateType") & "&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)

                            If TempType > 0 Then Response.Write "&TempType=" & TempType
                            Response.Write "'>&nbsp;设为系统默认</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<font color='gray'>&nbsp;设为系统默认&nbsp;&nbsp;</font>"
                        End If

                    Else

                        If rs("IsDefaultInProject") = False Then
                            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Set&DefaultType=2&TemplateType=" & rs("TemplateType") & "&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)

                            If TempType > 0 Then Response.Write "&TempType=" & TempType
                            Response.Write "'>&nbsp;设为方案默认</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<font color='gray'>&nbsp;设为方案默认&nbsp;&nbsp;</font>"
                        End If
                    End If

                    Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&ProjectName=" & Server.UrlEncode(rs("ProjectName")) & "&TemplateID=" & rs("TemplateID")

                    If TempType > 0 Then Response.Write "&TempType=" & TempType
                    Response.Write "'>修改模板</a>&nbsp;&nbsp;"
                    If rs("IsDefault") = False And rs("IsDefaultInProject") = False Then
                        Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Del&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)
                        If TempType > 0 Then Response.Write "&TempType=" & TempType
                        Response.Write "' onClick=""return confirm('确定要删除此版面设计模板吗？删除后你可以从回收站还原它。');"">删除模板</a>"
                    Else
                        Response.Write "<font color='gray'>删除模板</font>"
                    End If

                    Response.Write "&nbsp;&nbsp;<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=DoTemplateCopy&TemplateName=" & Server.UrlEncode(rs("TemplateName")) & "&TemplateType=" & rs("TemplateType") & "&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)

                    If TempType > 0 Then Response.Write "&TempType=" & TempType
                    Response.Write "'>复制模板</a><br>"
                End If

                Response.Write " </td>"
                Response.Write "</tr>"
            End If

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</table><br>" & vbCrLf
    Response.Write "        <input name=""Action"" type=""hidden""  value=""Del"">   " & vbCrLf
    Response.Write "        <input name=""ChannelID"" type=""hidden""  value=" & ChannelID & ">" & vbCrLf

    If TempType > 0 Then Response.Write "        <input name=""TempType"" type=""hidden""  value=" & TempType & ">" & vbCrLf
    Response.Write "        <input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" >选中所有模板" & vbCrLf
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf

    If downright = 1 Then
        Response.Write "        <input name='downright' value='1' type='hidden'>"
        Response.Write "        <input type=""submit"" value="" 彻底删除 "" name=""Del"" onclick='return confirm(""确定要彻底删除选中的模板吗？彻底删除后不可恢复。"");' >&nbsp;&nbsp;" & vbCrLf
        Response.Write "        &nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" 清空回收站 "" name=""Del"" onClick=""document.form1.downright.value='2'"">&nbsp;&nbsp;" & vbCrLf
        Response.Write "        &nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" 还 原 "" name=""Del"" onClick=""document.form1.downright.value='3'"">&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" 全部还原 "" name=""Del"" onClick=""document.form1.downright.value='4'"">&nbsp;&nbsp;" & vbCrLf
    Else
        Response.Write "        <input type=""submit"" value=""批量删除 "" name=""Del"" onclick='return confirm(""确定要删除选中的模板吗？删除后你可以从回收站还原它。"");' >&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" 批量复制 "" name=""ChannelCopyTemplate"" onClick=""document.form1.Action.value='DoTemplateCopy'"">&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" 批量替换 "" name=""BatchReplace"" onClick=""document.form1.Action.value='BatchReplace'"">&nbsp;&nbsp;" & vbCrLf
    End If

    Response.Write "                        <Input TYPE='hidden' Name='BatchTypeName' value='移动'>" & vbCrLf
    Response.Write "                        <Input TYPE='hidden' Name='ProjectName' value='" & ProjectName & "'>" & vbCrLf
    Response.Write "                        <Input TYPE='hidden' Name='TemplateProjectID' value='" & TemplateProjectID & "'>" & vbCrLf

    If downright = 0 Then
        If TemplateType > 0 Then
            Response.Write "<input type='button' name='buttonm' value='添加" & GetTemplateTypeName(TemplateType, ChannelID) & "' onclick=""window.location.href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Add&TemplateType=" & TemplateType & "&ProjectName=" & ProjectName
            If TempType > 0 Then Response.Write "&TempType=" & TempType
            Response.Write "'"">"
        End If
    End If

    Response.Write "<br><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write " <tr class=""tdbg"">"
    Response.Write "   <td width='20%' align='right'>"
    Response.Write "  模板内容查询：</td>"
    Response.Write "   <td width='30%'> <TEXTAREA NAME='TemplateSelectContent'  style='width:300px;height:40px' onMouseOver=""this.select()"" onClick=""javascript:{if (form1.TemplateSelectContent.value == '请在查询框内输入要查找的字符')form1.TemplateSelectContent.value=''; };"">请在查询框内输入要查找的字符</TEXTAREA><input name='TemplateSelect' value='0' type='hidden'></td>"
    Response.Write "   <td width='50%' align='left'> <input type='submit' value=' 查 询 ' onClick=""document.form1.TemplateSelect.value='1';document.form1.Action.value='Main'"">&nbsp;&nbsp; <font color='blue'>注：</font> 本功能可查询相应内容在哪些模板中使用过。</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"

End Sub

'=================================================
'过程名：CommonLabel
'作  用：调用常用函数标签
'=================================================
Sub CommonLabel(ByVal TemplateType)
    Response.Write "        <table align='left' border='0' id='CommonLabel" & TemplateType & "' cellpadding='0' cellspacing='1' width='550' height='100%' >"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "           <td width='120'> 常用超级函数标签:</td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetArticleList','文章列表函数标签',1,'GetList',800,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetArticleList.gif' border='0' width='18' height='18' alt='显示文章标题等信息'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicArticle','显示图片文章标签',1,'GetPic',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicArticle.gif' border='0' width='18' height='18' alt='显示图片文章'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicArticle','显示幻灯片文章标签',1,'GetSlide',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicArticle.gif' border='0' width='18' height='18' alt='显示幻灯片文章'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','文章自定义列表标签',1,'GetArticleCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetArticleCustom.gif' border='0' width='18' height='18' alt='文章自定义列表'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSoftList','下载列表函数标签',2,'GetList',800,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSoftList.gif' border='0' width='18' height='18' alt='显示软件标题'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicSoft','显示图片下载标签',2,'GetPic',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicSoft.gif' border='0' width='18' height='18' alt='显示图片下载'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicSoft','显示幻灯片下载标签',2,'GetSlide',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicSoft.gif' border='0' width='18' height='18' alt='显示幻灯片下载'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','下载自定义列表标签',2,'GetSoftCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSoftCustom.gif' border='0' width='18' height='18' alt='下载自定义列表'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPhotoList','图片列表函数标签',3,'GetList',800,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPhotoList.gif' border='0' width='18' height='18' alt='显示图片标题'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicPhoto','显示图片图文标签',3,'GetPic',700,550," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicPhoto.gif' border='0' width='18' height='18' alt='显示图片'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicPhoto','显示幻灯片图片标签',3,'GetSlide',700,550," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicPhoto.gif' border='0' width='18' height='18' alt='显示图片幻灯片'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','图片自定义列表标签',3,'GetPhotoCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPhotoList.gif' border='0' width='18' height='18' alt='图片自定义列表'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetProductList','商城列表函数标签',5,'GetList',800,750," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetProductList.gif' border='0' width='18' height='18' alt='显示商品标题'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicProduct','显示图片商城标签',5,'GetPic',700,600," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicProduct.gif' border='0' width='18' height='18' alt='显示商品图片'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicProduct','显示幻灯片商城标签',5,'GetSlide',700,460," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicProduct.gif' border='0' width='18' height='18' alt='显示商品幻灯片'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','商城自定义列表标签',5,'GetProductCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetProductCustom.gif' border='0' width='18' height='18' alt='商品自定义列表'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
End Sub
'=================================================
'过程名：ADD
'作  用：添加模板
'=================================================
Sub Add()

    Dim Num, strHead, Content
    Dim rsTemplateProject
    
    TemplateType = Request.QueryString("TemplateType")
    ProjectName = Request.QueryString("ProjectName")

    '返回js代码 num 为 大类或小类
    If TemplateType = 2 Then
        Num = 2
    Else
        Num = 1
    End If
    
    '加入模板预定头部 在添加时用到
    
    strHead = "<html>" & vbCrLf
    strHead = strHead & "<head>" & vbCrLf
    strHead = strHead & "<title>新模板标题</title>" & vbCrLf
    strHead = strHead & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
    strHead = strHead & "{$Skin_CSS} {$MenuJS}" & vbCrLf
    strHead = strHead & "</head>" & vbCrLf
    strHead = strHead & "<body leftmargin=0 topmargin=0 onmousemove='HideMenu()'>" & vbCrLf
    strHead = strHead & vbCrLf & "<!-- 请输入您要设计的代码 -->" & vbCrLf
    strHead = strHead & vbCrLf & "</body>" & vbCrLf
    strHead = strHead & "</html>" & vbCrLf
        
    '替换头部标签 Content 为替换后头部文件，用于编辑器显示css
    Content = Replace(strHead, "{$Skin_CSS}", "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>")
    Content = Replace(Content, "{$MenuJS}", "<script language='JavaScript' type='text/JavaScript' src='" & InstallDir & "js/menu.js'></script>")
    Content = Replace(Content, "{$InstallDir}", InstallDir)
    '预写3000行
    Dim strContenttemp, i
    For i = 1 To 3000
        If strContenttemp = "" Then
            strContenttemp = i & vbCrLf
        Else
            strContenttemp = strContenttemp & i & vbCrLf
        End If
    Next
   
    '调入js处理过程
    Call StrJS_Template
            
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp' >"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>添 加 新 模 板</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> 选择方案： </strong><select name='ProjectName' id='ProjectName' onChange=""window.location.href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Add&TemplateType=" & TemplateType & "&TempType=" & TempType & "&ProjectName='+this.value"">" & GetProject_Option(ProjectName) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> 模板类型： </strong><select name='TemplateType' id='TemplateType' onChange=""window.location.href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Add&TempType=" & TempType & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateType='+this.value"">" & GetTemplate_Option(PE_CLng(TemplateType)) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> 模板名称： </strong><input name='TemplateName' type='text' id='TemplateName' value='' size='50' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateStart1'></a>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'align=center id='Navigation1' style='display:'>"
    
    If TemplateType = 2 Then
        Response.Write "<b>大类模板：</b>当栏目含有子栏目时，就会调用此处内容显示！"
    Else
        Response.Write "<b> 模 板 内 容 ↓</b>"
    End If

    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation1 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""OpenNavigation(1)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation1 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""CloseNavigation(1)"">&nbsp;关闭标签导航栏</a></td></tr>"
    Response.Write "        </table>"

    Call CommonLabel(1)

    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' id=showAlgebra>"
    Response.Write "      <td>"
    Response.Write "       <table>"
    Response.Write "        <tr >"
    Response.Write "          <td width='20'><table id=showLabel style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0  width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=1'></iframe></td></tr></table></td>"
    Response.Write "          <td >"
    Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
    Response.Write "              <textarea id='txt_ln' name='rollContent'  COLS='5' ROWS='31' class='RomNumber' readonly>" & strContenttemp & "</textarea>" & vbCrLf
    Response.Write "            </td><td width='700'>"
    Response.Write "             <textarea name='Content' id='txt_main'  ROWS='30' COLS='117'  class='txt_main' wrap='OFF'  onkeydown='editTab()' onscroll=""show_ln('txt_ln','txt_main')""  wrap='on' onMouseUp=""setContent('get',1);setContent2(1)"">" & Server.HTMLEncode(strHead) & "</textarea></td></tr>"
    Response.Write "             <script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script>" & vbCrLf
    Response.Write "            </td></tr>"
    Response.Write "           </table>"
    Response.Write "          </td>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation3 ><td>&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1'  onclick=""OpenNavigation(1)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation3 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1' onclick=""CloseNavigation(1)"">&nbsp;关闭标签导航栏</a></td></tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' >"
    Response.Write "    <td><table><tr>"
    Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "         <input name='EditorAlgebra' type='button' id='EditorAlgebra' value=' 代码模式 '  onclick='LoadEditorAlgebra(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorMix' type='button' id='EditorMix' value=' 混合模式 '  disabled onclick='LoadEditorMix(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorEdit' type='button' id='EditorEdit' value=' 编辑模式 ' disabled onclick='LoadEditorEdit(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Copy' type='button' id='Copy' value=' 复制代码 ' onclick='copy(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Editorfullscreen' type='button' id='Editorfullscreen' value=' 全屏编辑 ' onclick='fullscreen(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorSkin' type='button' id='EditorSkin' value=' 修改风格 ' onClick=""return Templateskin()"" onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "       </td>"
    Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content');sizeContent(5,'rollContent')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content');sizeContent(-5,'rollContent')"">&nbsp;&nbsp;</td></tr>"
    Response.Write "     </tr></table>"
    Response.Write "    </td></tr>"
    Response.Write "    <a name='#TemplateEnd1'></a>"
    Response.Write "    <tr class='tdbg' id=showeditor style='display:none'>"
    Response.Write "      <td valign='top'>"
    Response.Write "       <table >"
    Response.Write "        <tr><td width='20'><td>"
    Response.Write "       <textarea name='EditorContent' style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent&TemplateType=1' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
    Response.Write "       </td></tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If TemplateType = 2 Then
        Response.Write "    <a name='#TemplateStart2'></a>"
        Response.Write "    <tr class='tdbg' >"
        Response.Write "      <td valign='top' align='center'>"
        Response.Write "     <b>小类模板：</b>当栏目没有子栏目时，就会调用此处内容显示</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='center'  align='left' valign='top'>"
        Response.Write "        <table align='left' width='200'  id='Navigation12' style='display:'>"
        Response.Write "          <tr id=OpenNavigation2 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""OpenNavigation(2)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
        Response.Write "          <tr id=CloseNavigation2 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""CloseNavigation(2)"">&nbsp;关闭标签导航栏</a></td></tr>"
        Response.Write "        </table>"

        Call CommonLabel(2)

        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg' id=showAlgebra2>"
        Response.Write "      <td>"
        Response.Write "       <table>"
        Response.Write "        <tr  >"
        Response.Write "          <td width='20'><table id=showLabel2 style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0 width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=2'></iframe></td></tr></table></td>"
        Response.Write "          <td >"
        Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
        Response.Write "           <textarea id='txt_ln2' name='rollContent2'  COLS='5' ROWS='31' class=RomNumber readonly>" & strContenttemp & "</textarea>" & vbCrLf
        Response.Write "            </td><td width='700'>"
        Response.Write "           <textarea name='Content2' id='txt_main2'  ROWS='30' COLS='117' wrap='OFF' id='TemplateContent2' class='txt_main'  onkeydown='editTab()' onscroll=""show_ln('txt_ln2','txt_main2')"" onMouseUp=""setContent('get',2);setContent2(2)"">" & Server.HTMLEncode(strHead) & "</textarea></td></tr>"
        Response.Write "           <script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln2').value += i + '\n';</script>" & vbCrLf
        Response.Write "            </td></tr>"
        Response.Write "           </table>"
        Response.Write "          </td>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg' >"
        Response.Write "    <td><table><tr>"
        Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "         <input name='EditorAlgebra' type='button' id='EditorAlgebra2' value=' 代码模式 '  onclick='LoadEditorAlgebra(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorMix2' type='button' id='EditorMix2' value=' 混合模式 ' disabled onclick='LoadEditorMix(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorEdit2' type='button' id='EditorEdit2' value=' 编辑模式 '  disabled onclick='LoadEditorEdit(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Copy2' type='button' id='Copy2' value=' 复制代码 '  onclick='copy(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Editorfullscreen2' type='button' id='Editorfullscreen2' value=' 全屏编辑 ' onclick='fullscreen(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorSkin2' type='button' id='EditorSkin' value=' 修改风格 ' onClick=""return Templateskin()"" onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "       </td>"
        Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content')"">&nbsp;&nbsp;</td></tr>"
        Response.Write "     </tr></table>"
        Response.Write "    <a name='#TemplateEnd2'></a>"
        Response.Write "    </td></tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "        <table align='left' width='200'>"
        Response.Write "          <tr id=OpenNavigation4 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""OpenNavigation(2)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
        Response.Write "          <tr id=CloseNavigation4 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""CloseNavigation(2)"">&nbsp;关闭标签导航栏</a></td></tr>"
        Response.Write "        </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"

        Response.Write "    <tr class='tdbg' id=showeditor2 style='display:none'>"
        Response.Write "      <td valign='top'>"
        Response.Write "       <table >"
        Response.Write "        <tr><td width='20'><td>"
        Response.Write "       <textarea name='EditorContent2' style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
        Response.Write "       <iframe ID='editor2' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent2&TemplateType=2' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
        Response.Write "       </td></tr></table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
    End If

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'>&nbsp;&nbsp;<input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'> 将此模板设为"
    Set rsTemplateProject = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rsTemplateProject.BOF And rsTemplateProject.EOF Then
        Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
        Exit Sub
    Else
        If ProjectName = rsTemplateProject("TemplateProjectName") Then
            Response.Write "系统"
        Else
            Response.Write "方案"
        End If
    End If
    Set rsTemplateProject = Nothing
    Response.Write "默认模板</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='50'  align='center'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "       <input name='Action' type='hidden' id='Action' value='SaveAdd'><input type='button' name='button' value=' 添 加 ' onClick='return CheckForm(" & Num & ");'>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    
End Sub

'=================================================
'过程名：Modify
'作  用：修改模板
'=================================================
Sub Modify()
    
    Dim TemplateID, TemplateContent, TemplateContent2
    Dim arrContent, rs, sql, Num, Content, Content2
    Dim strTemp
    Dim rsTemplateProject
    '获取模板ID
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
 
    If TemplateID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定TemplateID</li>"
        Exit Sub
    End If

    '得到模板内容
    sql = "select * from PE_Template where TemplateID=" & TemplateID
    
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的模板！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    TemplateContent = rs("TemplateContent")
    
    If rs("TemplateType") = 2 Then
        '给前台js为 栏目类型
        Num = 2
        arrContent = Split(TemplateContent, "{$$$}")
        TemplateContent = arrContent(0)
        TemplateContent2 = arrContent(1)
        Content = ShiftCharacter(TemplateContent)
        Content2 = ShiftCharacter(TemplateContent2)
    Else
        Num = 1
        Content = ShiftCharacter(TemplateContent)
    End If

    '与4.03 兼容 替换内容页显示不好看问题
    If rs("TemplateType") = 3 Then
        regEx.Pattern = "(\<noscript)([\s\S]*?)(\<\/noscript\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            strTemp = Match.value
            Content = Replace(Content, strTemp, "")
        Next
    End If

    '预写3000行
    Dim strContenttemp, i
    For i = 1 To 3000
        If strContenttemp = "" Then
            strContenttemp = i & vbCrLf
        Else
            strContenttemp = strContenttemp & i & vbCrLf
        End If
    Next

    '加载前台js
    Call StrJS_Template
    
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp' >"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22'  align='center'><strong> 修 改 模 板 设 置</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> 选择方案： </strong><select name='ProjectName' id='ProjectName' disabled>" & GetProject_Option(ProjectName) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> 模板类型： </strong><select name='TemplateType' disabled>" & GetTemplate_Option(PE_CLng(rs("TemplateType"))) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> 模板名称： </strong><input name='TemplateName' type='text' id='TemplateName' value='" & rs("TemplateName") & "' size='50' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateStart1'></a>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align=center>"
    If rs("TemplateType") = 2 Then
        Response.Write "<b>大类模板：</b>当栏目含有子栏目时，就会调用此处内容显示！"
    Else
        Response.Write "<b> 模 板 内 容 ↓</b>"
    End If
    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'align=center id='Navigation1' style='display:'>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation1 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""OpenNavigation(1)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation1 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""CloseNavigation(1)"">&nbsp;关闭标签导航栏</a></td></tr>"
    Response.Write "        </table>"

    Call CommonLabel(1)

    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg' id=showAlgebra>"
    Response.Write "      <td>"
    Response.Write "       <table>"
    Response.Write "        <tr >"
    Response.Write "          <td width='20'><table id=showLabel style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0  width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=1'></iframe></td></tr></table></td>"
    Response.Write "          <td>"
    Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
    Response.Write "            <textarea id='txt_ln' name='rollContent'  COLS='5' ROWS='31'   class=RomNumber readonly>" & strContenttemp & "</textarea>" & vbCrLf
    Response.Write "            </td><td width='700'>"
    Response.Write "            <textarea name='Content' id='txt_main'  ROWS='30' COLS='117'  wrap='OFF'  onkeydown='editTab()' onscroll=""show_ln('txt_ln','txt_main')"" wrap='on' onMouseUp=""setContent('get',1);setContent2(1)"" class='txt_main'>" & Server.HTMLEncode(TemplateContent) & "</textarea></td></tr>"
    Response.Write "            <script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script>" & vbCrLf
    Response.Write "            </td></tr>"
    Response.Write "           </table>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "    <td><table><tr>"
    Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "         <input name='EditorAlgebra' type='button' id='EditorAlgebra' value=' 代码模式 ' onclick='LoadEditorAlgebra(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorMix' type='button' id='EditorMix' value=' 混合模式 ' disabled onclick='LoadEditorMix(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorEdit' type='button' id='EditorEdit' value=' 编辑模式 ' disabled onclick='LoadEditorEdit(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Copy' type='button' id='Copy' value=' 复制代码 ' onclick='copy(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Editorfullscreen' type='button' id='Editorfullscreen' value=' 全屏编辑 ' onclick='fullscreen(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorSkin' type='button' id='EditorSkin' value=' 修改风格 ' onClick=""return Templateskin()""  onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "       </td>"
    Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content');sizeContent(5,'rollContent')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content');sizeContent(-5,'rollContent')"">&nbsp;&nbsp;</td></tr>"
    Response.Write "     </tr></table>"
    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation3 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1' onclick=""OpenNavigation(1)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation3 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1' onclick=""CloseNavigation(1)"">&nbsp;关闭标签导航栏</a></td></tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' id=showeditor style='display:none'>"
    Response.Write "      <td valign='top' >"
    Response.Write "       <table >"
    Response.Write "        <tr><td width='20'><td>"
    Response.Write "       <textarea name='EditorContent' style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent&TemplateType=1' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
    Response.Write "       </td></tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateEnd1'></a>"
    '当有小类模板时
    If rs("TemplateType") = 2 Then
        Response.Write "    <a name='#TemplateStart2'></a>"
        Response.Write "<tr class='tdbg'>"
        Response.Write "   <td align='center'  align='left' valign='top'>"
        Response.Write "    <b>小类模板：</b>当栏目没有子栏目时，就会调用此处内容显示</td>"
        Response.Write "   </td>"
        Response.Write "</tr>"
        Response.Write "<tr class='tdbg'>"
        Response.Write "   <td align='center'  align='left' valign='top'>"
        Response.Write "    <table align='left' width='200' id='Navigation12' style='display:'>"
        Response.Write "      <tr id=OpenNavigation2 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""OpenNavigation(2)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
        Response.Write "      <tr id=CloseNavigation2 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""CloseNavigation(2)"">&nbsp;关闭标签导航栏</a></td></tr>"
        Response.Write "    </table>"

        Call CommonLabel(2)

        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "   <tr class='tdbg' id=showAlgebra2>"
        Response.Write "      <td>"
        Response.Write "       <table>"
        Response.Write "        <tr >"
        Response.Write "          <td width='20'><table id=showLabel2 style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0  width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=2'></iframe></td></tr></table></td>"
        Response.Write "          <td>"
        Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
        Response.Write "            <textarea id='txt_ln2' name='rollContent2'  COLS='5' ROWS='31'   class=RomNumber readonly>" & strContenttemp & "</textarea>" & vbCrLf
        Response.Write "            </td><td width='700'>"
        Response.Write "            <textarea name='Content2' id='txt_main2'  ROWS='30' COLS='117' wrap='OFF' id='TemplateContent2' class='txt_main' onkeydown=""editTab()"" onscroll=""show_ln('txt_ln2','txt_main2')"" onMouseUp=""setContent('get',2);setContent2(2)"">" & Server.HTMLEncode(TemplateContent2) & "</textarea></td></tr>"
        Response.Write "            <script>for(var  i=3000; i<=3000; i++) document.getElementById('txt_ln2').value += i + '\n';</script>" & vbCrLf
        Response.Write "            </td></tr>"
        Response.Write "           </table>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "        <table align='left' width='200'>"
        Response.Write "          <tr id=OpenNavigation4 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""OpenNavigation(2)"">&nbsp;使用更多的标签&nbsp;</a></td></tr>"
        Response.Write "          <tr id=CloseNavigation4 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""CloseNavigation(2)"">&nbsp;关闭标签导航栏</a></td></tr>"
        Response.Write "        </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg' >"
        Response.Write "    <td><table><tr>"
        Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "         <input name='EditorAlgebra2' type='button' id='EditorAlgebra2' value=' 代码模式 ' onclick='LoadEditorAlgebra(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorMix2' type='button' id='EditorMix2' value=' 混合模式 ' disabled onclick='LoadEditorMix(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorEdit2' type='button' id='EditorEdit2' value=' 编辑模式 ' disabled onclick='LoadEditorEdit(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Copy2' type='button' id='Copy2' value=' 复制代码 ' onclick='copy(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Editorfullscreen2' type='button' id='Editorfullscreen2' value=' 全屏编辑 ' onclick='fullscreen(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorSkin' type='button' id='EditorSkin' value=' 修改风格 ' onClick=""return Templateskin()"" onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "       </td>"
        Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content2');sizeContent(5,'rollContent2')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content2');sizeContent(-5,'rollContent2')"">&nbsp;&nbsp;</td></tr>"
        Response.Write "     </tr></table>"
        Response.Write "    </td></tr>"
        Response.Write "  <tr class='tdbg'id=showeditor2 style='display:none'>"
        Response.Write "   <td valign='top' >"
        Response.Write "     <table >"
        Response.Write "      <tr><td width='20'><td>"
        Response.Write "       <textarea name='EditorContent2' style='display:none' >" & Server.HTMLEncode(Content2) & "</textarea>"
        Response.Write "       <iframe ID='editor2' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent2&TemplateType=2' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
        Response.Write "       </td>"
        Response.Write "      </tr>"
        Response.Write "     </table>"
        Response.Write "   </td>"
        Response.Write "</tr>"
    
    End If

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'>&nbsp;&nbsp;<input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'"

    If rs("IsDefault") = True Then Response.Write " checked"
    Response.Write "> 将此模板设为"

    Set rsTemplateProject = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rsTemplateProject.BOF And rsTemplateProject.EOF Then
        Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
        Exit Sub
    Else
        If ProjectName = rsTemplateProject("TemplateProjectName") Then
            Response.Write "系统"
        Else
            Response.Write "方案"
        End If
    End If
    Set rsTemplateProject = Nothing

    Response.Write "默认模板</td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateEnd2'></a>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='50'  align='center'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='TemplateID' type='hidden' id='TemplateID' value='" & TemplateID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'><input name='TemplateType' type='hidden' id='Action' value='" & rs("TemplateType") & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "        <input type='button' name='Submit2' value=' 保存修改结果 ' onClick='return CheckForm(" & Num & ");'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"

    rs.Close
    Set rs = Nothing
End Sub

'=================================================
'过程名：Save
'作  用：保存模板
'=================================================
Sub Save()
    
    Dim rs, sql, Action
    Dim TemplateID, ProjectName, TemplateName, IsDefault, IsDefaultInProject
    Dim DefaultType, setUpdateItem
    Dim TemplateContent, TemplateContent2, i
    
    '得到模板ID 名称 类型
    TemplateID = Trim(Request.Form("TemplateID"))
    TemplateName = Trim(Request.Form("TemplateName"))
    Action = Trim(Request.Form("Action"))
    TemplateType = Trim(Request.Form("TemplateType"))
    ProjectName = Trim(Request.Form("ProjectName"))
    '错误处理
    If Action = "SaveModify" Then
        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定TemplateID</li>"
        Else
            TemplateID = PE_CLng(TemplateID)
        End If
        Set rs = Conn.Execute("Select TemplateID,ProjectName From PE_Template Where TemplateID=" & TemplateID & "")
        If rs.BOF And rs.EOF Then
            Call WriteErrMsg("<li>系统中还没此模板！</li>", ComeUrl)
            Exit Sub
        Else
            ProjectName = rs("ProjectName")
        End If
        Set rs = Nothing
    End If
    
    If TemplateName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>模板名称不能为空！</li>"
    End If
         
    For i = 1 To Request.Form("Content").Count
        TemplateContent = TemplateContent & Request.Form("Content")(i)
    Next
    
    For i = 1 To Request.Form("Content2").Count
        TemplateContent2 = TemplateContent2 & Request.Form("Content2")(i)
    Next
    
    If TemplateType <> 2 Then
        TemplateContent = ShiftCharacterSave(TemplateContent)
    Else
        TemplateContent = ShiftCharacterSave(TemplateContent)
        TemplateContent2 = ShiftCharacterSave(TemplateContent2)
    End If
    
    If Len(TemplateName) > 50 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>模板标题不能超过50个字符！ </li>"
    End If
        
    If InStr(TemplateContent, "rsClass_") > 0 And Not ((TemplateType = 1 And ChannelID <> 0) Or TemplateType = 2 Or TemplateType = 101) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>除了频道首页和栏目大类模板外，其他模板中的标签均不可以使用 rsClass_开头的标签参数！</li>"
    End If

    If InStr(TemplateContent2, "rsClass_") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>栏目小类模板中不允许的使用 rsClass_开头的标签参数！ </li>"
    End If

    If TemplateType = 101 Then
        If InStr(TemplateContent, "【/ArticleList】") > 0 Or InStr(TemplateContent, "【/ProductList】") Or InStr(TemplateContent, "【/SoftList】") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>自定义列表模板不能包含自己类型的模板！ </li>"
        End If
    End If

    If UBound(Split(TemplateContent, "<!--")) > UBound(Split(TemplateContent, "-->")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请检查录入模板 &lt;!-- 数量大于 --&gt; 数量会引起注释后面模板不解析！ </li>"
    End If

    If TemplateType = 2 Then
        If UBound(Split(TemplateContent2, "<!--")) > UBound(Split(TemplateContent2, "-->")) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请检查小类模板内容 &lt;!-- 数量大于 --&gt; 数量会引起注释后面模板不解析！ </li>"
        End If
    End If
            
    If FoundErr = True Then Exit Sub
    '获得 是否 定义为默认
    IsDefault = Trim(Request("IsDefault"))

    '判断是否默认
    If IsDefault = "Yes" Then
        IsDefault = True
    Else
        IsDefault = False
    End If

    '执行默认选择
    If IsDefault = True Then
        '----------------------------------------------------
        '判断是否系统默认方案
        Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
        If rs.BOF And rs.EOF Then
            Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
            Exit Sub
        Else
            If ProjectName = rs("TemplateProjectName") Then
                DefaultType = 1
            Else
                DefaultType = 2
            End If
        End If
        Set rs = Nothing

        If DefaultType = 1 Then
            setUpdateItem = "IsDefault=" & PE_False & ",IsDefaultInProject=" & PE_False
        ElseIf DefaultType = 2 Then
            setUpdateItem = "IsDefaultInProject=" & PE_False
        End If
        Conn.Execute ("update PE_Template set " & setUpdateItem & " where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType & " and ProjectName='" & ProjectName & "'")
    End If
    '----------------------------------------------------
    '添加保存
    If Action = "SaveAdd" Then
        sql = "select top 1 * from PE_Template"
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open sql, Conn, 1, 3
        rs.addnew
        rs("ChannelID") = ChannelID
        rs("Templatetype") = TemplateType
        rs("ProjectName") = ProjectName
        rs("TemplateName") = TemplateName
        
        '对小类模板的判断
        If TemplateType = 2 Then
            rs("TemplateContent") = TemplateContent & "{$$$}" & TemplateContent2
        Else
            rs("TemplateContent") = TemplateContent
        End If
        If IsDefault = True Then
            If DefaultType = 1 Then
                rs("IsDefault") = True
                rs("IsDefaultInProject") = True
            Else
                rs("IsDefault") = False
                rs("IsDefaultInProject") = True
            End If
        Else
            rs("IsDefault") = False
            rs("IsDefaultInProject") = False
        End If
        rs.Update
        rs.Close
        Set rs = Nothing
        Call WriteSuccessMsg("成功添加新的模板：" & Trim(Request("TemplateName")), ComeUrl)
    Else
        '修改保存
        sql = "select * from PE_Template where TemplateID=" & TemplateID
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open sql, Conn, 1, 3

        If rs.BOF And rs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的版面设计模板！</li>"
        Else

            If rs("TemplateType") = 2 Then
                If TemplateContent2 = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>小类模板内容不能为空！</li>"
                    rs.Close
                    Set rs = Nothing
                    Exit Sub
                End If

                rs("TemplateContent") = TemplateContent & "{$$$}" & TemplateContent2
            Else
                rs("TemplateContent") = TemplateContent
            End If

            rs("TemplateName") = TemplateName

            If IsDefault = True Then
                If DefaultType = 1 Then
                    rs("IsDefault") = True
                    rs("IsDefaultInProject") = True
                Else
                    rs("IsDefault") = False
                    rs("IsDefaultInProject") = True
                End If
            End If
            rs.Update
            Call WriteSuccessMsg("保存模板成功！", ComeUrl)
        End If

        rs.Close
        Set rs = Nothing

        If IsDefault = True And TemplateType = 1 Then
            If ChannelID = 0 Then
                
                Dim FileExt_SiteIndex, FileName_Index
                FileExt_SiteIndex = arrFileExt(FileExt_SiteIndex)
                FileName_Index = "Index" & FileExt_SiteIndex
                
                If FileName_Index = "Index.asp" Then
                    ErrMsg = ErrMsg & "<li>因为网站配置中未启用网站首页生成HTML功能，所以不用生成首页。</li>"
                    Response.Write ErrMsg
                    Exit Sub
                End If
                
                If ObjInstalled_FSO = True Then
                    Response.Write "<br><iframe  width='100%' height='210' frameborder='0' src='Admin_CreateSiteIndex.asp'></iframe>"
                Else
                    ErrMsg = ErrMsg & "<li>因为网站不支持FSO 或 您的FSO已更名。</li>"
                    Response.Write ErrMsg
                    Exit Sub
                End If

            Else

                If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
                    Response.Write "<br><iframe  width='100%' height='210' frameborder='0' src='Admin_Create" & ModuleName & ".asp?ChannelID=" & ChannelID & "&CreateType=1&Action=CreateIndex&ShowBack=No'></iframe>"
                End If
            End If
        End If
    End If

    Call ClearSiteCache(0)
End Sub

'=================================================
'过程名：SetDefault
'作  用：保存指定的默认模板
'=================================================
Sub SetDefault()
    Dim TemplateID, DefaultType, setUpdateItem, setUpdateItem2, strTemp, ProjectName
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
    DefaultType = PE_CLng(Trim(Request("DefaultType")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If TemplateID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定TemplateID</li>"
        Exit Sub
    End If
        
    If DefaultType = 1 Then
        setUpdateItem = "IsDefault=" & PE_False & ",IsDefaultInProject=" & PE_False
        setUpdateItem2 = "IsDefault=" & PE_True & ",IsDefaultInProject=" & PE_True
        strTemp = "<li>成功将选定的模板,设置为<FONT style='font-size:12px' color='#008000'>系统默认</FONT>模板.</li>"
    ElseIf DefaultType = 2 Then
        setUpdateItem = "IsDefaultInProject=" & PE_False
        setUpdateItem2 = "IsDefaultInProject=" & PE_True
        strTemp = "<li>成功将选定的模板,设置为<FONT style='font-size:12px' color='#3366FF'>方案默认</FONT>模板.</li>"
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>设定的默认类型不对!</li>"
        Exit Sub
    End If

    Conn.Execute ("update PE_Template set " & setUpdateItem & " where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType & " and ProjectName='" & ProjectName & "'")
    Conn.Execute ("update PE_Template set " & setUpdateItem2 & " where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType & " and TemplateID=" & TemplateID)
    Call WriteSuccessMsg(strTemp, ComeUrl)
    Call ClearSiteCache(0)
End Sub

'=================================================
'过程名：DelTemplate
'作  用：删除指定模板
'=================================================
Sub DelTemplate()
    Dim TemplateID, rs, trs, sql, downright

    FoundErr = False

    'downright 0 删除到数据库 1 选定模板彻底删除 2 晴空回收站 3 选定的还原 4 全部还原
    downright = PE_CLng(Trim(Request("downright")))
    TemplateID = Trim(Request("TemplateID"))
	If IsValidID(TemplateID) = False Then
		TemplateID = ""
	End If

    If downright = 2 Or downright = 4 Then
        sql = "select * from PE_Template where Deleted=" & PE_True
    Else

        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定TemplateID</li>"
            Exit Sub
        End If

        If InStr(TemplateID, ",") > 0 Then
            sql = "select * from PE_Template where TemplateID In(" & TemplateID & ")"
        Else
            sql = "select * from PE_Template where TemplateID=" & PE_CLng(TemplateID)
        End If
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的版面设计模板！</li>"
    Else

        Do While Not rs.EOF

            If downright = 1 Or downright = 2 Then
                If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & rs("TemplateName") & "当前模板为方案默认模板，不能删除。</li><li>请先将默认模板改为其他模板后再来删除此模板。</li>"
                Else
                    Set trs = Conn.Execute("select TemplateID from PE_Template where ChannelID=" & ChannelID & " and IsDefault=" & PE_True & " and TemplateType=" & rs("TemplateType"))

                    If trs.BOF And trs.EOF Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>" & rs("TemplateName") & "找不到可替用的默认模板，所以不能删除当前模板。请先将其中一个模板改为默认模板后再来删除此模板。</li>"
                    Else
                        Select Case rs("TemplateType")

                            Case 1
                                Conn.Execute ("update PE_Channel set Template_Index=0 where ChannelID=" & ChannelID & " and Template_Index=" & rs("TemplateID"))

                            Case 2
                                Conn.Execute ("update PE_Class set TemplateID=0 where ChannelID=" & ChannelID & " and TemplateID=" & rs("TemplateID"))

                            Case 3
                                Conn.Execute ("update PE_Article set TemplateID=0 where ChannelID=" & ChannelID & " and TemplateID=" & rs("TemplateID"))

                            Case 4
                                Conn.Execute ("update PE_Special set TemplateID=0 where ChannelID=" & ChannelID & " and TemplateID=" & rs("TemplateID"))
                        End Select

                        TemplateType = rs("TemplateType")

                        If downright = 1 Then
                            ErrMsg = ErrMsg & "<li>成功删除 <font color=red>" & rs("TemplateName") & "</font>模板。并将使用此模板的栏目和文章改为使用默认模板。</li><br>"
                        End If

                        rs.Delete
                        rs.Update
                    End If

                    Set trs = Nothing
                End If

            ElseIf downright = 3 Or downright = 4 Then
                rs("Deleted") = False
                If downright = 3 Then
                    ErrMsg = "<FONT color='blue'>" & rs("TemplateName") & "</FONT>模板已经还原！"
                End If
            Else

                If rs("IsDefaultInProject") = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & rs("TemplateName") & "当前模板为方案默认模板，不能删除。</li><li>请先将默认模板改为其他模板后再来删除此模板。</li>"
                Else
                    rs("Deleted") = True
                    ErrMsg = ErrMsg & "成功删除<font color=red>" & rs("TemplateName") & "</font>模板。您可以在模板回收站恢复它们<br>"
                End If
            End If

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing


    If FoundErr = False Then
        If downright = 2 Then
            ErrMsg = "<li>成功的清空了模板回收站。</li>"
            
        ElseIf downright = 4 Then
            ErrMsg = "<li>成功将全部模板还原！</li>"
        End If
        Call WriteSuccessMsg(ErrMsg, ComeUrl)
    End If

End Sub

'=================================================
'过程名：BatchDefault
'作  用：批量设置默认
'=================================================
Sub BatchDefault()
    Dim sql, rs
    Dim iTemplateType, iChannelID, i, Num
    Dim rsTemplateProject, sqlTemplateProject, IsProjectDefault
    iChannelID = 0
    iTemplateType = 0
    i = 0
    Num = 1

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td>"
    Response.Write "    | <a href='Admin_Template.asp?Action=BatchDefault&ChannelID=0&ProjectName=" & Server.UrlEncode(ProjectName) & "'><FONT style='font-size:12px' " & IsFontChecked(ChannelID, 0) & ">网站通用模板</FONT></a>"
    i = 0
    sql = "SELECT DISTINCT t.ChannelID,c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID where c.Disabled=" & PE_False
    Set rs = Conn.Execute(sql)
        
    If rs.BOF And rs.EOF Then
        IsProjectDefault = False
    Else

        Do While Not rs.EOF
            Response.Write "    | <a href='Admin_Template.asp?Action=BatchDefault&ChannelID=" & rs("ChannelID") & "&ProjectName=" & Server.UrlEncode(ProjectName) & "'><FONT style='font-size:12px' " & IsFontChecked(rs("ChannelID"), ChannelID) & ">" & rs("ChannelName") & "频道模板</FONT></a>"

            If i > 3 Then
                Response.Write " | </td><tr class='title'><td>"
                i = 0
            Else
                i = i + 1
            End If

            rs.MoveNext
        Loop

        Response.Write " | "
    End If

    rs.Close
    Set rs = Nothing

    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    
    '是不是当前系统默认方案
    Set rs = Conn.Execute("select * from PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rs.BOF And rs.EOF Then
        IsProjectDefault = False
    Else
        If rs("TemplateProjectName") = ProjectName Then
            IsProjectDefault = True
        Else
            IsProjectDefault = False
        End If
    End If
    Set rs = Nothing

    sql = "select * from PE_Template where Deleted=" & PE_False & " and ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "' order by TemplateType,ChannelID"
        
    Set rs = Conn.Execute(sql)

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true&&e.zzz!=1)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "     <tr class='title' height='22'>"
    Response.Write "      <td width='30' align='center'><strong>选择</strong></td>"
    Response.Write "      <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "      <td width='150' align='center'><b>模板类型</b></td>"
    Response.Write "      <td height='22' align='center'><strong>模板名称</strong></td>"
    Response.Write "      <td width='85' align='center'><strong>是否"
    If IsProjectDefault = True Then
        Response.Write "系统"
    Else
        Response.Write "方案"
    End If
    Response.Write "默认</strong></td>"
    Response.Write "     </tr>"
    i = 0

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td width='100%' colspan='6' align='center'> 该 方 案 还 没 有 模 板</td></tr>"
    Else

        Do While Not rs.EOF

            If i > 0 And rs("TemplateType") <> iTemplateType Or i > 0 And rs("ChannelID") <> iChannelID Then
                Num = Num + 1
                Response.Write "<tr height='10'><td colspan='6'></td></tr>"
            End If

            iChannelID = rs("ChannelID")
            iTemplateType = rs("TemplateType")
            i = i + 1

            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "  <td width=""30"" align=""center"" height=""30"">" & vbCrLf
            Response.Write "    <input TYPE='radio' value='" & rs("TemplateID") & "' name=""TemplateID" & Num & """"
            If IsProjectDefault = True Then
                If rs("IsDefault") = True Then Response.Write "checked"
            Else
                If rs("IsDefaultInProject") = True Then Response.Write "checked"
            End If
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "      <td width='30' align='center'>" & rs("TemplateID") & "</td>"
            Response.Write "      <td width='150' align='center'>" & GetTemplateTypeName(rs("TemplateType"), rs("ChannelID")) & "</td>"
            Response.Write "      <td align='center'><a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&TemplateID=" & rs("TemplateID") & "'>" & rs("TemplateName") & "</a></td>"
            Response.Write "      <td width='80' align='center'><b>"

            If IsProjectDefault = True Then
                If rs("IsDefault") = True Then
                    Response.Write "√"
                Else
                    Response.Write "×"
                End If
            Else
                If rs("IsDefaultInProject") = True Then
                    Response.Write "√"
                Else
                    Response.Write "×"
                End If
            End If

            Response.Write "</td>"
            Response.Write "</tr>"

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "    <tr class=""tdbg""> " & vbCrLf
    Response.Write "      <td colspan=6 height=""30"" align='left'>" & vbCrLf
    Response.Write "        <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中所有模板&nbsp;<FONT style='font-size:12px' color='blue'>（注意选择所有模板后，如果模板类型有多个模板，系统将会选择该类型的最后一个模板）</FONT> &nbsp;&nbsp;"
    Response.Write "        <input name=""ProjectName"" type=""hidden""  value=" & ProjectName & ">   " & vbCrLf
    Response.Write "        <input name=""ContentNum"" type=""hidden""  value=" & Num & ">   " & vbCrLf
    Response.Write "        <input name=""Action"" type=""hidden""  value=""DoBatchDefault"">   " & vbCrLf
    Response.Write "        <input name=""ChannelID"" type=""hidden""  value=" & ChannelID & ">" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr> " & vbCrLf
    Response.Write "</table>  "
    Response.Write "<br><center><input type=""submit"" value="" 批 量 设 制 方 案 默 认 ""></center><br>" & vbCrLf
    Response.Write "</form>"
End Sub

'=================================================
'过程名：DoBatchDefault
'作  用：批量设置处理
'=================================================
Sub DoBatchDefault()
    Dim ContentNum, ProjectName, arrTemplateID, arrContent, i, DefaultType

    ContentNum = PE_CLng(Trim(Request("ContentNum")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
        
    For i = 1 To ContentNum
        arrTemplateID = arrTemplateID & PE_CLng(Trim(Request("TemplateID" & i & ""))) & ","
    Next

    '判断是否系统默认方案
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")

    If rs.BOF And rs.EOF Then
        Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
        Exit Sub
    Else

        If ProjectName = rs("TemplateProjectName") Then
            DefaultType = "IsDefault"
        Else
            DefaultType = "IsDefaultInProject"
        End If
    End If

    Set rs = Nothing

    arrTemplateID = Left(arrTemplateID, Len(arrTemplateID) - 1)
    If DefaultType = "IsDefaultInProject" Then
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_False & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "'")
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_True & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "' and TemplateID in (" & arrTemplateID & " )")
    Else
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_False & ",IsDefaultInProject=" & PE_False & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "'")
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_True & ",IsDefaultInProject=" & PE_True & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "' and TemplateID in (" & arrTemplateID & " )")
    End If
    Call WriteSuccessMsg("成功将选定的模板设置为默认模板", ComeUrl)
    Call ClearSiteCache(0)
End Sub

'=================================================
'过程名：Export
'作  用：导出模板
'=================================================
Sub Export()
    
    Dim rs, sql
    Dim trs, iCount, ModuleType, ProjectName
    
    '999999 为所有
    ModuleType = Trim(Request.Form("ModuleType"))
    If ReplaceBadChar(Trim(Request.QueryString("ProjectName"))) = "" Then
        ProjectName = ReplaceBadChar(Trim(Request.Form("ProjectName")))
    Else
        ProjectName = ReplaceBadChar(Trim(Request.QueryString("ProjectName")))
    End If


    If ModuleType = "" Then
        ModuleType = 999999
    End If
 
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>模板导出</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "            <select name='ProjectName' id='ProjectName' style='width:150px;'  onChange='document.myform.submit();' >"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何方案项目</option>"
    Else
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "            </select>"
    Response.Write "            <br>"
    Response.Write "             <select name='ModuleType' onChange='document.myform.submit();'>"
    Call GetAllModule("5.0", ModuleType)
    Response.Write "             </select>"
    Response.Write "            </td><td></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='TemplateID' size='2' multiple style='height:300px;width:450px;'>"
    
    '判断是否是所有（999999），首页（0），指定的模块
    If CLng(ModuleType) = 999999 Then
        sql = "select ChannelID,TemplateID,TemplateName,ProjectName from PE_Template where Deleted=" & PE_False & " And ProjectName='" & ProjectName & "'"
        If FoundInArr(AllModules, "Supply", ",") = False Then
            sql = sql & " And ChannelID <>999 "
        End If
        If FoundInArr(AllModules, "House", ",") = False Then
            sql = sql & " And ChannelID <>998 "
        End If
        If FoundInArr(AllModules, "Job", ",") = False Then
            sql = sql & " And ChannelID <>997 "
        End If
    ElseIf CLng(ModuleType) = 0 Then
        sql = "select ChannelID,TemplateID,TemplateName,ProjectName from PE_Template where ChannelID=0 And ProjectName='" & ProjectName & "'"
    Else
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,t.ProjectName,c.ModuleType from PE_Template t inner join PE_Channel c on t.ChannelID=c.ChannelID where c.ChannelType < 2 and c.Disabled=" & PE_False & " and c.ModuleType=" & PE_CLng(ModuleType) & " And t.ProjectName='" & ProjectName & "' Order by t.ChannelID asc"
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>没有任何模板</option>"
        '关闭提交按钮
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "         </select>"
    Response.Write "       </td>"
    Response.Write "       <td align='left'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' 选定所有 ' onclick='SelectAll()'>"
    Response.Write "       <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='30'>"
    Response.Write "        <td colspan='2'>目标数据库：<input name='TemplateMdb' type='text' id='TemplateMdb' value='../temp/Template.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> 先清空目标数据库</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='执行导出操作' onClick=""document.myform.Action.value='DoExport';"">"
    Response.Write "                  <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "                  <input name='Action' type='hidden' id='Action' value='Export'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

End Sub

'=================================================
'过程名：DoExport
'作  用：导出模板处理
'=================================================
Sub DoExport()
    On Error Resume Next
    Dim mdbname, tconn, trs, strSql, Table_PE_lable
    Dim TemplateID, rs, sql, FormatConn, rsLabel
    TemplateID = Trim(Request("TemplateID"))
    FormatConn = Request.Form("FormatConn")
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If

    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导出的模板</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导出模板数据库名"
    End If
        
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        Set tconn = Nothing
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    Table_PE_lable = True
    tconn.Execute ("select LabelName from PE_Label")

    If Err Then
        Table_PE_lable = False
    End If

    '判断PE_Label 表是否存在
    If Table_PE_lable = False Then
        strSql = "        CREATE TABLE PE_Label  ("
        strSql = strSql & "  LabelID counter PRIMARY KEY,"
        strSql = strSql & "  LabelName text(50),"
        strSql = strSql & "  LabelClass text(50),"
        strSql = strSql & "  PageNum int,"
        strSql = strSql & "  LabelType int,"
        strSql = strSql & "  reFlashTime int,"
        strSql = strSql & "  fieldlist text(50),"
        strSql = strSql & "  LabelIntro text(255),"
        strSql = strSql & "  Priority int,"
        strSql = strSql & "  LabelContent Memo,"
        strSql = strSql & "  AreaCollectionID int"
        strSql = strSql & " )"
        Set trs = tconn.Execute(strSql)
        Set trs = Nothing
    End If
      
    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_Template")
        tconn.Execute ("delete from PE_Label")
    End If

    Set rs = Conn.Execute("select t.ChannelID,t.TemplateID,t.TemplateName,t.TemplateType,t.TemplateContent,t.IsDefault,t.ProjectName,c.ModuleType from PE_Template t left join PE_Channel c on t.ChannelID=c.ChannelID where t.TemplateID in (" & TemplateID & ")  order by t.TemplateID")
 
    Dim i, iVersion
    iVersion = 4
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Template", tconn, 1, 3

    For i = 0 To trs.Fields.Count - 1
        If LCase(trs.Fields(i).name) = "moduletype" Then
            iVersion = 5
            Exit For
        End If
    Next

    If iVersion = 4 Then
        trs.Close
        tconn.Execute ("alter table [PE_Template]  add COLUMN ModuleType int null")
        trs.Open "select * from PE_Template", tconn, 1, 3
    End If

    Do While Not rs.EOF
        trs.addnew
        trs("TemplateID") = rs("TemplateID")
        trs("ChannelID") = rs("ChannelID")

        If rs("ModuleType") <> "" And Not IsNull(rs("ModuleType")) Then
            trs("ModuleType") = rs("ModuleType")
        Else
            trs("ModuleType") = 0
        End If

        trs("TemplateName") = rs("TemplateName")
        trs("TemplateType") = rs("TemplateType")
        trs("TemplateContent") = rs("TemplateContent")
        trs("IsDefault") = rs("IsDefault")
        trs.Update
        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    '这是导出标签
    Set trs = Conn.Execute("select * from PE_Label")
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open "select * from PE_Label", tconn, 1, 3
    
    If Not trs.EOF Then
        Do While Not trs.EOF
            rs.addnew
            rs("LabelName") = trs("LabelName")
            rs("LabelClass") = trs("LabelClass")
            rs("LabelType") = trs("LabelType")
            rs("PageNum") = trs("PageNum")
            rs("reFlashTime") = trs("reFlashTime")
            rs("fieldlist") = trs("fieldlist")
            rs("LabelIntro") = trs("LabelIntro")
            rs("Priority") = trs("Priority")
            rs("LabelContent") = trs("LabelContent")
            rs("AreaCollectionID") = trs("AreaCollectionID")
            rs.Update
            trs.MoveNext
        Loop
    End If

    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("已经成功将所选中的模板设置导出到指定的数据库中！", ComeUrl)
End Sub

'=================================================
'过程名：Import
'作  用：导入模板第一步
'=================================================
Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>模板导入（第一步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的模板数据库的文件名： "
    Response.Write "        <input name='TemplateMdb' type='text' id='TemplateMdb' value='../temp/Template.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'>"
    Response.Write "        <input name='ProjectName' type='hidden' id='Action' value='" & ProjectName & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：Import2
'作  用：导入模板第二步
'=================================================
Sub Import2()
    On Error Resume Next

    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    Dim ModuleType, ChannelName
    
    '获得下拉频道参数 999999 表示所有
    ModuleType = Trim(Request.Form("ModuleType"))

    If ModuleType = "" Then
        ModuleType = 999999
    Else
        ModuleType = PE_CLng(ModuleType)
    End If
    
    '获得导入模板数据库路径
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("Templatemdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "＄", "/") '防止外部链接安全问题

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入模板数据库名"
        Exit Sub
    End If

    '建立导入模板数据库
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If

    Dim i, iVersion
    iVersion = 4
    Set trs = tconn.Execute("select top 1 * from PE_Template")

    For i = 0 To trs.Fields.Count - 1

        If LCase(trs.Fields(i).name) = "moduletype" Then
            iVersion = 5
            Exit For
        End If

    Next

    Set trs = Nothing
    
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>模板导入（第二步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>将要导入的模板</strong></td>"
    Response.Write "            <td></td>"
    Response.Write "            <td><strong>要导入到那个频道</strong></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td colspan='3'height='25'>"

    If iVersion = 4 Then
        Response.Write "请选择频道："
    Else
        Response.Write "请选择模块："
    End If

    Response.Write "              <select name='ModuleType' onChange='document.myform.submit();'>"

    If iVersion = 4 Then
        Call GetAllModule("4.03", ModuleType)
    Else
        Call GetAllModule("5.0", ModuleType)
    End If

    Response.Write "              </select>"
    Response.Write "              <br>"
    Response.Write "             </td>"
    Response.Write "           </tr>"
    Response.Write "           <tr>"
    Response.Write "            <td>"
    
    '显示模板
    Response.Write "              <select name='TemplateID' size='2' multiple style='height:300px;width:250px;'>"
    
    '当导入模板为4.03型时
    If iVersion = 4 Then
        '查询选择是 指定 还是 用户自定义（－2） 还是 首页（0） 还是全部
        If ModuleType <> 999999 And ModuleType <> -2 Then
            sql = "select * from PE_Template where ChannelID = " & ModuleType & " Order by ChannelID asc"
        ElseIf ModuleType = -2 Then
            sql = "select ChannelID,TemplateID,TemplateName from PE_Template where ChannelID not in (0,1,2,3)"
        ElseIf ModuleType = 0 Then
            sql = "select ChannelID,TemplateID,TemplateName from PE_Template where ChannelID=0"
        Else
            sql = "select * from PE_Template"
        End If
    Else
        '5.0  查询选择是 指定 还是 首页（0） 还是全部
        If ModuleType <> 999999 And ModuleType <> -1 Then
            sql = "select * from PE_Template where ModuleType = " & ModuleType & " Order by ChannelID asc"
        ElseIf ModuleType = 0 Then
            sql = "select ChannelID,TemplateID,TemplateName from PE_Template where ChannelID=0"
        Else
            sql = "select * from PE_Template"
        End If
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1

    If rs.BOF And rs.EOF Then
        '没有模板时指定关闭提交按钮
        Response.Write "                <option value='0'>没有任何模板</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"
    Response.Write "                  </td>"
    Response.Write "                  <td width='80'>&nbsp;&nbsp;导入到&gt;&gt;</td>"
    Response.Write "                  <td>"
    Response.Write "                    <select name='TargetChannelID' size='2' multiple style='height:300px;width:250px;'"

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "                             >"
    
    If CLng(ModuleType) = 0 Then
        Response.Write "               <option value='0'>通用模板</option>" & vbCrLf
    Else
        '是所有，还是首页，还是指定模块
        If CLng(ModuleType) = 999999 Or CLng(ModuleType) = -2 Then
            sql = "select ChannelID,ChannelName from PE_Channel where ChannelType < 2 and Disabled=" & PE_False
        Else
            sql = "select ChannelID,ChannelName,ModuleType from PE_Channel where  ChannelType < 2 and Disabled=" & PE_False & " and ModuleType=" & ModuleType & "  Order by ChannelID asc"
        End If

        Set rs = Conn.Execute(sql)

        If rs.BOF And rs.EOF Then
            Response.Write "              <option value='0'>您还没有建立此类型的频道</option>"
        Else
            If CLng(ModuleType) = 999999 Then
                Response.Write "            <option value='0'>通用模板</option>" & vbCrLf
            End If

            Do While Not rs.EOF
                Response.Write "           <option value='" & rs("ChannelID") & "'>" & rs("ChannelName") & "</option>"
                rs.MoveNext
            Loop
        End If

        rs.Close
        Set rs = Nothing
    End If
    
    Response.Write "                    </select>"
    Response.Write "                   </td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='10'></td></tr>"
    Response.Write "                  <tr>"
    Response.Write "                    <td height='25' align='center'><b> 提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "                    <td height='25' align='center'></td>"
    Response.Write "                    <td height='25' align='center'><b> 提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='20'></td></tr>"
    Response.Write "                  <tr><td colspan='3' height='25' align='center'><input type='submit' name='Submit' value=' 导入模板 ' onClick=""document.myform.Action.value='DoImport';"""

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "                      ></td>"
    Response.Write "                 </tr>"
    Response.Write "               </table>"
    Response.Write "               <input name='TemplateMdb' type='hidden' id='TemplateMdb' value='" & mdbname & "'>"
    Response.Write "               <input name='Action' type='hidden' id='Action' value='Import2'>"
    Response.Write "               <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "               <br>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "       </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：DoImport
'作  用：导入模板保存
'=================================================
Sub DoImport()
    On Error Resume Next
    
    Dim trs, crs, mdbname, tconn
    Dim TemplateID, TargetChannelID, rs, sql, rsLabel, Table_PE_lable
    TemplateID = Trim(Request("TemplateID"))
    TargetChannelID = Trim(Request("TargetChannelID"))
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If
    If IsValidID(TargetChannelID) = False Then
        TargetChannelID = ""
    End If

    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导入的模板</li>"
    End If

    If TargetChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导入的频道模块</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导出模板数据库名"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
        
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    If FoundInArr(TargetChannelID, 0, ",") = True Then
      
        Set trs = tconn.Execute(" select * from PE_Template where TemplateID in (" & TemplateID & ") and ChannelID=0 order by TemplateID")
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open "select top 1 * from PE_Template", Conn, 1, 3

        Do While Not trs.EOF
            rs.addnew
            rs("ChannelID") = 0
            rs("TemplateName") = trs("TemplateName")
            rs("TemplateType") = trs("TemplateType")
            rs("TemplateContent") = trs("TemplateContent")
            rs("ProjectName") = ProjectName
            rs("IsDefault") = False
            rs("IsDefaultInProject") = False
            rs("Deleted") = False
            rs.Update
            trs.MoveNext
        Loop
    
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If
       
    Dim i, iVersion
    iVersion = 4
    Set crs = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelID in (" & TargetChannelID & ")")
    Set trs = tconn.Execute(" select * from PE_Template where TemplateID in (" & TemplateID & ")  order by TemplateID")

    For i = 0 To trs.Fields.Count - 1

        If LCase(trs.Fields(i).name) = "moduletype" Then
            iVersion = 5
            Exit For
        End If

    Next

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open "select top 1 * from PE_Template", Conn, 1, 3
    
    Do While Not crs.EOF
        trs.MoveFirst
        If Not trs.EOF Then
            Do While Not trs.EOF
                If iVersion = 5 Then   '如果是5.0版的模板数据库
                    If trs("ModuleType") = crs("ModuleType") Then
                        rs.addnew
                        rs("ChannelID") = crs("ChannelID")
                        rs("TemplateName") = trs("TemplateName")
                        rs("TemplateType") = trs("TemplateType")
                        rs("TemplateContent") = trs("TemplateContent")
                        rs("ProjectName") = ProjectName
                        rs("IsDefault") = False
                        rs("IsDefaultInProject") = False
                        rs("Deleted") = False
                        rs.Update
                    End If
                Else  '如果是4.0版的模板数据库
                    rs.addnew
                    rs("ChannelID") = crs("ChannelID")
                    rs("TemplateName") = trs("TemplateName")
                    rs("TemplateType") = trs("TemplateType")
                    rs("TemplateContent") = trs("TemplateContent")
                    rs("ProjectName") = ProjectName
                    rs("IsDefault") = False
                    rs("IsDefaultInProject") = False
                    rs("Deleted") = False
                    rs.Update
                End If
                trs.MoveNext
            Loop
        End If
        crs.MoveNext
    Loop
    Set crs = Nothing
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    '判断PE_Label 表是否存在
    Table_PE_lable = True
    tconn.Execute ("select LabelName from PE_Label")

    If Err Then
        Table_PE_lable = False
    End If

    If Table_PE_lable = True Then
        '这是导入标签
        Set rsLabel = tconn.Execute("select * from PE_Label")
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Label", Conn, 1, 3
        
        Do While Not rsLabel.EOF
            Set rs = Conn.Execute("select LabelName from PE_Label where LabelName='" & rsLabel("LabelName") & "'")

            If rs.BOF And rs.EOF Then
                trs.addnew
                trs("LabelName") = rsLabel("LabelName")
                trs("LabelType") = rsLabel("LabelType")
                trs("LabelIntro") = rsLabel("LabelIntro")
                trs("Priority") = rsLabel("Priority")
                trs("LabelContent") = rsLabel("LabelContent")
                trs.Update
            End If

            rsLabel.MoveNext
        Loop
        
        Set trs = Nothing
        Set rsLabel = Nothing
        rs.Close
        Set rs = Nothing
    End If
    
    tconn.Close
    Set tconn = Nothing
       
    Call WriteSuccessMsg("已经成功从指定的数据库中导入选中的模板！", ComeUrl & "?ChannelID=" & ChannelID & "&TempType=" & TempType & "&Action=Import2&Templatemdb=" & Replace(mdbname, "/", "＄") & "")

End Sub

'=================================================
'过程名：ChannelCopyTemplate
'作  用：频道复制模板
'=================================================
Sub ChannelCopyTemplate()

    Dim rs, sql
    Dim trs, iCount
    Dim TemplateChannelID, ModuleType
    
    '获得模板频道ID
    TemplateChannelID = Trim(Request.Form("TemplateChannelID"))
    If IsValidID(TemplateChannelID) = False Then
        TemplateChannelID = ""
    End If

    If TemplateChannelID = "" Then
        TemplateChannelID = 999999
    End If
    
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>方案频道模板复制</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>选择要复制的频道模板</strong></td>"
    Response.Write "            <td></td>"
    Response.Write "            <td><strong>要复制到那个频道</strong></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td colspan='3'height='25'>"
    '显示下拉系统已有的频道
    sql = "select ChannelID,ChannelName,ModuleType from PE_Channel where ModuleType <> 0"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3
    Response.Write "             <select name='TemplateChannelID' onChange='document.myform.submit();'>"

    If rs.BOF And rs.EOF Then
        Response.Write "           <option value=''>您还没有建立频道</option>"
    Else

        Do While Not rs.EOF
            Response.Write "       <option value='" & rs("ChannelID") & "'"

            If rs("ChannelID") = PE_CLng(TemplateChannelID) Then Response.Write " selected"
            Response.Write ">" & rs("ChannelName") & "</option>" & vbCrLf
            rs.MoveNext
        Loop

        Response.Write "           <option value='999999'"

        If PE_CLng(TemplateChannelID) = 999999 Then Response.Write " selected"
        Response.Write ">所有频道</option>" & vbCrLf
    End If

    Response.Write "         </select>"
    rs.Close
    Set rs = Nothing
    
    Response.Write "             <br>"
    Response.Write "            </td>"
    Response.Write "           </tr>"
    Response.Write "           <tr>"
    Response.Write "             <td>"
    Response.Write "               <select name='TemplateID' size='2' multiple style='height:300px;width:250px;'>"
           
    '判断是所有还是指定
    If PE_CLng(TemplateChannelID) = 999999 Then
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,c.ModuleType from PE_Template t inner join PE_Channel c on t.ChannelID=c.ChannelID where t.ProjectName='" & ProjectName & "'"
    Else
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,c.ModuleType from PE_Template t inner join PE_Channel c on t.ChannelID=c.ChannelID where t.ChannelID=" & TemplateChannelID & " And t.ProjectName='" & ProjectName & "'  Order by t.ChannelID asc"
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "      <option value=''>没有任何模板</option>"
        '关闭提交按钮
        iCount = 0
        '-999999 为 没有模板
        ModuleType = -999999
    Else
        '得到值证明有内容开起提交按钮
        iCount = rs.RecordCount
        '得到模块的类型
        ModuleType = rs("ModuleType")

        Do While Not rs.EOF
            Response.Write "  <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "               </select>"
    Response.Write "              </td>"
    Response.Write "              <td width='80'>&nbsp;&nbsp;复制到&gt;&gt;></td>"
    Response.Write "              <td>"
    Response.Write "                <select name='TargetChannelID' size='2' multiple style='height:300px;width:250px;'>"

    '判没有模板时
    If ModuleType = -999999 Then
        Response.Write "              <option value=''>没有可复制的模板</option>" & vbCrLf
    Else

        '判断是全部 还是 指定
        If PE_CLng(TemplateChannelID) = 999999 Then
            sql = "select ChannelID,ChannelName from PE_Channel where ModuleType<>0"
        Else
            sql = "select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelID<>" & TemplateChannelID & " and ModuleType=" & ModuleType & "  Order by ChannelID asc"
        End If

        Set rs = Conn.Execute(sql)

        If rs.BOF And rs.EOF Then
            iCount = 0
            Response.Write "          <option value=''>您还没有建立相同类型的频道</option>"
        Else

            Do While Not rs.EOF
                Response.Write "      <option value='" & rs("ChannelID") & "'>" & rs("ChannelName") & "</option>"
                rs.MoveNext
            Loop

        End If

        rs.Close
        Set rs = Nothing
    End If

    Response.Write "                </select>"
    Response.Write "               </td>"
    Response.Write "             </tr>"
    Response.Write "             <tr><td colspan='3' height='10'></td></tr>"
    Response.Write "             <tr>"
    Response.Write "              <td height='25' align='center'><b> 提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "              <td height='25' align='center'></td>"
    Response.Write "              <td height='25' align='center'><b> 提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "             </tr>"
    Response.Write "             <tr><td colspan='3' height='20'></td></tr>"
    Response.Write "             <tr><td colspan='3' height='25' align='center'><input type='submit' name='Submit' value=' 复制模板 ' onClick=""document.myform.Action.value='DoCopy';"""

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "                  ></td>"
    Response.Write "             </tr>"
    Response.Write "           </table>"
    Response.Write "           <input name='Action' type='hidden' id='Action' value='ChannelCopyTemplate'>"
    Response.Write "          <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "          <br>"
    Response.Write "        </td>"
    Response.Write "      </tr>"
    Response.Write "   </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：DoCopy
'作  用：频道复制模板保存
'=================================================
Sub DoCopy()
    ' On Error Resume Next
    Dim trs, crs
    Dim TemplateID, TargetChannelID, rs, sql
    TemplateID = Trim(Request("TemplateID"))
    TargetChannelID = Trim(Request("TargetChannelID"))
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If
    If IsValidID(TargetChannelID) = False Then
        TargetChannelID = ""
    End If

    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要复制的模板</li>"
    End If

    If TargetChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定目标频道</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If

    Set crs = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelID in (" & TargetChannelID & ")")
    Set trs = Conn.Execute("select T.ProjectName,T.TemplateName,T.TemplateType,T.TemplateContent,T.ChannelID,C.ChannelName,C.ModuleType from PE_Template T inner join PE_Channel c on T.ChannelID=C.ChannelID where T.TemplateID in (" & TemplateID & ")  order by T.TemplateID")

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open "select top 1 * from PE_Template", Conn, 1, 3
    
    Do While Not crs.EOF

        If Not trs.EOF Then

            Do While Not trs.EOF

                If trs("ChannelID") <> crs("ChannelID") And trs("ModuleType") = crs("ModuleType") Then
                    rs.addnew
                    rs("ChannelID") = crs("ChannelID")
                    rs("TemplateName") = trs("TemplateName")
                    rs("TemplateType") = trs("TemplateType")
                    rs("TemplateContent") = trs("TemplateContent")
                    rs("IsDefault") = False
                    rs("ProjectName") = trs("ProjectName")
                    rs("IsDefaultInProject") = False
                    rs("Deleted") = False
                    rs.Update
                End If

                trs.MoveNext
            Loop

            trs.MoveFirst
        End If

        crs.MoveNext
    Loop

    Set crs = Nothing
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    Call WriteSuccessMsg("已经成功完成了模板复制！", ComeUrl & "?ChannelID=" & ChannelID & "&TempType=" & TempType & "&Action=ChannelCopyTemplate")
End Sub

'=================================================
'过程名：DoTemplateCopy
'作  用：模板复制处理
'=================================================
Sub DoTemplateCopy()
    Dim sql, rs, trs, TemplateID, TemplateName, FoundErr, ErrMsg
    FoundErr = False

    TemplateID = Trim(Request("TemplateID"))
    TemplateName = Trim(Request("TemplateName"))
    ProjectName = Trim(Request("ProjectName"))
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If

    
    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数错误，项目的ID不对！</li>"
    End If
    
    If FoundErr <> True Then
        If InStr(TemplateID, ",") = 0 Then
            Set trs = Conn.Execute("Select TemplateID,ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,Deleted,IsDefaultInProject from PE_Template Where TemplateID=" & TemplateID)
        Else
            Set trs = Conn.Execute("Select TemplateID,ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,Deleted,IsDefaultInProject from PE_Template Where TemplateID in (" & TemplateID & ")")
        End If

        If trs.BOF And trs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br>参数错误，没有找到该模板项目！"
        Else
            Set rs = Server.CreateObject("adodb.recordset")
            rs.Open "select top 1 * from PE_Template", Conn, 1, 3

            Do While Not trs.EOF
                rs.addnew
                rs("ChannelID") = trs("ChannelID")
                rs("TemplateName") = trs("TemplateName") & " 备份"
                rs("TemplateType") = trs("TemplateType")
                rs("TemplateContent") = trs("TemplateContent")
                rs("IsDefault") = False
                rs("ProjectName") = trs("ProjectName")
                rs("IsDefaultInProject") = False
                rs("Deleted") = trs("Deleted")
                rs.Update
                ErrMsg = ErrMsg & "<br>新的模板保存为：<font color=red>" & rs("TemplateName") & "</font>"
                trs.MoveNext
            Loop

            rs.Close
            Set rs = Nothing
        End If

        Set trs = Nothing
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    Else
        Response.Write "<br>"
        Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
        Response.Write "  <tr align='center' class='title'><td height='22'><strong>恭喜您！</strong></td></tr>" & vbCrLf
        Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & TemplateName & " 模板备份完成." & ErrMsg & "<br></td></tr>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg'><td>"
        Response.Write "</td></tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
      '  Response.Write "<meta http-equiv='refresh' content=3;url='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=" & TemplateType & "&TempType=" & TempType & "&ProjectName=" & ProjectName & "&TemplateProjectID=" & TemplateProjectID & "'>"
        Call Refresh("Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=" & TemplateType & "&TempType=" & TempType & "&ProjectName=" & ProjectName & "&TemplateProjectID=" & TemplateProjectID,3)
    End If

    Call CloseConn
End Sub

'=================================================
'过程名：BatchReplace
'作  用：批量替换
'=================================================
Sub BatchReplace()

    Dim rs, sql
    Dim ModuleType, TemplateID, TemplateChannelID
    Dim BatchType, BatchContent, TemplateType, TemplateReplace, TemplateReplaceStart, TemplateReplaceEnd, TemplateReplaceResult
    Dim ProjectName

    TemplateType = PE_CLng(Trim(Request.Form("TemplateType")))
    TemplateID = ReplaceBadChar(Trim(Request.Form("TemplateID")))
    TemplateChannelID = ReplaceBadChar(Trim(Request.Form("TemplateChannelID")))
    BatchType = PE_CLng(Trim(Request.Form("BatchType")))
    BatchContent = PE_CLng(Trim(Request.Form("BatchContent")))
    TemplateReplace = Trim(Request.Form("TemplateReplace"))
    TemplateReplaceStart = Trim(Request.Form("TemplateReplaceStart"))
    TemplateReplaceEnd = Trim(Request.Form("TemplateReplaceEnd"))
    TemplateReplaceResult = Trim(Request.Form("TemplateReplaceResult"))

    If ReplaceBadChar(Trim(Request.QueryString("ProjectName"))) = "" Then
        ProjectName = ReplaceBadChar(Trim(Request.Form("ProjectName")))
    Else
        ProjectName = ReplaceBadChar(Trim(Request.QueryString("ProjectName")))
    End If

    If TemplateType = 0 Then
        TemplateType = 1
    End If

    If BatchType = 0 Then
        BatchType = 1
    End If
        
    '获得下拉频道参数 999999 表示所有
    ModuleType = Trim(Request.Form("ModuleType"))

    If ModuleType = "" Then
        ModuleType = 999999
    Else
        ModuleType = PE_CLng(ModuleType)
    End If

    Response.Write "<form method=""post"" action=""Admin_Template.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td height='22' colspan='2' align='center'><b>模板批量替换管理</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td class='tdbg' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "          <tr class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left' ><INPUT TYPE='radio' NAME='TemplateType' value='1' " & IsRadioChecked(TemplateType, 1) & ""
    Response.Write " onClick=""javascript:TemplateID.style.display='';TemplateChannelID.style.display='none';ProjectName.style.display='none';ModuleType.style.display='none';BatchTemplateID.style.display='none';"""
    Response.Write " ><b>选择要被替换<FONT color='red'>模板</FONT>ID</b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "          <td align='left' ><INPUT TYPE='radio' NAME='TemplateType' value='2' " & IsRadioChecked(TemplateType, 2) & ""
    Response.Write " onClick=""javascript:TemplateID.style.display='none';TemplateChannelID.style.display='';ProjectName.style.display='none';ModuleType.style.display='none';BatchTemplateID.style.display='none';"""
    Response.Write " ><b>选择要被替换<FONT color='blue'>频道</FONT>ID</b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left'><INPUT TYPE='radio' NAME='TemplateType' value='3' " & IsRadioChecked(TemplateType, 3) & ""
    Response.Write " onClick=""javascript:TemplateID.style.display='none';TemplateChannelID.style.display='none';ProjectName.style.display=''; ModuleType.style.display='';BatchTemplateID.style.display='';"""
    Response.Write "><b>选择要被替换的<FONT color='#339900'>方案</Font></b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left' ><INPUT TYPE='Text' NAME='TemplateID' id='TemplateID' value='" & TemplateID & " ' size='40' " & IsStyleDisplay(TemplateType, 1) & "></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left'>" & vbCrLf
    sql = "SELECT DISTINCT t.ChannelID, c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID"
    Set rs = Conn.Execute(sql)
    Response.Write "<select name='TemplateChannelID' id='TemplateChannelID' size='2' multiple style='height:300px;width:250px;'  " & IsStyleDisplay(TemplateType, 2) & ">"

    If rs.BOF And rs.EOF Then
        Response.Write "<option value="" selected>还没有添加频道！</option> "
    Else
        Response.Write "<option selected value=" & rs("ChannelID") & ">" & rs("ChannelName") & "</option>"
        rs.MoveNext

        Do While Not rs.EOF
            Response.Write "<option value=" & rs("ChannelID") & ">" & rs("ChannelName") & "</option>"
            rs.MoveNext
        Loop

    End If

    Response.Write "</select>"
    rs.Close
    Set rs = Nothing
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf

    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td>" & vbCrLf
    Response.Write "            <select name='ProjectName' id='ProjectName' style='width:150px;'  onChange='document.form1.submit();' " & IsStyleDisplay(TemplateType, 3) & ">"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何方案项目</option>"
    Else
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "            </select>"
    Response.Write "            <br>"
    '显示下拉系统已有的频道
    Response.Write "              <select name='ModuleType' id='ModuleType' onChange='document.form1.submit();'  " & IsStyleDisplay(TemplateType, 3) & ">"
    Call GetAllModule("5.0", ModuleType)
    Response.Write "              </select>"
    Response.Write "             <br>"
    '显示模板
    Response.Write "              <select name='BatchTemplateID' id='BatchTemplateID' size='2' multiple style='height:300px;width:250px;'  " & IsStyleDisplay(TemplateType, 3) & ">"

    '5.0  查询选择是 指定 还是 首页（0） 还是全部
    If ModuleType <> 999999 And ModuleType <> -1 And ModuleType <> 0 Then
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,t.TemplateType,t.TemplateContent,t.IsDefault,c.ModuleType,t.ProjectName from PE_Template t left join PE_Channel c on t.ChannelID=c.ChannelID where c.ModuleType=" & ModuleType & " And t.ProjectName='" & ProjectName & "' order by t.TemplateID"
    ElseIf ModuleType = 0 Then
        sql = "select ChannelID,TemplateID,TemplateName,ProjectName from PE_Template where ChannelID=0  And ProjectName='" & ProjectName & "'"
    Else
        sql = "select * from PE_Template  Where ProjectName='" & ProjectName & "'"
        If FoundInArr(AllModules, "Supply", ",") = False Then
            sql = sql & " And ChannelID <> 999 "
        End If
        If FoundInArr(AllModules, "House", ",") = False Then
            sql = sql & " And ChannelID <> 998 "
        End If
        If FoundInArr(AllModules, "Job", ",") = False Then
            sql = sql & " And ChannelID <> 997 "
        End If
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        '没有模板时指定关闭提交按钮
        Response.Write "                <option value='0'>没有任何模板</option>"
    Else
        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "              </select>"
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td valign='top'>" & vbCrLf
    Response.Write "       <table width='100%' height='400' border='0' cellpadding='0' cellspacing='1'>"
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right""><strong>替换内容：</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='BatchContent' value='1' " & IsRadioChecked(BatchContent, 1) & " >模板名称&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='BatchContent' value='2' " & IsRadioChecked(BatchContent, 2) & " >模板内容</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right""><strong>替换类型：</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='BatchType' value='1' onClick=""javascript:PE_TemplateReplaceStart.style.display='none';PE_TemplateReplaceEnd.style.display='none';PE_TemplateReplace.style.display='';"" " & IsRadioChecked(BatchType, 1) & ">简单替换&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='BatchType' value='2'  onClick=""javascript:PE_TemplateReplaceStart.style.display='';PE_TemplateReplaceEnd.style.display='';PE_TemplateReplace.style.display='none';"" " & IsRadioChecked(BatchType, 2) & ">高级替换</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplace' " & IsStyleDisplay(BatchType, 1) & "> " & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg"" align=""right""><strong> 要替换的代码：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplace' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplace & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplaceStart' " & IsStyleDisplay(BatchType, 2) & "> " & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg"" align=""right"" ><strong> 要替换的开始代码：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplaceStart' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplaceStart & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplaceEnd' " & IsStyleDisplay(BatchType, 2) & "> " & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg"" align=""right"" ><strong> 要替换的结束代码：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplaceEnd' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplaceEnd & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplaceResult'>" & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg""  align=""right""><strong> 要替换后的代码：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplaceResult' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplaceResult & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg""> " & vbCrLf
    Response.Write "           <td colspan=""2"" align=""center"" class=""tdbg"" height=""50"">" & vbCrLf
    Response.Write "            <input name=""Action"" type=""hidden"" id=""Action"" value=""BatchReplace"">" & vbCrLf
    Response.Write "            <input name=""ChannelID"" type=""hidden"" id=""ChannelID"" value=" & ChannelID & ">" & vbCrLf
    Response.Write "            <input name=""Cancel"" type=""button"" id=""Cancel"" value=""返回上一步"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input  type=""submit"" name=""Submit"" value="" 开始替换 "" onClick=""document.form1.Action.value='DoBatchReplace';"" >" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "       </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "     </tr>" & vbCrLf
    Response.Write " </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

'=================================================
'过程名：DoBatchReplace
'作  用：批量替换处理
'=================================================
Sub DoBatchReplace()
    Dim rs, sql
    Dim TemplateType, TemplateID, TemplateChannelID, BatchTemplateID
    Dim BatchType, BatchContent, BatchDataType, TemplateReplace, TemplateReplaceStart, TemplateReplaceEnd, TemplateReplaceResult
    Dim FoundErr, ErrMsg
        
    FoundErr = False
    TemplateType = PE_CLng(Trim(Request.Form("TemplateType")))
    TemplateID = ReplaceBadChar(Trim(Request.Form("TemplateID")))
    BatchTemplateID = Trim(Request.Form("BatchTemplateID"))
    TemplateChannelID = Trim(Request.Form("TemplateChannelID"))
    BatchType = PE_CLng(Trim(Request.Form("BatchType")))
    BatchContent = PE_CLng(Trim(Request.Form("BatchContent")))
    TemplateReplace = Trim(Request.Form("TemplateReplace"))
    TemplateReplaceStart = Trim(Request.Form("TemplateReplaceStart"))
    TemplateReplaceEnd = Trim(Request.Form("TemplateReplaceEnd"))
    TemplateReplaceResult = Trim(Request.Form("TemplateReplaceResult"))
    If IsValidID(BatchTemplateID) = False Then
        BatchTemplateID = ""
    End If
    If IsValidID(TemplateChannelID) = False Then
        TemplateChannelID = ""
    End If

    If TemplateType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择要替换的模板类型</li>"
    End If

    If BatchContent = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择模板内容类型</li>"
    End If

    If BatchType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择模板替换字符类型</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If TemplateType = 1 Then
        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>没有模板ID号,请返回输入要替换的模板ID</li>"
        End If

    ElseIf TemplateType = 2 Then

        If TemplateChannelID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>没有频道ID号,请返回输入要替换的模板频道ID</li>"
        End If

    ElseIf TemplateType = 3 Then

        If BatchTemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>没有模板ID号,请返回输入要替换的模板ID</li>"
        End If

        TemplateID = BatchTemplateID
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>选择的模板类型不对</li>"
    End If

    If BatchType = 1 Then
        If TemplateReplace = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入要替换的代码不能为空</li>"
        End If

    ElseIf BatchType = 2 Then

        If TemplateReplaceStart = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入要替换的开始代码不能为空</li>"
        End If

        If TemplateReplaceEnd = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入要替换后的结束代码不能为空</li>"
        End If

    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>选择模板替换字符类型不对</li>"
    End If

    If TemplateReplaceResult = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>输入要替换后的代码不能为空</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<li>正在替换数据</li>&nbsp;&nbsp;"
    Set rs = Server.CreateObject("ADODB.Recordset")

    sql = "select TemplateID,ChannelID,TemplateName,TemplateContent from PE_Template where "

    If TemplateType = 1 Or TemplateType = 3 Then
        If InStr(TemplateID, ",") > 0 Then
            sql = sql & " TemplateID in (" & TemplateID & ")"
        Else
            sql = sql & " TemplateID=" & TemplateID
        End If

    ElseIf TemplateType = 2 Then

        If InStr(TemplateChannelID, ",") > 0 Then
            sql = sql & " ChannelID in (" & TemplateChannelID & ")"
        Else
            sql = sql & " ChannelID=" & TemplateChannelID
        End If
    End If

    rs.Open sql, Conn, 1, 3

    If BatchContent = 1 Then
        BatchDataType = "Name"
    Else
        BatchDataType = "Content"
    End If

    Do While Not rs.EOF
        If BatchType = 1 Then
            If InStr(rs("Template" & BatchDataType & ""), TemplateReplace) <> 0 Then
                rs("Template" & BatchDataType & "") = Replace(rs("Template" & BatchDataType & ""), TemplateReplace, TemplateReplaceResult)
                Response.Write "<br>&nbsp;&nbsp;" & rs("TemplateName") & "..<font color='#009900'>模板替换成功！</font>"
            Else
                Response.Write "<br>&nbsp;&nbsp;" & rs("TemplateName") & "..<font color='#FF0000'>模板替换代码不存在,不用替换！</font>"
            End If

        ElseIf BatchType = 2 Then
            rs("Template" & BatchDataType & "") = BatchReplaceString(rs("Template" & BatchDataType & ""), TemplateReplaceStart, TemplateReplaceEnd, TemplateReplaceResult, rs("TemplateName"))
        End If

        rs.Update
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Response.Write "<br><center> <a href='Admin_Template.asp?Action=Main' >返回模板管理</a> </center>"
End Sub

'=================================================
'函数名：StrJS_Template
'作  用：显示当前频道的模板类型
'=================================================
Sub StrJS_Template()
    Dim TrueSiteUrl
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
     
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "    addeditorcss=false;" & vbCrLf
    Response.Write "    addeditorcss2=false;" & vbCrLf
    Response.Write "    var strTemplateLabel;" & vbCrLf
    Response.Write "    var strTemplateLabel2;" & vbCrLf
    Response.Write "    function ResumeError() {" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    window.onerror = ResumeError;" & vbCrLf
    Response.Write "    function sizeContent(num,objname){" & vbCrLf
    Response.Write "        var obj = document.getElementById(objname);" & vbCrLf
    Response.Write "        if (parseInt(obj.rows)+num>=1) {" & vbCrLf
    Response.Write "            obj.rows = parseInt(obj.rows) + num;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (num>0){" & vbCrLf
    Response.Write "            obj.width=""90%"";" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function copy(num) {" & vbCrLf
    Response.Write "        if (num==1) {" & vbCrLf
    Response.Write "            var content= document.form1.Content.value;" & vbCrLf
    Response.Write "            document.form1.Content.value=content;" & vbCrLf
    Response.Write "            document.form1.Content.focus();" & vbCrLf
    Response.Write "            document.form1.Content.select();" & vbCrLf
    Response.Write "            textRange = document.form1.Content.createTextRange();" & vbCrLf
    Response.Write "            textRange.execCommand(""Copy"");" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else {" & vbCrLf
    Response.Write "            document.form1.Content2.focus();" & vbCrLf
    Response.Write "            document.form1.Content2.select();" & vbCrLf
    Response.Write "            textRange = document.form1.Content2.createTextRange();" & vbCrLf
    Response.Write "            textRange.execCommand(""Copy"");" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function LoadEditorAlgebra(num){" & vbCrLf
    Response.Write "        if (num==1){" & vbCrLf
    Response.Write "            document.form1.Content.rows=30;" & vbCrLf
    Response.Write "            document.form1.rollContent.rows=31;" & vbCrLf
    Response.Write "            showAlgebra.style.display="""";" & vbCrLf
    Response.Write "            showeditor.style.display=""none"";" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation1.style.display="""";" & vbCrLf
    Response.Write "            CommonLabel1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display="""";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen.disabled=true;" & vbCrLf
    Response.Write "            document.form1.Copy.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss==false){" & vbCrLf
    Response.Write "                addeditorcss=true;" & vbCrLf
    Response.Write "                setContent('get',1)" & vbCrLf
    Response.Write "                editor.yToolbarsCss();" & vbCrLf
    Response.Write "                editor.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('get',1)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            document.form1.Content2.rows=30;" & vbCrLf
    Response.Write "            document.form1.rollContent2.rows=31;" & vbCrLf
    Response.Write "            showAlgebra2.style.display="""";" & vbCrLf
    Response.Write "            showeditor2.style.display=""none"";" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation12.style.display="""";" & vbCrLf
    Response.Write "            CommonLabel2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display="""";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen2.disabled=true;" & vbCrLf
    Response.Write "            document.form1.Copy2.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss2==false){" & vbCrLf
    Response.Write "                addeditorcss2=true;" & vbCrLf
    Response.Write "                setContent('get',2);" & vbCrLf
    Response.Write "                editor2.yToolbarsCss();" & vbCrLf
    Response.Write "                editor2.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('get',2)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " function LoadEditorEdit(num){" & vbCrLf
    Response.Write "        if (num==1){" & vbCrLf
    Response.Write "            showAlgebra.style.display=""none"";" & vbCrLf
    Response.Write "            showeditor.style.display="""";" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy.disabled=true;" & vbCrLf
    Response.Write "            if (addeditorcss==false){" & vbCrLf
    Response.Write "                addeditorcss=true;" & vbCrLf
    Response.Write "                setContent('set',1);" & vbCrLf
    Response.Write "                editor.yToolbarsCss();" & vbCrLf
    Response.Write "                editor.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('set',1)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            showAlgebra2.style.display=""none"";" & vbCrLf
    Response.Write "            showeditor2.style.display="""";" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation12.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen2.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy2.disabled=true;" & vbCrLf
    Response.Write "            if (addeditorcss2==false){" & vbCrLf
    Response.Write "                addeditorcss2=true;" & vbCrLf
    Response.Write "                setContent('set',2);" & vbCrLf
    Response.Write "                editor2.yToolbarsCss();" & vbCrLf
    Response.Write "                editor2.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('set',2)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function LoadEditorMix(num){" & vbCrLf
    Response.Write "        if (num==1){" & vbCrLf
    Response.Write "            document.form1.Content.rows=10;" & vbCrLf
    Response.Write "            document.form1.rollContent.rows=11;" & vbCrLf
    Response.Write "            showeditor.style.display="""";" & vbCrLf
    Response.Write "            showAlgebra.style.display="""";" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss==false){" & vbCrLf
    Response.Write "                addeditorcss=true;" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "                editor.showBorders()" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            document.form1.Content2.rows=10;" & vbCrLf
    Response.Write "            document.form1.rollContent2.rows=11;" & vbCrLf
    Response.Write "            showAlgebra2.style.display="""";" & vbCrLf
    Response.Write "            showeditor2.style.display="""";" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation12.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen2.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy2.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss2==false){" & vbCrLf
    Response.Write "                addeditorcss2=true;" & vbCrLf
    Response.Write "                editor2.yToolbarsCss();" & vbCrLf
    Response.Write "                editor2.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function OpenNavigation(TemplateType) {" & vbCrLf
    Response.Write "        if (TemplateType==1){" & vbCrLf
    Response.Write "            showLabel.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation3.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='0,*';" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            showLabel2.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='0,*';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function CloseNavigation(TemplateType) {" & vbCrLf
    Response.Write "        if (TemplateType==1){" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='200,*';" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation4.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='200,*';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function setContent2(num) {" & vbCrLf
    Response.Write "    if (num==1){" & vbCrLf
    Response.Write "        form1.Content.focus();" & vbCrLf
    Response.Write "        strTemplateLabel = document.selection.createRange();" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        form1.Content2.focus();" & vbCrLf
    Response.Write "        strTemplateLabel2 = document.selection.createRange();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function insertTemplateLabel(strLabel,insertTemplateType) {" & vbCrLf
    Response.Write "    if (insertTemplateType==1){" & vbCrLf
    Response.Write "        strTemplateLabel.text = strLabel" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strTemplateLabel2.text = strLabel" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "function SuperFunctionLabel(url,label,title,ModuleType,ChannelShowType,iwidth,iheight,TemplateType){" & vbCrLf
    Response.Write "    if (TemplateType==1){" & vbCrLf
    Response.Write "        form1.Content.focus();" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        form1.Content2.focus();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var str = document.selection.createRange();" & vbCrLf
    Response.Write "    var arr = showModalDialog(url+""?ChannelID=" & ChannelID & "&Action=Add&LabelName=""+label+""&Title=""+title+""&ModuleType=""+ModuleType+""&ChannelShowType=""+ChannelShowType+""&InsertTemplate=1"", """", ""dialogWidth:""+iwidth+""px; dialogHeight:""+iheight+""px; help: no; scroll:yes; status: yes""); " & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        str.text = arr;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function fullscreen(num) {" & vbCrLf
    Response.Write "    window.open (""../Editor/editor_fullscreen.asp?ChannelID=" & ChannelID & "&num=""+num+"""", """", ""toolbar=no, menubar=no, top=0,left=0,width=1024,height=768, scrollbars=no, resizable=no,location=no, status=no"");" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function Templateskin(){" & vbCrLf
    Response.Write "    if(confirm('您确定要转入风格设计，如果您没有保存当前操作的模板请保存模板。')){" & vbCrLf
    Response.Write "        window.location.href='Admin_Skin.asp?Action=Modify&SkinID=1&IsDefault=-1';" & vbCrLf
    Response.Write "    }  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function show_ln(txt_ln,txt_main){" & vbCrLf
    Response.Write "    var txt_ln  = document.getElementById(txt_ln);" & vbCrLf
    Response.Write "    var txt_main  = document.getElementById(txt_main);" & vbCrLf
    Response.Write "    txt_ln.scrollTop = txt_main.scrollTop;" & vbCrLf
    Response.Write "    while(txt_ln.scrollTop != txt_main.scrollTop)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        txt_ln.value += (i++) + '\n';" & vbCrLf
    Response.Write "        txt_ln.scrollTop = txt_main.scrollTop;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    return;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function editTab(){" & vbCrLf
    Response.Write "    var code, sel, tmp, r" & vbCrLf
    Response.Write "    var tabs=''" & vbCrLf
    Response.Write "    event.returnValue = false" & vbCrLf
    Response.Write "    sel =event.srcElement.document.selection.createRange()" & vbCrLf
    Response.Write "    r = event.srcElement.createTextRange()" & vbCrLf
    Response.Write "    switch (event.keyCode){" & vbCrLf
    Response.Write "        case (8) :" & vbCrLf
    Response.Write "        if (!(sel.getClientRects().length > 1)){" & vbCrLf
    Response.Write "            event.returnValue = true" & vbCrLf
    Response.Write "            return" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        code = sel.text" & vbCrLf
    Response.Write "        tmp = sel.duplicate()" & vbCrLf
    Response.Write "        tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)" & vbCrLf
    Response.Write "        sel.setEndPoint('startToStart', tmp)" & vbCrLf
    Response.Write "        sel.text = sel.text.replace(/\t/gm, '')" & vbCrLf
    Response.Write "        code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')" & vbCrLf
    Response.Write "        r.findText(code)" & vbCrLf
    Response.Write "        r.select()" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "    case (9) :" & vbCrLf
    Response.Write "        if (sel.getClientRects().length > 1){" & vbCrLf
    Response.Write "            code = sel.text" & vbCrLf
    Response.Write "            tmp = sel.duplicate()" & vbCrLf
    Response.Write "            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)" & vbCrLf
    Response.Write "            sel.setEndPoint('startToStart', tmp)" & vbCrLf
    Response.Write "            sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')" & vbCrLf
    Response.Write "            code = code.replace(/\r\n/g, '\r\t')" & vbCrLf
    Response.Write "            r.findText(code)" & vbCrLf
    Response.Write "            r.select()" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            sel.text = '\t'" & vbCrLf
    Response.Write "            sel.select()" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "    case (13) :" & vbCrLf
    Response.Write "        tmp = sel.duplicate()" & vbCrLf
   ' Response.write "        tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)" & vbCrLf
   ' Response.write "        tmp.setEndPoint('endToEnd', sel)" & vbCrLf
    Response.Write "        for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'" & vbCrLf
    Response.Write "        sel.text = '\r\n'+tabs" & vbCrLf
    Response.Write "        sel.select()" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "    default  :" & vbCrLf
    Response.Write "        event.returnValue = true" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    '=================================================
    '作  用：这是客户端处理
    '=================================================
    Response.Write "<script language=""VBScript"">" & vbCrLf
    Response.Write "    Dim Strsave,Strsave2,addeditorcss3" & vbCrLf
    Response.Write "    Dim regEx, Match, Matches, StrBody,strTemp,strTemp2,strMatch,arrMatch,i" & vbCrLf
    Response.Write "    Dim Content1,Content2,Content3,Content4,TemplateContent,TemplateContent2,TemplateContent3,arrContent,EditorContent" & vbCrLf
    Response.Write "    Set regEx = New RegExp" & vbCrLf
    Response.Write "    regEx.IgnoreCase = True" & vbCrLf
    Response.Write "    regEx.Global = True" & vbCrLf
    Response.Write "    Strsave=""A""" & vbCrLf
    Response.Write "    Strsave2=""A""" & vbCrLf
    Response.Write "    Sub CheckForm(Num)" & vbCrLf
    Response.Write "        if document.form1.TemplateName.value="""" then" & vbCrLf
    Response.Write "            alert ""模板名称不能为空！""" & vbCrLf
    Response.Write "            document.form1.TemplateName.focus()" & vbCrLf
    Response.Write "            exit sub" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if document.form1.Content.value="""" then" & vbCrLf
    Response.Write "            alert ""模板主内容不能为空！""" & vbCrLf
    Response.Write "            editor.HtmlEdit.focus()" & vbCrLf
    Response.Write "            exit sub" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Num=2 then" & vbCrLf
    Response.Write "            if document.form1.Content2.value="""" then " & vbCrLf
    Response.Write "                alert ""小类模板主内容不能为空！""" & vbCrLf
    Response.Write "                exit sub" & vbCrLf
    Response.Write "            End if" & vbCrLf
    Response.Write "            if Strsave=""B"" then setContent ""get"",1" & vbCrLf
    Response.Write "            if Strsave2=""B"" then setContent ""get"",2" & vbCrLf
    Response.Write "            document.form1.EditorContent.value=""""" & vbCrLf
    Response.Write "            document.form1.EditorContent2.value=""""" & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            if Strsave=""B"" then setContent ""get"",1" & vbCrLf
    Response.Write "            document.form1.EditorContent.value=""""" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        form1.submit" & vbCrLf
    Response.Write "    End Sub" & vbCrLf

    Call Resumeblank
    '=================================================
    '作  用：客户端交替传值
    '=================================================
    Response.Write "    function setContent(zhi,TemplateType)" & vbCrLf
    Response.Write "    if zhi=""get"" then" & vbCrLf
    Response.Write "        if TemplateType=1 then" & vbCrLf
    Response.Write "            if Strsave=""A"" then Exit Function" & vbCrLf
    Response.Write "            Strsave=""A""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content.value" & vbCrLf
    Response.Write "            TemplateContent2= editor.HtmlEdit.document.body.innerHTML" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            if Strsave2=""A"" then Exit Function" & vbCrLf
    Response.Write "            Strsave2=""A""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content2.value" & vbCrLf
    Response.Write "            TemplateContent2= editor2.HtmlEdit.document.body.innerHTML" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if TemplateContent="""" then " & vbCrLf
    Response.Write "            alert ""您删除了代码框网页，请您务必填写网页 ！""" & vbCrLf
    Response.Write "            Exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body>"")=0 then" & vbCrLf
    Response.Write "            regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "            Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "            For Each Match In Matches" & vbCrLf
    Response.Write "                StrBody = Match.Value" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "            If StrBody = """"  Then" & vbCrLf
    Response.Write "                alert ""您加载的文本框没有包含 <body> 或您没有给body 参数这会使网页很难看,请最少给出 <body> ！""" & vbCrLf
    Response.Write "                Exit function" & vbCrLf
    Response.Write "            End If " & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            StrBody=""<body>"" " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "        if ubound(arrContent)=0 then " & vbCrLf
    Response.Write "           alert ""您加载的文本框没有包含 <body> 或您没有给body 参数这会使网页很难看,请最少给出 <body> ！""" & vbCrLf
    Response.Write "           exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "        Content2 = arrContent(1)" & vbCrLf
    Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\}['|""""]\>""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(TemplateContent2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            regEx.Pattern = ""\{\$(.*?)\}""" & vbCrLf
    Response.Write "            Set strTemp = regEx.Execute(replace(Match.Value,"" "",""""))" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
    Response.Write "                TemplateContent2 = Replace(TemplateContent2, Match.Value, ""<!--"" & strTemp2 & ""-->"")" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\$\>""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(TemplateContent2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            regEx.Pattern = ""\#\[(.*?)\]\#""" & vbCrLf
    Response.Write "            Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""&amp;"", ""&"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""&13;&10;"",vbCrLf)" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""&9;"",vbTab)" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""′"",""'"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
    Response.Write "                TemplateContent2 = Replace(TemplateContent2, Match.Value, strTemp2)" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        TemplateContent2=Resumeblank(TemplateContent2)" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2,""{$InstallDir}{$rsClass_ClassUrl}"",""{$rsClass_ClassUrl}"")" & vbCrLf
    Response.Write "        regEx.Pattern = ""\{\$InstallDir\}editor.asp(.[^\<]*?)\#""" & vbCrLf
    Response.Write "        TemplateContent2 = regEx.Replace(TemplateContent2, ""#"")" & vbCrLf
    Response.Write "        if TemplateType =1 then" & vbCrLf
    Response.Write "            document.form1.Content.value=Content1& vbCrLf &TemplateContent2& vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            document.form1.Content2.value=Content1 & vbCrLf &TemplateContent2& vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "    Else" & vbCrLf
    Response.Write "        if TemplateType =1 then    " & vbCrLf
    Response.Write "            if Strsave=""B"" then Exit Function" & vbCrLf
    Response.Write "            Strsave=""B""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content.value" & vbCrLf
    Response.Write "        else " & vbCrLf
    Response.Write "            if Strsave2=""B"" then Exit Function" & vbCrLf
    Response.Write "            Strsave2=""B""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content2.value" & vbCrLf
    Response.Write "        End if    " & vbCrLf
    Response.Write "        if TemplateContent="""" then " & vbCrLf
    Response.Write "            alert ""您删除了代码框网页，请您务必填写网页 ！""" & vbCrLf
    Response.Write "            Exit function" & vbCrLf
    Response.Write "           " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body>"")=0 then" & vbCrLf
    Response.Write "            regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "            Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "            For Each Match In Matches" & vbCrLf
    Response.Write "                StrBody = Match.Value" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "            If StrBody = """"  Then" & vbCrLf
    Response.Write "                alert ""您加载的文本框没有包含 <body> 或您没有给body 参数这会使网页很难看,请最少给出 <body> ！""" & vbCrLf
    Response.Write "                Exit function" & vbCrLf
    Response.Write "            End If " & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            StrBody=""<body>"" " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "        if ubound(arrContent)=0 then " & vbCrLf
    Response.Write "           alert ""您加载的文本框没有包含 <body> 或您没有给body 参数这会使网页很难看,请最少给出 <body> ！""" & vbCrLf
    Response.Write "           exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "        Content2 = arrContent(1)" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""<!--$"", ""{$--"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""$-->"", ""--}"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""<!--{$"", ""{$"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""}-->"", ""}"")" & vbCrLf
    Response.Write "        '图片替换JS" & vbCrLf
    Response.Write "        regEx.Pattern = ""(\<Script)([\s\S]*?)(\<\/Script\>)""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(Content2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            strTemp = Replace(Match.Value, ""<"", ""[!"")" & vbCrLf
    Response.Write "            strTemp = Replace(strTemp, "">"", ""!]"")" & vbCrLf
    Response.Write "            strTemp = Replace(strTemp, ""'"", ""′"")" & vbCrLf
    Response.Write "            strTemp = ""<IMG alt='#"" & strTemp & ""#' src=""""" & InstallDir & "editor/images/jscript.gif"""" border=0 $>""" & vbCrLf
    Response.Write "            Content2 = Replace(Content2, Match.Value, strTemp)" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        '图片替换超级标签" & vbCrLf
    Response.Write "        regEx.Pattern = ""(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct|\{\$GetPositionList|\{\$GetSearchResult)\((.*?)\)\}""" & vbCrLf
    Response.Write "        Content2 = regEx.Replace(Content2, ""<IMG src=""""" & InstallDir & "editor/images/label.gif"""" border=0 zzz='$1($2)}'>"")" & vbCrLf
    Response.Write "        regEx.Pattern = ""\{\$InstallDir\}""" & vbCrLf
    Response.Write "        Content2 = regEx.Replace(Content2,""http://" & TrueSiteUrl & InstallDir & """)" & vbCrLf
    Response.Write "        if TemplateType=1 then" & vbCrLf
    Response.Write "            editor.HtmlEdit.document.body.innerHTML=Content2" & vbCrLf
    Response.Write "            editor.showBorders()" & vbCrLf
    Response.Write "            editor.showBorders()" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            editor2.HtmlEdit.document.body.innerHTML=Content2" & vbCrLf
    Response.Write "            editor2.showBorders()" & vbCrLf
    Response.Write "            editor2.showBorders()" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "    End if" & vbCrLf
    Response.Write "    End function" & vbCrLf
    Response.Write "    function setstatus()" & vbCrLf '为323 版兼容editor.asp 无效过程
    Response.Write "    end function" & vbCrLf
    Response.Write "</script>" & vbCrLf
    
End Sub

'=================================================
'过程名：fullscreen
'作  用：全屏模式
'=================================================
Sub fullscreen()
    Dim TrueSiteUrl

    If ChannelID = 0 Then
        Response.Write "频道参数丢失！"
        Response.End
    End If
            
    Response.Write "<HTML>" & vbCrLf
    Response.Write "<HEAD>" & vbCrLf
    Response.Write "<TITLE>HtmlEdit - 全屏编辑</TITLE>" & vbCrLf
    Response.Write "<META http-equiv=Content-Type content=""text/html; charset=gb2312"">" & vbCrLf
    Response.Write "</HEAD>" & vbCrLf
    Response.Write "<body leftmargin=0 topmargin=0 onunload=""Minimize()"">" & vbCrLf
    Response.Write "<input type=""hidden"" id=""ContentFullScreen"" name=""ContentFullScreen"" value="""">" & vbCrLf
    Response.Write "<script language=VBScript>" & vbCrLf
    Response.Write "   Dim Matches, Match, arrContent, Content1, Content2,Content3,Content5" & vbCrLf
    Response.Write "   Dim strTemp, strTemp2, StrBody,TemplateContent" & vbCrLf
    Response.Write "   Set regEx = New RegExp" & vbCrLf

    If Request.QueryString("num") = 1 Then
        Response.Write "ContentFullScreen.value=opener.editor.HtmlEdit.document.body.innerHTML" & vbCrLf
        Response.Write "TemplateContent= opener.document.form1.Content.value" & vbCrLf
    Else
        Response.Write "ContentFullScreen.value =opener.editor2.HtmlEdit.document.body.innerHTML" & vbCrLf
        Response.Write "TemplateContent= opener.document.form1.Content2.value" & vbCrLf
    End If

    Response.Write "   ContentFullScreen.value =""<html><head><META http-equiv=Content-Type content=text/html; charset=gb2312><link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'></head><body leftmargin=0 topmargin=0 >"" & ContentFullScreen.value" & vbCrLf
    Response.Write "   document.Write ""<iframe ID='EditorFullScreen' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&TemplateType=3&tContentid=ContentFullScreen' frameborder='0' scrolling=no width='100%' HEIGHT='100%'></iframe>""" & vbCrLf
    
    Call Resumeblank

    Response.Write "Function Minimize()" & vbCrLf
    Response.Write "       regEx.IgnoreCase = True" & vbCrLf
    Response.Write "       regEx.Global = True" & vbCrLf
    Response.Write "       regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "       Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "         For Each Match In Matches" & vbCrLf
    Response.Write "            StrBody = Match.Value" & vbCrLf
    Response.Write "         Next" & vbCrLf
    Response.Write "         arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "         Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "         Content2 = arrContent(1)" & vbCrLf
    Response.Write "         Content5 = EditorFullScreen.HtmlEdit.document.Body.innerHTML" & vbCrLf
    Response.Write "         regEx.Pattern = ""\<IMG(.[^\<]*)\}['|""""]\>""" & vbCrLf
    Response.Write "         Set Matches = regEx.Execute(Content5)" & vbCrLf
    Response.Write "         For Each Match In Matches" & vbCrLf
    Response.Write "             regEx.Pattern = ""\{\$(.*?)\}""" & vbCrLf
    Response.Write "             Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "             For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
    Response.Write "                Content5 = Replace(Content5, Match.Value, ""<!--""&strTemp2&""-->"")" & vbCrLf
    Response.Write "             Next" & vbCrLf
    Response.Write "         Next" & vbCrLf
    Response.Write "         regEx.Pattern = ""\<IMG(.[^\<]*)\$\>""" & vbCrLf
    Response.Write "         Set Matches = regEx.Execute(Content5)" & vbCrLf
    Response.Write "         For Each Match In Matches" & vbCrLf
    Response.Write "         regEx.Pattern = ""\#(.*?)\#""" & vbCrLf
    Response.Write "         Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "               strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
    Response.Write "               strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
    Response.Write "               strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
    Response.Write "               Content5 = Replace(Content5, Match.Value, strTemp2)" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "         Next" & vbCrLf
    Response.Write "        Content5=Replace(Content5, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        Content5=Replace(Content5, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf
    
    If Request.QueryString("num") = 1 Then
        Response.Write "opener.editor.HtmlEdit.document.body.innerHTML=Resumeblank(EditorFullScreen.getHTML())" & vbCrLf
        Response.Write "opener.document.form1.Content.value=Content1& vbCrLf & Resumeblank(Content5) & vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
        Response.Write "opener.editor.showBorders()" & vbCrLf
        Response.Write "opener.editor.showBorders()" & vbCrLf
    Else
        Response.Write "opener.editor2.HtmlEdit.document.body.innerHTML=Resumeblank(EditorFullScreen.getHTML())" & vbCrLf
        Response.Write "opener.document.form1.Content2.value=Content1& vbCrLf & Resumeblank(Content5) & vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
        Response.Write "opener.editor2.showBorders()" & vbCrLf
        Response.Write "opener.editor2.showBorders()" & vbCrLf
        
    End If

    Response.Write "    Set regEx = Nothing" & vbCrLf
    Response.Write "End function" & vbCrLf
    Response.Write "function setstatus()" & vbCrLf '这两个兼容editor.asp多能调用
    Response.Write "End function" & vbCrLf
    Response.Write "function setContent(zhi,TemplateType)" & vbCrLf
    Response.Write "End function" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "   setTimeout(""EditorFullScreen.showBorders()"",2000);" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "</BODY>" & vbCrLf
    Response.Write "</HTML>" & vbCrLf
   
End Sub

'=================================================
'过程名：ShiftCharacter
'作  用：替换标签为图片显示
'参  数：要替换的数据    Content
'=================================================
Function ShiftCharacter(ByVal Content)

    Dim strTemp, StrBody, arrContent, ContentHead, arrMatch, strMatch, i, TrueSiteUrl
    
    '替换文件的注解函数符，解决不显示问题
    
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
    
    Content = Replace(Content, "<!--{$", "{$")
    Content = Replace(Content, "}-->", "}")


    regEx.Pattern = "(\<body\>)"
    Content = regEx.Replace(Content, "<body>")

    Set Matches = regEx.Execute(Content)
    For Each Match In Matches
        StrBody = Match.value
    Next

    If InStr(Content, "<body>") = 0 Then
        regEx.Pattern = "(\<body)(.[^\<]*)(\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            StrBody = Match.value
        Next
    Else
        StrBody = "<body>"
    End If
    
    arrContent = Split(Content, StrBody)
    
    If UBound(arrContent) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您加载的模板没有包含 <body> 或您没有给body 参数这会使网页很难看,请最少给出 <body> ！</li>"
        Exit Function
    End If
    
    ContentHead = arrContent(0) & StrBody
    Content = arrContent(1)
    
    '图片替换JS
    regEx.Pattern = "(\<Script)([\s\S]*?)(\<\/Script\>)"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        strTemp = Replace(Match.value, "<", "[!")
        strTemp = Replace(strTemp, ">", "!]")
        strTemp = Replace(strTemp, "'", """")
        strTemp = "<IMG alt='#" & strTemp & "#' src=""" & InstallDir & "editor/images/jscript.gif"" border=0 $>"
        Content = Replace(Content, Match.value, strTemp)
    Next
    
    '图片替换超级标签
    regEx.Pattern = "(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct|\{\$GetPositionList|\{\$GetSearchResult)\((.*?)\)\}"
    Content = regEx.Replace(Content, "<IMG src=""" & InstallDir & "editor/images/label.gif"" border=0 zzz='$1($2)}'>")
      
    Content = ContentHead & vbCrLf & Content
    '替换文件标签 转换为 css文件
    Content = Replace(Content, "{$Skin_CSS}", "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>")
    Content = Replace(Content, "{$MenuJS}", "<script language='JavaScript' type='text/JavaScript' src='" & InstallDir & "js/menu.js'></script>")
    '替换文件中的路径
    Content = Replace(Content, "[InstallDir_ChannelDir]", "http://" & TrueSiteUrl & InstallDir & ChannelDir & "/")
    Content = Replace(Content, "{$InstallDir}", "http://" & TrueSiteUrl & InstallDir)
    
    ShiftCharacter = Content
    
End Function

'=================================================
'过程名：ShiftCharacterSave
'作  用：替换标签为图片显示
'参  数：要替换的数据    Content
'=================================================
Function ShiftCharacterSave(Content)

    Dim NullBody
    Dim strTemp, strTemp2, Match2, strSiteUrl, strPhotoJs
    
    '将绝对地址转化为相对地址
    strSiteUrl = Request.ServerVariables("HTTP_REFERER")
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/") - 1)
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/"))
     
    '使用正则判断大类，小类文件中是否有 <body> 怕有些网友删除文本框中的<body>
    regEx.Pattern = "(\<body\>)"
    Set Matches = regEx.Execute(Content)
    For Each Match In Matches
        NullBody = Match.value
    Next

    If NullBody = "" Then
        regEx.Pattern = "(\<body)(.[^\<]*)(\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            NullBody = Match.value
        Next
    Else
        NullBody = "<body>"
    End If
    
    If NullBody = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您加载的模板没有包含  &lt;<font color=blue>body</font>&gt;  这在网页中是不允许的！</li>"
        ShiftCharacterSave = False
        Exit Function
    End If
      
    regEx.Pattern = "\<IMG(.[^\<]*)\}['|""]>"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        regEx.Pattern = "\{\$(.*?)\}"
        Set strTemp = regEx.Execute(Match.value)

        For Each Match2 In strTemp
            strTemp2 = Replace(Match2.value, "?", """")
            Content = Replace(Content, Match.value, "<!--" & strTemp2 & "-->")
        Next
    Next
    
    '处理图片JS标签代码
    regEx.Pattern = "\<IMG(.[^\<]*)\$\>"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        regEx.Pattern = "\#(.*?)\#"
        Set strTemp = regEx.Execute(Match.value)

        For Each Match2 In strTemp
            strTemp2 = Replace(Match2.value, "?", "?")
            strTemp2 = Replace(strTemp2, "&amp;", "&")
            strTemp2 = Replace(strTemp2, "&13;&10;", "vbCrLf")
            strTemp2 = Replace(strTemp2, "&9;", "vbTab")
            strTemp2 = Replace(strTemp2, "[!", "<")
            strTemp2 = Replace(strTemp2, "!]", ">")
            Content = Replace(Content, Match.value, strTemp2)
        Next
    Next

    '处理编辑器问题
    Content = Replace(Content, "{$InstallDir}{$rsClass_ClassUrl}", "{$rsClass_ClassUrl}")
    Content = Replace(Content, "{$InstallDir}{$ArticleUrl}", "{$ArticleUrl}")
    Content = Replace(Content, "{$InstallDir}{$SoftUrl}", "{$SoftUrl}")
    Content = Replace(Content, "{$InstallDir}{$PhotoUrl}", "{$PhotoUrl}")
    Content = Replace(Content, "{$InstallDir}{$ProductUrl}", "{$ProductUrl}")

    '解决编辑器过滤为标签值为空问题
    Content = Replace(Content, "{$InstallDir}", "[$InstallDir]")
    regEx.Pattern = "(\s)+(value|title|src|href)(\s)*\=(\s)*\{\$(.[^\<\{]*)\}"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        strTemp = Replace(Trim(Match.value), "{$", """{$")
        strTemp = Replace(strTemp, "}", "}""")
        Content = Replace(Content, Match.value, " " & strTemp)
    Next

    Content = Replace(Content, "[$InstallDir]", "{$InstallDir}")

    '解决正文页用户删除图片js 问题
    strPhotoJs = "<script language=""JavaScript"">" & vbCrLf
    strPhotoJs = strPhotoJs & "<!--" & vbCrLf
    strPhotoJs = strPhotoJs & "//改变图片大小" & vbCrLf
    strPhotoJs = strPhotoJs & "function resizepic(thispic)" & vbCrLf
    strPhotoJs = strPhotoJs & "{" & vbCrLf
    'strPhotoJs = strPhotoJs & "if(thispic.width>700) thispic.width=700;" & vbCrLf
    strPhotoJs = strPhotoJs & "  return true;" & vbCrLf
    strPhotoJs = strPhotoJs & "}" & vbCrLf
    strPhotoJs = strPhotoJs & "//无级缩放图片大小" & vbCrLf
    strPhotoJs = strPhotoJs & "function bbimg(o)" & vbCrLf
    strPhotoJs = strPhotoJs & "{" & vbCrLf
    'strPhotoJs = strPhotoJs & "  var zoom=parseInt(o.style.zoom, 10)||100;" & vbCrLf
    'strPhotoJs = strPhotoJs & "  zoom+=event.wheelDelta/12;" & vbCrLf
    'strPhotoJs = strPhotoJs & "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    strPhotoJs = strPhotoJs & "  return true;" & vbCrLf
    strPhotoJs = strPhotoJs & "}" & vbCrLf
    strPhotoJs = strPhotoJs & "-->" & vbCrLf
    strPhotoJs = strPhotoJs & "</script>" & vbCrLf
    strPhotoJs = strPhotoJs & "</head>" & vbCrLf

    If TemplateType = 3 Then
        If InStr(Content, "resizepic(thispic)") <= 0 Or InStr(Content, "bbimg(o)") <= 0 Then
            Content = Replace(Content, "</head>", strPhotoJs)
        End If
    End If
    
    ShiftCharacterSave = Content
End Function


'**************************************************
'函数名：BatchReplaceString
'作  用：批量替换处理函数
'参  数：TemplateContent ----模板数据
'参  数：TemplateReplaceStart ----获得要替换的开头代码
'参  数：TemplateReplaceEnd ----获得要替换的结束代码
'参  数：TemplateReplaceResult ----要替换的代码
'参  数：TemplateName ----模板名称
'返回值：True  ----已创建
'**************************************************
Function BatchReplaceString(TemplateContent, _
                                    TemplateReplaceStart, _
                                    TemplateReplaceEnd, _
                                    TemplateReplaceResult, _
                                    TemplateName)

    If InStr(TemplateContent, TemplateReplaceStart) = 0 Or InStr(TemplateContent, TemplateReplaceEnd) = 0 Then
        BatchReplaceString = TemplateContent
        Response.Write "<br>&nbsp;&nbsp;" & TemplateName & "..<font color='#FF0000'>模板替换开始代码 或 结束不存在,不用替换！</font>"
        Exit Function
    End If

    If GetBody(TemplateContent, TemplateReplaceStart, TemplateReplaceEnd, True, True) = "" Then
        BatchReplaceString = TemplateContent
        Response.Write "<br>&nbsp;&nbsp;" & TemplateName & "..<font color='#FF0000'>模板替换开始代码 或 结束找寻位置不对,不用替换！</font>"
        Exit Function
    End If

    BatchReplaceString = Replace(TemplateContent, GetBody(TemplateContent, TemplateReplaceStart, TemplateReplaceEnd, True, True), TemplateReplaceResult)
    Response.Write "<br>&nbsp;&nbsp;" & TemplateName & "..<font color='#009900'>模板替换成功！</font>"
End Function

'=================================================
'过程名：Resumeblank
'作  用：排序客户端 html
'=================================================
Sub Resumeblank()

    Response.Write "Function  Resumeblank(byval Content)" & vbCrLf
    Response.Write " Dim strHtml,strHtml2,Num,Numtemp,Strtemp" & vbCrLf
    Response.Write "   strHtml=Replace(Content, ""<DIV"", ""<div"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DIV>"", vbCrLf & ""</div>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<DD>"", vbCrLf & ""<dd>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<DT>"", vbCrLf & ""<dt>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<DL>"", vbCrLf & ""<dl>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DD>"", vbCrLf & ""</dd>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DT>"", vbCrLf & ""</dt>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DL>"", vbCrLf & ""</dl>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TABLE"", ""<table"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TABLE>"", vbCrLf & ""</table>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TBODY>"", """")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TBODY>"","""" & vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TR"", ""<tr"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TR>"", vbCrLf & ""</tr>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TD"", ""<td"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TD>"", ""</td>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<!--"", vbCrLf & ""<!--"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<SELECT"",vbCrLf & ""<Select"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</SELECT>"",vbCrLf & ""</Select>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<OPTION"",vbCrLf & ""  <Option"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</OPTION>"",""</Option>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<INPUT"",vbCrLf & ""  <Input"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<script"",vbCrLf & ""<script"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""&amp;"",""&"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""{$--"",vbCrLf & ""<!--$"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""--}"",""$-->"")" & vbCrLf
    Response.Write "   arrContent = Split(strHtml,vbCrLf)" & vbCrLf
    Response.Write "    For i = 0 To UBound(arrContent)" & vbCrLf
    Response.Write "        Numtemp=false" & vbCrLf
    Response.Write "        if Instr(arrContent(i),""<table"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<table"" and Strtemp <>""</table>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<table""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<tr"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<tr"" and Strtemp<>""</tr>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<tr""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<td"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<td"" and Strtemp<>""</td>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<td""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</table>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</table>"" and Strtemp<>""<table"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</table>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</tr>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</tr>"" and Strtemp<>""<tr"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</tr>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</td>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</td>"" and Strtemp<>""<td"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</td>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<!--"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Num< 0 then Num = 0" & vbCrLf
    Response.Write "        if trim(arrContent(i))<>"""" then" & vbCrLf
    Response.Write "            if i=0 then" & vbCrLf
    Response.Write "                strHtml2= string(Num,"" "") & arrContent(i) " & vbCrLf
    Response.Write "            elseif Numtemp=True then" & vbCrLf
    Response.Write "                strHtml2= strHtml2 & vbCrLf & string(Num,"" "") & arrContent(i) " & vbCrLf
    Response.Write "            else" & vbCrLf
    Response.Write "                strHtml2= strHtml2 & vbCrLf & arrContent(i) " & vbCrLf
    Response.Write "            end if" & vbCrLf
    Response.Write "        end if" & vbCrLf
    Response.Write "    Next" & vbCrLf
    Response.Write "    Resumeblank=strHtml2" & vbCrLf
    Response.Write "End function" & vbCrLf
End Sub

'**************************************************
'函数名：IsFontChecked
'作  用：是否是默认,默认显示红色
'参  数：Compare1-----比较值1
'参  数：Compare2-----比较值2
'返回值：替换后字符串
'**************************************************
Function IsFontChecked(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsFontChecked = " color='red'"
    Else
        IsFontChecked = ""
    End If
End Function

'**************************************************
'函数名：IsFontChecked2
'作  用：是否是默认,默认显示红色
'参  数：Compare1-----比较值1
'参  数：Compare2-----比较值2
'返回值：替换后字符串
'**************************************************
Function IsFontChecked2(ByVal Compare1, ByVal Compare2, ByVal IsOnlinePayment1, ByVal IsOnlinePayment2)
    If Compare1 = Compare2 Then
        If IsOnlinePayment1 = IsOnlinePayment2 Then
            IsFontChecked2 = " color='red'"
        End If
    Else
        IsFontChecked2 = ""
    End If
End Function


'**************************************************
'函数名：IsRadioChecked
'作  用：单选,多选默认
'参  数：Compare1-----比较值1
'参  数：Compare2-----比较值2
'返回值：替换后字符串
'**************************************************
Function IsRadioChecked(ByVal Compare1, _
                                ByVal Compare2)

    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If

End Function

'=================================================
'过程名：GetProject_Option
'作  用：调用所属方案
'参  数：iProjectName  ----方案名称
'=================================================
Function GetProject_Option(iProjectName)
    Dim sqlProject, rsProject, strProject

    sqlProject = "select * from PE_TemplateProject"
    Set rsProject = Conn.Execute(sqlProject)

    If rsProject.BOF And rsProject.EOF Then
    Else

        Do While Not rsProject.EOF
            strProject = strProject & "<option value='" & rsProject("TemplateProjectName") & "'"

            If rsProject("TemplateProjectName") = iProjectName Then
                strProject = strProject & " selected"
            End If

            strProject = strProject & ">" & rsProject("TemplateProjectName")

            If rsProject("IsDefault") = True Then
                strProject = strProject & "（默认）"
            End If

            strProject = strProject & "</option>"
            rsProject.MoveNext
        Loop

    End If

    rsProject.Close
    Set rsProject = Nothing
    GetProject_Option = strProject
End Function

'=================================================
'过程名：GetAllModule
'作  用：显示下拉菜单的导入模块
'=================================================
Sub GetAllModule(SystemType, ModuleType)
    Response.Write "<option " & OptionValue(CLng(ModuleType), 0) & ">通用模块</option>"
    Response.Write "<option " & OptionValue(CLng(ModuleType), 1) & ">文章模块</option>" & vbCrLf
    Response.Write "<option " & OptionValue(CLng(ModuleType), 2) & ">下载模块</option>" & vbCrLf
    Response.Write "<option " & OptionValue(CLng(ModuleType), 3) & ">图片模块</option>" & vbCrLf
    
    If SystemType = "4.03" Then
        Response.Write "<option " & OptionValue(CLng(ModuleType), -2) & ">用户自定义模块</option>" & vbCrLf
    Else
        Response.Write "<option " & OptionValue(CLng(ModuleType), 4) & ">留言模块</option>" & vbCrLf
        Response.Write "<option " & OptionValue(CLng(ModuleType), 5) & ">商城模块</option>" & vbCrLf
        If FoundInArr(AllModules, "Supply", ",") Then
            Response.Write "<option " & OptionValue(CLng(ModuleType), 6) & ">供求模块</option>" & vbCrLf
        End If
        If FoundInArr(AllModules, "House", ",") Then
            Response.Write "<option " & OptionValue(CLng(ModuleType), 7) & ">房产模块</option>" & vbCrLf
        End If
        If FoundInArr(AllModules, "Job", ",") Then
            Response.Write "<option " & OptionValue(CLng(ModuleType), 8) & ">招聘模块</option>" & vbCrLf
        End If
    End If

    Response.Write "<option " & OptionValue(CLng(ModuleType), 999999) & ">所有模块</option>" & vbCrLf
End Sub

'=================================================
'函数名：GetTemplate_Option
'作  用：频道模板类型下拉选择
'参  数：CurrentTemplateType --- 代入的模板值
'=================================================
Function GetTemplate_Option(CurrentTemplateType)
    Dim strTemp
    If ChannelID > 0 Then
        If ModuleType = 4 Then
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">留言首页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">留言发表模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 4) & ">留言回复模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">留言搜索页模板</option>" & vbCrLf
        Else
            Select Case ModuleType
            Case 7 '房产模板
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">频道首页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 2) & ">栏目模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 7) & ">推荐页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 8) & ">热门页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 30) & ">出售内容页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 31) & ">出租内容页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 32) & ">求购内容页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 33) & ">求租内容页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 34) & ">合租内容页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">搜索页模板</option>" & vbCrLf
            Case 8 '招聘模板
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">频道首页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">职位搜索页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">职位申请页模板</option>" & vbCrLf
            Case Else
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">频道首页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 2) & ">栏目页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">内容页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 4) & ">专题页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 22) & ">专题列表页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">搜索页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 6) & ">最新" & ChannelShortName & "页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 7) & ">推荐" & ChannelShortName & "页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 8) & ">热门" & ChannelShortName & "页模板</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 16) & ">评论" & ChannelShortName & "页模板</option>" & vbCrLf

                If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Then
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 23) & ">更多相关" & ChannelShortName & "页模板</option>" & vbCrLf
                End If

                If ModuleType = 1 Then
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 17) & ">打印页模板</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 20) & ">告诉好友页模板</option>" & vbCrLf
                ElseIf ModuleType = 5 Then
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 9) & ">购物车模板</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 10) & ">收银台模板</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 11) & ">订单预览页模板</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 12) & ">订购成功页模板</option>" & vbCrLf
                    'strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 13) & ">在线支付第一步模板</option>" & vbCrLf
                    'strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 14) & ">在线支付第二步模板</option>" & vbCrLf
                    'strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 15) & ">在线支付第三步模板</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 19) & ">特价商品页模板</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 21) & ">商城帮助页模板</option>" & vbCrLf
                End If
            End Select
        End If

    Else

        If TempType = 1 Then
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 102) & ">会员中心通用模板</option>" & vbCrLf		
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 8) & ">会员信息页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 9) & ">会员列表页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 18) & ">会员注册页模板（许可协议）</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 19) & ">会员注册页模板（必填项目）</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 20) & ">会员注册页模板（选填项目）</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 21) & ">会员注册页模板（注册结果）</option>" & vbCrLf
        Else
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">网站首页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">网站搜索页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 4) & ">网站公告模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 22) & ">网站公告列表页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">网站友情页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 6) & ">网站调查页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 7) & ">版权声明页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 29) & ">全站专题列表页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 30) & ">全站专题页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 10) & ">作者显示页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 11) & ">作者列表页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 12) & ">来源显示页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 13) & ">来源列表页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 103) & ">匿名投稿模板</option>" & vbCrLf
			
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 14) & ">厂商显示页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 15) & ">厂商列表页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 16) & ">品牌显示页模板</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 17) & ">品牌列表页模板</option>" & vbCrLf

            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 101) & ">自定义列表模板</option>" & vbCrLf
        End If
    End If

    GetTemplate_Option = strTemp
End Function

'=================================================
'函数名：GetTemplateTypeName
'作  用：显示当前频道的模板类型
'参  数：iTemplateType --- 代入的模板值
'=================================================
Function GetTemplateTypeName(iTemplateType, _
                                     ChannelID)

    If ChannelID > 0 Then
        If ModuleType = 4 Then

            Select Case iTemplateType

                Case 1
                    GetTemplateTypeName = "留言首页模板"

                Case 3
                    GetTemplateTypeName = "留言发表模板"

                Case 4
                    GetTemplateTypeName = "留言回复模板"

                Case 5
                    GetTemplateTypeName = "留言搜索页模板"
            End Select

        Else

            Select Case iTemplateType
            Case 1
                GetTemplateTypeName = "频道首页模板"
            Case 2
                GetTemplateTypeName = "频道栏目模板"
            Case 3
                GetTemplateTypeName = "频道内容页模板"
            Case 4
                GetTemplateTypeName = "频道专题页模板"
            Case 5
                GetTemplateTypeName = "频道搜索页模板"
            Case 6
                GetTemplateTypeName = "最新" & ChannelShortName & "页模板"
            Case 7
                GetTemplateTypeName = "推荐" & ChannelShortName & "页模板"
            Case 8
                GetTemplateTypeName = "热点" & ChannelShortName & "页模板"
            Case 16
                GetTemplateTypeName = "评论" & ChannelShortName & "页模板"
            Case 9
                GetTemplateTypeName = "购物车模板"
            Case 10
                GetTemplateTypeName = "收银台模板"
            Case 11
                GetTemplateTypeName = "预览订单模板"
            Case 12
                GetTemplateTypeName = "订购成功页模板"
            'Case 13
            '   GetTemplateTypeName = "在线支付第一步模板"
            'Case 14
            '   GetTemplateTypeName = "在线支付第二步模板"
            'Case 15
            '   GetTemplateTypeName = "在线支付第三步模板"
            Case 17
                GetTemplateTypeName = "打印模板"
            Case 101
                GetTemplateTypeName = "自定义列表模板"
            Case 19
                GetTemplateTypeName = "特价商品页模板"
            Case 20
                GetTemplateTypeName = "告诉好友页模板"
            Case 21
                GetTemplateTypeName = "商城帮助页模板"
            Case 22
                GetTemplateTypeName = "频道专题列表页模板"
            Case 23
                GetTemplateTypeName = "更多相关" & ChannelShortName & "页模板"
            '增加房产模块的相关模板类型
            Case 30
                GetTemplateTypeName = "出售内容页模板"
            Case 31
                GetTemplateTypeName = "出租内容页模板"
            Case 32
                GetTemplateTypeName = "求购内容页模板"
            Case 33
                GetTemplateTypeName = "求租内容页模板"
            Case 34
                GetTemplateTypeName = "合租内容页模板"
            '***************End********************
            End Select
        End If
    Else
        Select Case iTemplateType
        Case 1
            GetTemplateTypeName = "网站首页模板"
        Case 3
            GetTemplateTypeName = "网站搜索页模板"
        Case 4
            GetTemplateTypeName = "网站公告页模板"
        Case 5
            GetTemplateTypeName = "友情链接页模板"
        Case 6
            GetTemplateTypeName = "网站调查页模板"
        Case 7
            GetTemplateTypeName = "版权声明页模板"
        Case 8
            GetTemplateTypeName = "会员信息页模板"
        Case 102	
            GetTemplateTypeName = "会员中心通用模板"			
        Case 9
            GetTemplateTypeName = "会员列表页模板"
        Case 10
            GetTemplateTypeName = "作者显示页模板"
        Case 11
            GetTemplateTypeName = "作者列表页模板"
        Case 12
            GetTemplateTypeName = "来源显示页模板"
        Case 13
            If IsOnlinePayment = 0 Then
                GetTemplateTypeName = "来源列表页模板"
            'Else
            '   GetTemplateTypeName = "在线支付第一步模板"
            End If
        Case 103
            GetTemplateTypeName = "匿名投稿模板"			
        Case 14
            If IsOnlinePayment = 0 Then
                GetTemplateTypeName = "厂商显示页模板"
            'Else
            '   GetTemplateTypeName = "在线支付第二步模板"
            End If
        Case 15
            If IsOnlinePayment = 0 Then
                GetTemplateTypeName = "厂商列表页模板"
            'Else
            '   GetTemplateTypeName = "在线支付第三步模板"
            End If
        Case 16
            GetTemplateTypeName = "品牌显示页模板"
        Case 17
            GetTemplateTypeName = "品牌列表页模板"
        Case 101
            GetTemplateTypeName = "自定义列表模板"
        Case 18
            GetTemplateTypeName = "会员注册页模板（许可协议）"
        Case 19
            GetTemplateTypeName = "会员注册页模板（必填项目）"
        Case 20
            GetTemplateTypeName = "会员注册页模板（选填项目）"
        Case 21
            GetTemplateTypeName = "会员注册页模板（注册结果）"
        Case 22
            GetTemplateTypeName = "公告列表页模板"
        Case 29
            GetTemplateTypeName = "全站专题列表页模板"
        Case 30
            GetTemplateTypeName = "全站专题页模板"
        End Select

    End If

    If iTemplateType = 0 Then
        GetTemplateTypeName = "当前类型所有模板"
    End If

End Function
%>
