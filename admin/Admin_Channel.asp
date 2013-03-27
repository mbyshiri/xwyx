<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Channel"   '其他权限

rsGetAdmin.Close
Set rsGetAdmin = Nothing

Response.Write "<html><head><title>频道管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("频 道 管 理", 10002)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Channel.asp'>频道管理首页</a>&nbsp;|&nbsp;<a href='Admin_Channel.asp?Action=Add'>添加新频道</a>&nbsp;|&nbsp;<a href='Admin_Channel.asp?Action=Order'>频道排序</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Action = Trim(Request("Action"))
Select Case Action
Case "Add"
    Call AddChannel
Case "SaveAdd"
    Call SaveAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Disabled"
    Call DisabledChannel(0)
Case "UnDisabled"
    Call DisabledChannel(1)
Case "Del"
    Call DelChannel
Case "Order"
    Call order
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "UpdateData"
    Call UpdateData
Case "UpdateChannelFiles"
    Call UpdateChannelFiles
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "频道管理操作失败，失败原因：" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsChannelList, sqlChannelList
    sqlChannelList = "select * from PE_Channel Where 1=1"
    If Not (FoundInArr(AllModules, "Supply", ",")) Then
        sqlChannelList = sqlChannelList & " And ModuleType<>6"
    End If
    If Not (FoundInArr(AllModules, "Job", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>8"
    End If
    If Not (FoundInArr(AllModules, "House", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>7"
    End If
    sqlChannelList = sqlChannelList & " order by OrderID "
    Set rsChannelList = Conn.Execute(sqlChannelList)
    
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "    <td align='center'><strong>频道名称</strong></td>"
    Response.Write "    <td width='54' align='center'><strong>打开方式</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>频道类型</strong></td>"
    Response.Write "    <td width='120' align='center'><strong>频道目录/链接地址</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>项目名称</strong></td>"
    Response.Write "    <td width='54' align='center'><strong>功能模块</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>生成HTML方式</strong></td>"
    Response.Write "    <td width='54' align='center'><strong>频道状态</strong></td>"
    Response.Write "    <td width='110' align='center'><strong>操作</strong></td>"
    Response.Write "    <td width='65' align='center'><strong>频道更新</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    Do While Not rsChannelList.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td align='center'>" & rsChannelList("ChannelID") & "</td>"
        Response.Write "    <td align='center'><a href='Admin_Channel.asp?Action=Modify&iChannelID=" & rsChannelList("ChannelID") & "' title='" & rsChannelList("ReadMe") & "'>" & rsChannelList("ChannelName") & "</a></td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("OpenType") = 0 Then
            Response.Write "<font color=green>原窗口</font>"
        Else
            Response.Write "新窗口"
        End If
        Response.Write "</td>"
        Response.Write "<td width='60' align='center'>"
        Select Case rsChannelList("ChannelType")
        Case 0
            Response.Write "<font color=blue>系统频道</font>"
        Case 1
            Response.Write "<font color=green>内部频道</font>"
        Case 2
            Response.Write "<font color=red>外部频道</font>"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='120' style='word-wrap:break-word'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write "目录：" & rsChannelList("ChannelDir")
        Else
            Response.Write "<font color=red>链接：" & rsChannelList("LinkUrl") & "</font>"
        End If
        Response.Write "</td>"
        Response.Write "    <td width='60' align='center'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write rsChannelList("ChannelShortName")
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write GetModuleTypeName(rsChannelList("ModuleType"))
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='60' align='center'>"
        Select Case rsChannelList("UseCreateHTML")
        Case 0
            Response.Write "不生成"
        Case 1
            Response.Write "全部生成"
        Case 2
            Response.Write "部分生成1"
        Case 3
            Response.Write "部分生成2"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("Disabled") = True Then
            Response.Write "<font color=red>已禁用</font>"
        Else
            Response.Write "正常"
        End If
        Response.Write "</td>"
        Response.Write "<td width='110' align='center'>"
        Response.Write "<a href='Admin_Channel.asp?Action=Modify&iChannelID=" & rsChannelList("ChannelID") & "'>修改</a>&nbsp;&nbsp;"
        If rsChannelList("Disabled") = True Then
            Response.Write "<a href='Admin_Channel.asp?Action=UnDisabled&iChannelID=" & rsChannelList("ChannelID") & "'>启用</a>&nbsp;&nbsp;"
        Else
            Response.Write "<a href='Admin_Channel.asp?Action=Disabled&iChannelID=" & rsChannelList("ChannelID") & "'>禁用</a>&nbsp;&nbsp;"
        End If
        If rsChannelList("ChannelType") > 0 Then
            Response.Write "<a href='Admin_Channel.asp?Action=Del&iChannelID=" & rsChannelList("ChannelID") & "' onClick=""return confirm('确定要删除此频道吗？');"">删除</a>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<td width='65' align='center'>"
        If rsChannelList("ChannelType") < 2 And rsChannelList("ModuleType") <> 4 And rsChannelList("ModuleType") <> 8 Then
            Response.Write "<a href='Admin_Channel.asp?Action=UpdateData&iChannelID=" & rsChannelList("ChannelID") & "'>数据</a>&nbsp;"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        If rsChannelList("ChannelType") = 1 And rsChannelList("ModuleType") <> 4 Then
            Response.Write "<a href='Admin_Channel.asp?Action=UpdateChannelFiles&iChannelID=" & rsChannelList("ChannelID") & "'>文件</a>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "</tr>"
        rsChannelList.MoveNext
    Loop
    Response.Write "</table>"
    rsChannelList.Close
    Set rsChannelList = Nothing
    Response.Write "<form name='form1' method='post' action='Admin_Channel.asp'><div align='center'>"
    Response.Write "<input type='hidden' name='Action' value='UpdateData'>"
    Response.Write "<input type='submit' name='submit' value='更新所有频道的数据' onclick=""document.form1.Action.value='UpdateData'""> "
    Response.Write "<input type='submit' name='submit' value='更新所有频道的文件' onclick=""document.form1.Action.value='UpdateChannelFiles'"">"
    Response.Write "</div></form>"
End Sub

Sub order()
    Dim rsChannelList, sqlChannelList, iCount, i, j
    'sqlChannelList = "select * from PE_Channel order by OrderID"
    sqlChannelList = "select * from PE_Channel Where 1=1"
    If Not (FoundInArr(AllModules, "Supply", ",")) Then
        sqlChannelList = sqlChannelList & " And ModuleType<>6"
    End If
    If Not (FoundInArr(AllModules, "Job", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>8"
    End If
    If Not (FoundInArr(AllModules, "House", ",")) Then
        sqlChannelList = sqlChannelList & " And  ModuleType<>7"
    End If
    sqlChannelList = sqlChannelList & " order by OrderID "
    Set rsChannelList = Server.CreateObject("Adodb.RecordSet")
    rsChannelList.Open sqlChannelList, Conn, 1, 1
    iCount = rsChannelList.RecordCount
    j = 1
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write " <td width='32' align='center'><strong>序号</strong></td>"
    Response.Write " <td height='22' align='center'><strong> 频道名称</strong></td>"
    Response.Write " <td width='54' align='center'><strong>打开方式</strong></td>"
    Response.Write " <td width='80' align='center'><strong>频道类型</strong></td>"
    Response.Write " <td width='120' align='center'><strong>频道目录/</strong><strong>链接地址</strong></td>"
    Response.Write " <td width='80' align='center'><strong>功能模块</strong></td>"
    Response.Write " <td width='240' colspan='2' align='center'><strong>操作</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    Do While Not rsChannelList.EOF
        Response.Write "<tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "<td width='32' align='center'>" & rsChannelList("OrderID") & "</td>"
        Response.Write "    <td align='center'><a href='Admin_Channel.asp?Action=Modify&iChannelID=" & rsChannelList("ChannelID") & "' title='" & nohtml(rsChannelList("ReadMe")) & "'>" & rsChannelList("ChannelName") & "</a></td>"
        Response.Write "<td width='54' align='center'>"
        If rsChannelList("OpenType") = 0 Then
            Response.Write "<font color=green>原窗口</font>"
        Else
            Response.Write "新窗口"
        End If
        Response.Write "</td>"
        Response.Write "<td width='80' align='center'>"
        Select Case rsChannelList("ChannelType")
        Case 0
            Response.Write "<font color=blue>系统频道</font>"
        Case 1
            Response.Write "<font color=green>内部频道</font>"
        Case 2
            Response.Write "<font color=red>外部频道</font>"
        End Select
        Response.Write "</td>"
        Response.Write "<td width='120'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write "目录：" & rsChannelList("ChannelDir")
        Else
            Response.Write "<font color=red>链接：" & rsChannelList("LinkUrl") & "</font>"
        End If
        Response.Write "</td>"
        Response.Write "<td width='80' align='center'>"
        If rsChannelList("ChannelType") <= 1 Then
            Response.Write GetModuleTypeName(rsChannelList("ModuleType"))
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "<form action='Admin_Channel.asp?Action=UpOrder' method='post'>"
        Response.Write "  <td width='120' align='center'>"
        If j > 1 Then
            Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select>"
            Response.Write "<input type=hidden name=iChannelID value=" & rsChannelList("ChannelID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsChannelList("OrderID") & ">&nbsp;<input type=submit name=Submit value=修改>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form>"
        Response.Write "<form action='Admin_Channel.asp?Action=DownOrder' method='post'>"
        Response.Write "  <td width='120' align='center'>"
        If iCount > j Then
            Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select>"
            Response.Write "<input type=hidden name=iChannelID value=" & rsChannelList("ChannelID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsChannelList("OrderID") & ">&nbsp;<input type=submit name=Submit value=修改>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form></tr>"
        j = j + 1
        rsChannelList.MoveNext
    Loop
    Response.Write "</table>"
    rsChannelList.Close
    Set rsChannelList = Nothing
End Sub

Sub AddChannel()
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Channel.asp'>频道管理</a>&nbsp;&gt;&gt;&nbsp;添加新频道</td></tr></table>"
    Response.Write "<form method='post' action='Admin_Channel.asp' name='myform' onSubmit='return CheckForm();'>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>频道设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>前台样式</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>上传选项</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>生成选项</td>" & vbCrLf
    If IsCustom_Content = True Then
        Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>自设内容</td>" & vbCrLf
    End If
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong> 频道名称：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelName' type='text' id='ChannelName' size='49' maxlength='30'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>频道图片：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelPicUrl' type='text' id='ChannelPicUrl' size='49' maxlength='200'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>频道说明：</strong><br>鼠标移至频道名称上时将显示设定的说明文字（不支持HTML）</td>" & vbCrLf
    Response.Write "      <td valign='middle'><textarea name='ReadMe' cols='40' rows='3' id='ReadMe'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道类型：</strong><br><font color=red>请慎重选择，频道一旦添加后就不能再更改频道类型。</font></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function HideTabTitle(displayValue,tempType){" & vbCrLf
    Response.Write "  for (var i = 1; i < TabTitle.length; i++) {" & vbCrLf
    Response.Write "    if(tempType==0&&i==2) {" & vbCrLf
    Response.Write "        TabTitle[i].style.display='none';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    else{" & vbCrLf
    Response.Write "        TabTitle[i].style.display=displayValue;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<input name='ChannelType' type='radio' value='2'  onclick=""HideTabTitle('none');ChannelSetting.style.display='none'"" ><font color=blue><b>外部频道</b></font>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;外部频道指链接到本系统以外的地址中。当此频道准备链接到网站中的其他系统时，请使用这种方式。<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;外部频道的链接地址：<input name='LinkUrl' type='text' id='LinkUrl' value='' size='40' maxlength='200'>"
    Response.Write "   <br><br>" & vbCrLf
    Response.Write "   <input name='ChannelType' type='radio' value='1' checked"
    If ObjInstalled_FSO = False Then Response.Write " disabled "
    Response.Write " onclick=""HideTabTitle('',1);ChannelSetting.style.display=''"">"
    Response.Write "<font color=blue><b>系统内部频道</b></font>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;系统内部频道指的是在本系统现有功能模块（新闻、文章、图片等）基础上添加新的频道，新频道具备和所使用功能模块完全相同的功能。例如，添加一个名为“网络学院”的新频道，新频道使用“文章”模块的功能，则新添加的“网络学院”频道具有原文章频道的所有功能。<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;此功能需要服务器支持FSO才可用。<br>" & vbCrLf
    Response.Write "      <table id='ChannelSetting' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF' style='display:'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>频道使用的功能模块：</strong></td>"
    Response.Write "          <td><select name='ModuleType' id='ModuleType'>"
    Response.Write "          <option value='1' selected>文章</option>"
    Response.Write "          <option value='2'>下载</option>"
    Response.Write "          <option value='3'>图片</option>"
    Response.Write "          </select>&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>请慎重选择，频道一旦添加后就不能修改此项。</font></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>频道目录：</strong>（频道英文名）<br>"
    Response.Write "          <font color='#FF0000'>只能是英文，不能带空格或“\”、“/”等符号。</font><br>"
    Response.Write "          <font color='#0000FF'>样例：</font>News、Article、Soft</td>"
    Response.Write "          <td><input name='ChannelDir' type='text' id='ChannelDir' size='20' maxlength='50'>  <font color='#FF0000'>*&nbsp;&nbsp;&nbsp;&nbsp;请慎重录入，频道一旦添加后就不能修改此项。</font></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong><font color=red>频道链接地址（子域名）：</font></strong><br>如果您要在前台将此频道做为主站的一个<font color='red'>独立子站点</font>来访问，请输入完整的网址（如：http://news.powereasy.net）；否则，请保持为空。</td>"
    Response.Write "          <td><input name='ChannelUrl' type='input' value='' size='30' maxlength='100'"
    If SiteUrlType = 0 Then Response.Write " disabled"
    Response.Write "> <font color='red'>* 不能带目录</font><br>如果要启用此功能，必须在“网站选项”中将“链接地址方式”改为“绝对路径”</td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>项目名称：</strong><br>例如：频道名称为“网络学院”，其项目名称为“文章”或“教程”</td>"
    Response.Write "          <td><input name='ChannelShortName' type='text' id='ChannelShortName' size='20' maxlength='30'> <font color='#FF0000'>*</font></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='200' class='tdbg5'><strong>项目单位：</strong><br>例如：“篇”、“条”、“个”</td>"
    Response.Write "          <td><input name='ChannelItemUnit' type='text' id='ChannelItemUnit' size='10' maxlength='30'></td>"
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>打开方式：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='OpenType' value='0'>在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='OpenType' type='radio' value='1' checked>在新窗口打开" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>禁用本频道：</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='Disabled' value='True'>是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='Disabled' type='radio' value='False' checked>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道权限：</strong><br><font color='red'>频道权限为继承关系，当频道设为“认证频道”时，其下的栏目设为“开放栏目”也无效。相反，如果频道设为“开放频道”，其下的栏目可以自由设置权限。</font></td>"
    Response.Write "      <td>"
    Response.Write "        <table>"
    Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' value='0' checked>开放频道</td><td>任何人（包括游客）可以浏览此频道下的信息。可以在栏目设置中再指定具体的栏目权限。</td></tr>"
    Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' value='1'>认证频道</td><td>游客不能浏览，并在下面指定允许浏览的会员组。如果频道设置为认证频道，则此频道的“生成HTML”选项只能设为“不生成HTML”。</td></tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>允许浏览此频道的会员组：</strong><br>如果频道权限设置为“认证频道”，请在此设置允许浏览此频道的会员组</td>"
    Response.Write "      <td>" & GetUserGroup("", "") & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>本频道的审核级别：</strong><br>设定为需要审核时，如果某组会员有“发表信息不需审核”的特权，则此会员组不受此限。</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='CheckLevel' type='radio' value='0'>不需审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>在本频道发表信息不需要管理员审核</font><br>"
    Response.Write "        <input name='CheckLevel' type='radio' value='1' checked>一级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>需要栏目审核员进行审核（注：此级别为最小级别，下同）</font><br>"
    Response.Write "        <input name='CheckLevel' type='radio' value='2'>二级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>初审需要栏目审核员审核，终审需要栏目总编进行审核</font><br>"
    Response.Write "        <input name='CheckLevel' type='radio' value='3'>三级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>初审需要栏目审核员审核，二审需要栏目总编审核，终审需要频道管理员审核</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf	
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>不属于任何栏目的评论权限：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input name='EnableComment' type='checkbox' value='True' checked>允许发表评论<br>"
    Response.Write "        <input name='CheckComment' type='checkbox' value='True' checked>评论需要审核"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>退稿时站内短信/Email通知内容：</strong><br>不支持HTML代码</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfReject' cols='60' rows='4'>非常抱歉的告诉您，您的{$ChannelShortName}《{$Title}》因为以下几个原因未被录用：" & vbCrLf & "1、" & vbCrLf & "2、" & vbCrLf & "3、" & vbCrLf & vbCrLf & "期待着您的再次投稿！</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>稿件被采用时站内短信/Email通知内容：</strong><br>不支持HTML代码</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfPassed' cols='60' rows='4'>恭喜您！您的{$ChannelShortName}《{$Title}》已经被录用！" & vbCrLf & "非常感谢您的投稿！</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>频道META关键词：</strong><br>针对搜索引擎设置的关键词<br>多个关键词请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>频道META网页描述：</strong><br>针对搜索引擎设置的网页描述<br>多个描述请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>添加/修改信息时的界面设置：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input name='arrEnabledTabs' type='checkbox' value='Charge' checked>显示“收费选项”书签<br>"
    Response.Write "        <input name='arrEnabledTabs' type='checkbox' value='Vote' checked>显示“调查设置”书签<br>"
    Response.Write "        <input name='arrEnabledTabs' type='checkbox' value='SoftParameter' checked>显示“软件参数”书签（仅对下载模块有效）<br>"
    Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Recieve' checked>显示“签收设置”书签（仅对文章模块有效）<br>"
    Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Copyfee' checked>显示“稿费设置”书签（仅对文章模块有效）<br>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>默认稿费标准：</strong>(单位：元/千字)</td>"
    Response.Write "      <td><input name='MoneyPerKw' type='text'id='MoneyPerKw' size='10' maxlength='10'> <font color=red>元/千字</font></Td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>是否在频道栏显示频道名称：</strong></td>"
    Response.Write "      <td><input name='ShowName' type='radio' value='True' checked>是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>是否在导航栏显示频道名称：</strong></td>"
    Response.Write "      <td><input name='ShowNameOnPath' type='radio' value='True' checked>是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>是否在本频道显示树状导航菜单：</strong></td>"
    Response.Write "      <td><input name='ShowClassTreeGuide' type='radio' value='True' checked>是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowClassTreeGuide' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>前台是否显示标题后的省略号：</strong><br>当模板中指定标题长度小于标题实际长度时，可以决定是否在标题后面显示省略号</td>"
    Response.Write "      <td><input name='ShowSuspensionPoints' type='radio' value='True' checked>是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowSuspensionPoints' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>本频道热点的点击数最小值：</strong></td>"
    Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='500' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>多少天内更新的信息为新信息：</strong></td>"
    Response.Write "      <td><input name='DaysOfNew' type='text' id='DaysOfNew' value='7' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>顶部导航栏每行显示的栏目数：</strong></td>"
    Response.Write "      <td><input name='MaxPerLine' type='text' id='MaxPerLine' value='10' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>信息列表中作者名称的长度：</strong></td>"
    Response.Write "      <td><input name='AuthorInfoLen' type='text' id='AuthorInfoLen' value='8' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道首页专题列表的数量：</strong></td>"
    Response.Write "      <td><input name='JS_SpecialNum' type='text' id='JS_SpecialNum' value='10' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道首页的每页信息数：</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_Index' type='text' id='MaxPerPage_Index' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>搜索结果页的每页信息数：</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>最新信息页的每页信息数：</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_New' type='text' id='MaxPerPage_New' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>热门信息页的每页信息数：</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_Hot' type='text' id='MaxPerPage_Hot' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>推荐信息页的每页信息数：</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_Elite' type='text' id='MaxPerPage_Elite' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>专题列表页的每页信息数：</strong></td>"
    Response.Write "      <td><input name='MaxPerPage_SpecialList' type='text' id='MaxPerPage_SpecialList' value='20' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>本频道的默认风格：</strong></td>"
    Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(0) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>顶部栏目菜单的显示方式：</strong><br>更改此参数后需要刷新栏目JS方生效。</td>"
    Response.Write "      <td><select name='TopMenuType' id='TopMenuType'>" & GetMenuType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>底部栏目导航的显示方式：</strong><br>更改此参数后需要刷新栏目JS方生效。</td>"
    Response.Write "      <td><select name='ClassGuideType' id='ClassGuideType'>" & GetGuideType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf



    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>是否允许在本频道上传文件：</strong></td>"
    Response.Write "      <td><input name='EnableUploadFile' type='radio' value='True' checked>是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='EnableUploadFile' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Dim UploadDir
    Randomize
    UploadDir = "UploadFiles_" & CInt(Rnd * 8999 + 1000)
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>上传文件的保存目录：</strong><br><font color='red'>你可以定期或不定期的更改上传目录，以防其他网站盗链</font></td>"
    Response.Write "      <td><input name='UploadDir' type='text' id='UploadDir' value='" & UploadDir & "' size='20' maxlength='20'>&nbsp;&nbsp;<font color='red'>只能是英文和数字，不能带空格或“\”、“/”等符号。</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>允许上传的最大文件大小：</strong></td>"
    Response.Write "      <td><input name='MaxFileSize' type='text' id='MaxFileSize' value='1024' size='10' maxlength='10'> KB&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>允许上传的文件类型：</strong><br>多种文件类型之间以“|”分隔</td>"
    Response.Write "      <td><table>"
    Response.Write "          <tr><td>图片类型：</td><td><input name='UpFileType' type='text' id='UpFileType' value='gif|jpg|jpeg|jpe|bmp|png' size='50' maxlength='200'></td></tr>"
    Response.Write "          <tr><td>Flash文件：</td><td><input name='UpFileType' type='text' id='UpFileType' value='swf' size='50' maxlength='50'></td></tr>"
    Response.Write "          <tr><td>Windows媒体：</td><td><input name='UpFileType' type='text' id='UpFileType' value='mid|mp3|wmv|asf|avi|mpg' size='50' maxlength='200'></td></tr>"
    Response.Write "          <tr><td>Real媒体：</td><td><input name='UpFileType' type='text' id='UpFileType' value='ram|rm|ra' size='20' maxlength='200'></td></tr>"
    Response.Write "          <tr><td>其他文件：</td><td><input name='UpFileType' type='text' id='UpFileType' value='rar|exe|doc|zip' size='50' maxlength='200'></td></tr>"
    Response.Write "      </table></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf



    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>生成HTML方式：</font></strong><br>服务器支持FSO才能启用“生成HTML”功能<br>请谨慎选择！以后在每一次更改生成方式前，你最好先删除所有以前生成的文件，然后在保存频道参数后再重新生成所有文件。</td>"
    Response.Write "      <td>" & GetUseCreateHTML(0, 0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>栏目、专题列表更新页数：</font></strong><br>添加内容后自动更新的栏目及专题列表页数。</td>"
    Response.Write "      <td><input name='UpdatePages' type='text' id='UpdatePages' value='3' size='5' maxlength='5'> 页 <font color='#FF0000'>*</font>&nbsp;&nbsp;<font color='blue'>如：更新页数设为3，则每次自动更新前三页，第4页以后的分页为固定生成的页面，当新增内容数超过一页，则再生成一个固定页面，在总记录数不是每页记录数的整数倍时，交叉页（第3、4页）会有部分记录重复。</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr><td colspan='2'><font color='red'><b>以下参数仅当“生成HTML方式”设为后三者时才有效。<br>请谨慎选择！以后在每一次更改以下参数前，你最好先删除所有以前生成的文件，然后在保存参数设置后再重新生成所有文件。</b></font></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>自动生成HTML时的生成方式：</font></strong><br>添加/修改信息时，系统可以自动生成有关页面文件，请在这里选择自动生成时的方式。</td>"
    Response.Write "      <td>" & GetAutoCreateType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>栏目列表文件的存放位置：</font></strong></td>"
    Response.Write "      <td>" & GetListFileType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>目录结构方式：</font></strong></td>"
    Response.Write "      <td>" & GetStructureType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>内容页文件的命名方式：</font></strong></td>"
    Response.Write "      <td>" & GetFileNameType(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>频道首页的扩展名：</font></strong></td>"
    Response.Write "      <td>" & arrFileExt_Index(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>栏目页、专题页的扩展名：</font></strong></td>"
    Response.Write "      <td>" & arrFileExt_List(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>内容页的扩展名：</font></strong></td>"
    Response.Write "      <td>" & arrFileExt_Item(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    If IsCustom_Content = True Then
        Call EditCustom_Content("Add", "", "Channel")
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "      <input  type='submit' name='Submit' value=' 添 加 '> &nbsp; <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Channel.asp'"" style='cursor:hand;'></td>"
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "</form>"

    Call ShowChekcFormJS
End Sub

Sub ShowChekcFormJS()
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(document.myform.ChannelName.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('请输入频道名称！');" & vbCrLf
    Response.Write "    document.myform.ChannelName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.ChannelType[1].checked==true){" & vbCrLf
    Response.Write "    if(document.myform.ChannelDir.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('请输入频道目录！');" & vbCrLf
    Response.Write "      document.myform.ChannelDir.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(document.myform.ChannelShortName.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('请输入项目名称！');" & vbCrLf
    Response.Write "      document.myform.ChannelShortName.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    if(document.myform.LinkUrl.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('请输入频道的链接地址！');" & vbCrLf
    Response.Write "      document.myform.LinkUrl.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

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
End Sub

Sub Modify()
    Dim iChannelID, rsChannel
    iChannelID = Trim(Request("iChannelID"))
    If iChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改频道ID</li>"
        Exit Sub
    Else
        iChannelID = PE_CLng(iChannelID)
    End If
    Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & iChannelID)
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的频道！</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Sub
    End If
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
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

    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Channel.asp'>频道管理</a>&nbsp;&gt;&gt;&nbsp;修改频道设置：<font color='red'>" & rsChannel("ChannelName") & "</font></td></tr></table>"
    Response.Write "<form method='post' action='Admin_Channel.asp' name='myform' onSubmit='return CheckForm();'>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    If rsChannel("ChannelType") <> 2 Then
        If rsChannel("ModuleType") <> 4 Then
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>频道设置</td>" & vbCrLf
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>前台样式</td>" & vbCrLf
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>上传选项</td>" & vbCrLf			
        End If
        If rsChannel("ModuleType") = 4 Then		
             Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>上传选项</td>" & vbCrLf
        End If					
        '刘永涛，供求频道屏蔽
        If rsChannel("ModuleType") <> 6 And rsChannel("ModuleType") <> 4 And rsChannel("ModuleType") <> 7 And rsChannel("ModuleType") <> 8 Then
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>生成选项</td>" & vbCrLf
            If IsCustom_Content = True Then
                Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>自设内容</td>" & vbCrLf
            End If
        End If
        '2005-12-23
    End If
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong> 频道名称：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelName' type='text' id='ChannelName' size='49' maxlength='30' value='" & rsChannel("ChannelName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>频道图片：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelPicUrl' type='text' id='ChannelPicUrl' size='49' maxlength='200' value='" & rsChannel("ChannelPicUrl") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>频道说明：</strong><br>鼠标移至频道名称上时将显示设定的说明文字（不支持HTML）</td>" & vbCrLf
    Response.Write "      <td valign='middle'><textarea name='ReadMe' cols='40' rows='3' id='ReadMe'>" & rsChannel("ReadMe") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道类型：</strong><br><font color=red>请慎重选择，频道一旦添加后就不能再更改频道类型。</font></td>" & vbCrLf
    Response.Write "      <td><input name='ChannelType' type='radio' value='2'"
    If rsChannel("ChannelType") > 1 Then
        Response.Write " checked "
    Else
        Response.Write " disabled"
    End If
    Response.Write " onClick=""ChannelSetting.style.display='none'""><font color=blue><b>外部频道</b></font></legend>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;外部频道指链接到本系统以外的地址中。当此频道准备链接到网站中的其他系统时，请使用这种方式。<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;外部频道的链接地址："
    Response.Write "      <input name='LinkUrl' type='text' id='LinkUrl' value='" & rsChannel("LinkUrl") & "' size='40' maxlength='200'"
    If rsChannel("ChannelType") <= 1 Then
        Response.Write " disabled"
    End If
    Response.Write "><br><br>"
    Response.Write "<input name='ChannelType' type='radio' value='1'"
    If rsChannel("ChannelType") <= 1 Then
        Response.Write " checked"
    Else
        Response.Write " disabled"
    End If
    Response.Write "><font color=blue><b>系统内部频道</b></font></legend>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;系统内部频道指的是在本系统现有功能模块（新闻、文章、图片等）基础上添加新的频道，新频道具备和所使用功能模块完全相同的功能。例如，添加一个名为“网络学院”的新频道，新频道使用“文章”模块的功能，则新添加的“网络学院”频道具有原文章频道的所有功能。<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;此功能需要服务器支持FSO才可用。<br>"
    Response.Write "     <table id='ChannelSetting' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'"
    If rsChannel("ChannelType") > 1 Then Response.Write " style='display:none'"
    Response.Write ">"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td colspan='2'><strong>内部频道参数设置</strong></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道使用的功能模块：</strong></td>"
    Response.Write "      <td><select name='ModuleType' id='ModuleType' disabled>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 1) & ">文章</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 2) & ">下载</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 3) & ">图片</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 4) & ">留言板</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 5) & ">商城</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 6) & ">供求</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 7) & ">房产</option>"
    Response.Write "      <option " & OptionValue(rsChannel("ModuleType"), 8) & ">人才</option>"
    Response.Write "      </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>频道目录：</strong>（频道英文名）<br>"
    Response.Write "        <font color='#FF0000'>只能是英文，不能带空格或“\”、“/”等符号。</font><br><font color='#0000FF'>样例：</font>News、Article、Soft</td>"
    Response.Write "      <td><input name='ChannelDir' type='text' id='ChannelDir' value='" & rsChannel("ChannelDir") & "' size='20' maxlength='50' disabled>"
    If rsChannel("ChannelType") <= 1 Then Response.Write "<input name='ChannelDir' type='hidden' id='ChannelDir' value='" & rsChannel("ChannelDir") & "'>"
    Response.Write "<font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    If rsChannel("ModuleType") <> 4 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>频道链接地址（子域名）：</font></strong><br>如果您要在前台将此频道做为主站的一个<font color='red'>独立子站点</font>来访问，请输入完整的网址（如：http://news.powereasy.net）；否则，请保持为空。</td>"
        Response.Write "      <td><input name='ChannelUrl' type='input' size='30' maxlength='100'"
        If SiteUrlType = 0 Then
            Response.Write " disabled"
        Else
            Response.Write " value='" & rsChannel("LinkUrl") & "'"
        End If
        Response.Write "> <font color='red'>* 不能带目录</font><br>如果要启用此功能，必须在“网站选项”中将“链接地址方式”改为“绝对路径”</td>"
        Response.Write "    </tr>" & vbCrLf
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>项目名称：</strong><br>例如：频道名称为“网络学院”，其项目名称为“文章”或“教程”</td>"
    Response.Write "      <td><input name='ChannelShortName' type='text' id='ChannelShortName' size='20' maxlength='30' value='" & rsChannel("ChannelShortName") & "'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' class='tdbg5'><strong>项目单位：</strong><br>例如：“篇”、“条”、“个”</td>"
    Response.Write "      <td><input name='ChannelItemUnit' type='text' id='ChannelItemUnit' size='10' maxlength='30' value='" & rsChannel("ChannelItemUnit") & "'></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "       </table>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='200' class='tdbg5'><strong>打开方式：</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='OpenType' " & RadioValue(rsChannel("OpenType"), 0) & ">在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='OpenType' type='radio' " & RadioValue(rsChannel("OpenType"), 1) & ">在新窗口打开</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='200' class='tdbg5'><strong>禁用本频道：</strong></td>"
    Response.Write "      <td><input name='Disabled' type='radio' " & RadioValue(rsChannel("Disabled"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='Disabled' type='radio' " & RadioValue(rsChannel("Disabled"), False) & ">否</td>"
    Response.Write "    </tr>" & vbCrLf
    If rsChannel("ModuleType") = 4 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>是否在频道栏显示频道名称：</strong></td>"
        Response.Write "      <td><input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), False) & ">否</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>是否在导航栏显示频道名称：</strong></td>"
        Response.Write "      <td><input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), False) & ">否</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>是否启用留言的审核功能：</strong></td>"
        Response.Write "      <td>"
        Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 1) & ">是&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 0) & ">否&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "        <br><font color='#0000FF'>设定为需要审核时，如果某组会员有“发表信息不需审核”的特权，则此会员组不受此限。</font>"
        Response.Write "      </td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>本频道的默认风格：</strong></td>"
        Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(rsChannel("DefaultSkinID")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
    End If
    Response.Write "  </tbody>" & vbCrLf

    '刘永涛 屏蔽与供求及人才无关的频道
    If rsChannel("ModuleType") = 6 Or rsChannel("ModuleType") = 7 Or rsChannel("ModuleType") = 8 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        If rsChannel("ModuleType") = 6 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>频道推荐的收费设置：</strong><br>信息在此频道设为推荐，每天要扣除的会员点数.</td>"
            Response.Write "      <td>"
            Response.Write "        推荐信息扣除&nbsp;<INPUT TYPE='text' NAME='CommandChannelPoint' MaxLength='5' Size='5' Value='" & PE_CLng(rsChannel("CommandChannelPoint")) & "'>&nbsp;点数/天"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        If rsChannel("ModuleType") <> 8 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>本频道的审核级别：</strong><br>设定为需要审核时，如果某组会员有“发表信息不需审核”的特权，则此会员组不受此限。</td>"
            Response.Write "      <td>"
            Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 0) & ">不需审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>在本频道发表信息不需要管理员审核</font><br>"
            Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 1) & ">一级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>需要栏目审核员进行审核（注：此级别为最小级别）</font><br>"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='200' class='tdbg5'><strong>频道META关键词：</strong><br>针对搜索引擎设置的关键词<br>多个关键词请用,号分隔</td>" & vbCrLf
        Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsChannel("Meta_Keywords") & "</textarea></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='200' class='tdbg5'><strong>频道META网页描述：</strong><br>针对搜索引擎设置的网页描述<br>多个描述请用,号分隔</td>" & vbCrLf
        Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsChannel("Meta_Description") & "</textarea></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        If rsChannel("ModuleType") <> 4 Then
            Response.Write "  <tbody id='Tabs' style='display:none'>"
            If rsChannel("ModuleType") <> 5 Then
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>频道权限：</strong><br><font color='red'>频道权限为继承关系，当频道设为“认证频道”时，其下的栏目设为“开放栏目”也无效。相反，如果频道设为“开放频道”，其下的栏目可以自由设置权限。</font></td>"
                Response.Write "      <td>"
                Response.Write "        <table>"
                Response.Write "     <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' " & RadioValue(rsChannel("ChannelPurview"), 0) & ">开放频道</td><td>任何人（包括游客）可以浏览此频道下的信息。可以在栏目设置中再指定具体的栏目权限。</td></tr>"
                Response.Write "     <tr><td width='80' valign='top'><input type='radio' name='ChannelPurview' " & RadioValue(rsChannel("ChannelPurview"), 1) & ">认证频道</td><td>游客不能浏览，并在下面指定允许浏览的会员组。如果频道设置为认证频道，则此频道的“生成HTML”选项只能设为“不生成HTML”。</td></tr>"
                Response.Write "        </table>"
                Response.Write "      </td>"
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>允许浏览此频道的会员组：</strong><br>如果频道权限设置为“认证频道”，请在此设置允许浏览此频道的会员组</td>"
                Response.Write "      <td>" & GetUserGroup(rsChannel("arrGroupID") & "", "") & "</td>"
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>本频道的审核级别：</strong><br>设定为需要审核时，如果某组会员有“发表信息不需审核”的特权，则此会员组不受此限。</td>"
                Response.Write "      <td>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 0) & ">不需审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>在本频道发表信息不需要管理员审核</font><br>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 1) & ">一级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>需要栏目审核员进行审核（注：此级别为最小级别，下同）</font><br>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 2) & ">二级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>初审需要栏目审核员审核，终审需要栏目总编进行审核</font><br>"
                Response.Write "        <input name='CheckLevel' type='radio' " & RadioValue(rsChannel("CheckLevel"), 3) & ">三级审核&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>初审需要栏目审核员审核，二审需要栏目总编审核，终审需要频道管理员审核</font>"
                Response.Write "      </td>"
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='300' class='tdbg5'><strong>不属于任何栏目的评论权限：</strong></td>"
                Response.Write "      <td>"
                Response.Write "        <input name='EnableComment' type='checkbox' value='True' "
                If PE_CBool(rsChannel("EnableComment")) = True Then Response.write "checked"
                Response.Write ">允许发表评论<br>"
                Response.Write "        <input name='CheckComment' type='checkbox' value='True' "
                If PE_CBool(rsChannel("CheckComment")) = True Then Response.write "checked"		
                Response.Write ">评论需要审核"
                Response.Write "      </td>"
                Response.Write "    </tr>" & vbCrLf				
                Response.Write "    <tr class='tdbg'>" & vbCrLf
                Response.Write "      <td width='200' class='tdbg5'><strong>退稿时站内短信/Email通知内容：</strong><br>不支持HTML代码</td>" & vbCrLf
                Response.Write "      <td><textarea name='EmailOfReject' cols='60' rows='4'>" & rsChannel("EmailOfReject") & "</textarea></td>" & vbCrLf
                Response.Write "    </tr>" & vbCrLf
                Response.Write "    <tr class='tdbg'>" & vbCrLf
                Response.Write "      <td width='200' class='tdbg5'><strong>稿件被采用时站内短信/Email通知内容：</strong><br>不支持HTML代码</td>" & vbCrLf
                Response.Write "      <td><textarea name='EmailOfPassed' cols='60' rows='4'>" & rsChannel("EmailOfPassed") & "</textarea></td>" & vbCrLf
                Response.Write "    </tr>" & vbCrLf
            End If
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='200' class='tdbg5'><strong>频道META关键词：</strong><br>针对搜索引擎设置的关键词<br>多个关键词请用,号分隔</td>" & vbCrLf
            Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsChannel("Meta_Keywords") & "</textarea></td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>" & vbCrLf
            Response.Write "      <td width='200' class='tdbg5'><strong>频道META网页描述：</strong><br>针对搜索引擎设置的网页描述<br>多个描述请用,号分隔</td>" & vbCrLf
            Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsChannel("Meta_Description") & "</textarea></td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf

            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>添加/修改信息时的界面设置：</strong></td>"
            Response.Write "      <td>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Charge'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Charge", ",") = True Then Response.Write " checked"
            Response.Write ">显示“收费选项”书签<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Vote'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Vote", ",") = True Then Response.Write " checked"
            Response.Write ">显示“调查设置”书签<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='SoftParameter'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "SoftParameter", ",") = True Then Response.Write " checked"
            Response.Write ">显示“软件参数”书签（仅对下载模块有效）<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Recieve'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Recieve", ",") = True Then Response.Write " checked"
            Response.Write ">显示“签收设置”书签（仅对文章模块有效）<br>"
            Response.Write "<input name='arrEnabledTabs' type='checkbox' value='Copyfee'"
            If FoundInArr(rsChannel("arrEnabledTabs"), "Copyfee", ",") = True Then Response.Write " checked"
            Response.Write ">显示“稿费设置”书签（仅对文章模块有效）<br>"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
            If rsChannel("ModuleType") = 1 Then
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>默认稿费标准：</strong>(单位：元/千字)</td>"
                Response.Write "      <td><input name='MoneyPerKw' type='text'id='MoneyPerKw' size='10' maxlength='10' value='" & rsChannel("MoneyPerKw") & "'> <font color=red>元/千字</font></Td>"
                Response.Write "    </tr>" & vbCrLf
            End If
            Response.Write "  </tbody>" & vbCrLf
        End If
    End If

    '2005-12-23 供求信息频道
    '刘永涛 屏蔽与供求和房产及人才无关的频道
    If rsChannel("ModuleType") = 6 Or rsChannel("ModuleType") = 7 Or rsChannel("ModuleType") = 8 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>本频道的首页模板</strong></td>"
        Response.Write "      <td><select name='Template_Index' id='Template_Index'>" & GetTemplate_Option(iChannelID, 1, rsChannel("Template_Index")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>是否在频道栏显示频道名称：</strong></td>"
        Response.Write "      <td><input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), False) & ">否</td>"
        Response.Write "    </tr>" & vbCrLf
        If rsChannel("ModuleType") <> 8 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>本频道热点的点击数最小值：</strong></td>"
            Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='" & rsChannel("HitsOfHot") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        If rsChannel("ModuleType") = 6 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>频道首页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Index' type='text' id='MaxPerPage_Index' value='" & rsChannel("MaxPerPage_Index") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>搜索结果页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='" & rsChannel("MaxPerPage_SearchResult") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>最新信息页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_New' type='text' id='MaxPerPage_New' value='" & rsChannel("MaxPerPage_New") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>热门信息页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Hot' type='text' id='MaxPerPage_Hot' value='" & rsChannel("MaxPerPage_Hot") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>推荐信息页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Elite' type='text' id='MaxPerPage_Elite' value='" & rsChannel("MaxPerPage_Elite") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>专题列表页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SpecialList' type='text' id='MaxPerPage_SpecialList' value='" & rsChannel("MaxPerPage_SpecialList") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>是否在导航栏显示频道名称：</strong></td>"
        Response.Write "      <td><input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), False) & ">否</td>"
        Response.Write "    </tr>" & vbCrLf
        If rsChannel("ModuleType") = 6 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>多少天内更新的信息为新信息：</strong></td>"
            Response.Write "      <td><input name='DaysOfNew' type='text' id='DaysOfNew' value='" & rsChannel("DaysOfNew") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>顶部导航栏每行显示的栏目数：</strong></td>"
            Response.Write "      <td><input name='MaxPerLine' type='text' id='MaxPerLine' value='" & rsChannel("MaxPerLine") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>顶部栏目菜单的显示方式：</strong><br>更改此参数后需要刷新栏目JS方生效。</td>"
            Response.Write "      <td><select name='TopMenuType' id='TopMenuType'>" & GetMenuType_Option(rsChannel("TopMenuType")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>本频道的默认风格：</strong></td>"
        Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(rsChannel("DefaultSkinID")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        If rsChannel("ModuleType") <> 4 Then
            Response.Write "  <tbody id='Tabs' style='display:none'>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>本频道的首页模板</strong></td>"
            Response.Write "      <td><select name='Template_Index' id='Template_Index'>" & GetTemplate_Option(iChannelID, 1, rsChannel("Template_Index")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>是否在频道栏显示频道名称：</strong></td>"
            Response.Write "      <td><input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowName' type='radio' " & RadioValue(rsChannel("ShowName"), False) & ">否</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>是否在导航栏显示频道名称：</strong></td>"
            Response.Write "      <td><input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowNameOnPath' type='radio' " & RadioValue(rsChannel("ShowNameOnPath"), False) & ">否</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>是否在本频道显示树状导航菜单：</strong></td>"
            Response.Write "      <td><input name='ShowClassTreeGuide' type='radio' " & RadioValue(rsChannel("ShowClassTreeGuide"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowClassTreeGuide' type='radio' " & RadioValue(rsChannel("ShowClassTreeGuide"), False) & ">否</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>前台是否显示标题后的省略号：</strong><br>当模板中指定标题长度小于标题实际长度时，可以决定是否在标题后面显示省略号</td>"
            Response.Write "      <td><input name='ShowSuspensionPoints' type='radio' " & RadioValue(rsChannel("ShowSuspensionPoints"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowSuspensionPoints' type='radio' " & RadioValue(rsChannel("ShowSuspensionPoints"), False) & ">否</td>"
            Response.Write "    </tr>" & vbCrLf
            If rsChannel("ModuleType") <> 5 Then
                Response.Write "    <tr class='tdbg'>"
                Response.Write "      <td width='200' class='tdbg5'><strong>本频道热点的点击数最小值：</strong></td>"
                Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='" & rsChannel("HitsOfHot") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
                Response.Write "    </tr>" & vbCrLf
            End If
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>多少天内更新的信息为新信息：</strong></td>"
            Response.Write "      <td><input name='DaysOfNew' type='text' id='DaysOfNew' value='" & rsChannel("DaysOfNew") & "' size='10' maxlength='10'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>顶部导航栏每行显示的栏目数：</strong></td>"
            Response.Write "      <td><input name='MaxPerLine' type='text' id='MaxPerLine' value='" & rsChannel("MaxPerLine") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>信息列表中作者名称的长度：</strong></td>"
            Response.Write "      <td><input name='AuthorInfoLen' type='text' id='AuthorInfoLen' value='" & rsChannel("AuthorInfoLen") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>频道首页专题列表的数量：</strong></td>"
            Response.Write "      <td><input name='JS_SpecialNum' type='text' id='JS_SpecialNum' value='" & rsChannel("JS_SpecialNum") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>频道首页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Index' type='text' id='MaxPerPage_Index' value='" & rsChannel("MaxPerPage_Index") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>搜索结果页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='" & rsChannel("MaxPerPage_SearchResult") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>最新信息页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_New' type='text' id='MaxPerPage_New' value='" & rsChannel("MaxPerPage_New") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>热门信息页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Hot' type='text' id='MaxPerPage_Hot' value='" & rsChannel("MaxPerPage_Hot") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>推荐信息页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_Elite' type='text' id='MaxPerPage_Elite' value='" & rsChannel("MaxPerPage_Elite") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>专题列表页的每页信息数：</strong></td>"
            Response.Write "      <td><input name='MaxPerPage_SpecialList' type='text' id='MaxPerPage_SpecialList' value='" & rsChannel("MaxPerPage_SpecialList") & "' size='10' maxlength='3'> <font color='#FF0000'>*</font></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>本频道的默认风格：</strong></td>"
            Response.Write "      <td><select name='DefaultSkinID' id='DefaultSkinID'>" & Admin_GetSkin_Option(rsChannel("DefaultSkinID")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>顶部栏目菜单的显示方式：</strong><br>更改此参数后需要刷新栏目JS方生效。</td>"
            Response.Write "      <td><select name='TopMenuType' id='TopMenuType'>" & GetMenuType_Option(rsChannel("TopMenuType")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='200' class='tdbg5'><strong>底部栏目导航的显示方式：</strong><br>更改此参数后需要刷新栏目JS方生效。</td>"
            Response.Write "      <td><select name='ClassGuideType' id='ClassGuideType'>" & GetGuideType_Option(rsChannel("ClassGuideType")) & "</select></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "  </tbody>" & vbCrLf
        End If
    End If

  '  If rsChannel("ModuleType") <> 4 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>是否允许在本频道上传文件：</strong></td>"
        Response.Write "      <td><input name='EnableUploadFile' type='radio' " & RadioValue(rsChannel("EnableUploadFile"), True) & ">是 &nbsp;&nbsp;&nbsp;&nbsp;<input name='EnableUploadFile' type='radio' " & RadioValue(rsChannel("EnableUploadFile"), False) & ">否</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>上传文件的保存目录：</strong><br><font color='red'>你可以定期或不定期的更改上传目录，以防其他网站盗链</font></td>"
        Response.Write "      <td><input name='UploadDir' type='text' id='UploadDir' value='" & rsChannel("UploadDir") & "' size='20' maxlength='20'>&nbsp;&nbsp;<font color='red'>只能是英文和数字，不能带空格或“\”、“/”等符号。</font></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>允许上传的最大文件大小：</strong></td>"
        Response.Write "      <td><input name='MaxFileSize' type='text' id='MaxFileSize' value='" & rsChannel("MaxFileSize") & "' size='10' maxlength='10'> KB&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font></td>"
        Response.Write "    </tr>" & vbCrLf
        Dim arrFileType
        If rsChannel("UpFileType") & "" = "" Then
            arrFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")
        Else
            arrFileType = Split(rsChannel("UpFileType"), "$")
            If UBound(arrFileType) < 4 Then
                arrFileType = Split("gif|jpg|jpeg|jpe|bmp|png$swf$mid|mp3|wmv|asf|avi|mpg$ram|rm|ra$rar|exe|doc|zip", "$")
            End If
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong>允许上传的文件类型：</strong><br>多种文件类型之间以“|”分隔</td>"
        Response.Write "      <td><table>"
        Response.Write "          <tr><td>图片类型：</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(0)) & "' size='50' maxlength='200'></td></tr>"
        Response.Write "          <tr><td>Flash文件：</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(1)) & "' size='50' maxlength='50'></td></tr>"
        Response.Write "          <tr><td>Windows媒体：</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(2)) & "' size='50' maxlength='200'></td></tr>"
        Response.Write "          <tr><td>Real媒体：</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(3)) & "' size='20' maxlength='200'></td></tr>"
        Response.Write "          <tr><td>其他文件：</td><td><input name='UpFileType' type='text' id='UpFileType' value='" & Trim(arrFileType(4)) & "' size='50' maxlength='200'></td></tr>"
        Response.Write "      </table></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
  '  End If

    '刘永涛 屏蔽供求中不用的信息
    If rsChannel("ModuleType") <> 6 Or rsChannel("ModuleType") <> 7 Or rsChannel("ModuleType") <> 4 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>生成HTML方式：</font></strong><br>服务器支持FSO才能启用“生成HTML”功能<br>每一次更改生成方式后，你需要先删除所有以前的文件，再重新生成所有文件。</td>"
        Response.Write "      <td>" & GetUseCreateHTML(rsChannel("UseCreateHTML"), rsChannel("ModuleType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>栏目、专题列表更新页数：</font></strong><br>添加内容后自动更新的栏目及专题列表页数。</td>"
        Response.Write "      <td><input name='UpdatePages' type='text' id='UpdatePages' value='" & rsChannel("UpdatePages") & "' size='5' maxlength='5'> 页 <font color='#FF0000'>*</font>&nbsp;&nbsp;<font color='blue'>如：更新页数设为3，则每次自动更新前三页，第4页以后的分页为固定生成的页面，当新增内容数超过一页，则再生成一个固定页面，在总记录数不是每页记录数的整数倍时，交叉页（第3、4页）会有部分记录重复。</font></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td colspan='2'><font color='red'><b>以下参数仅当“生成HTML方式”设为后三者时才有效。<br>请谨慎选择！在每一次更改以下参数前，你最好先删除所有以前生成的文件，然后在保存参数设置后再重新生成所有文件。</b></font></td></tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>自动生成HTML时的生成方式：</font></strong><br>添加/修改信息时，系统可以自动生成有关页面文件，请在这里选择自动生成时的方式。</td>"
        Response.Write "      <td>" & GetAutoCreateType(rsChannel("AutoCreateType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>栏目列表文件的存放位置：</font></strong></td>"
        Response.Write "      <td>" & GetListFileType(rsChannel("ListFileType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>目录结构方式：</font></strong><br>每一次更改目录结构后，你需要先删除所有以前的文件，再重新生成所有文件。<br>免费版不支持目录结构修改。</td>"
        Response.Write "      <td>" & GetStructureType(rsChannel("StructureType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>内容页文件的命名方式：</font></strong></td>"
        Response.Write "      <td>" & GetFileNameType(rsChannel("FileNameType")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>频道首页的扩展名：</font></strong></td>"
        Response.Write "      <td>" & arrFileExt_Index(rsChannel("FileExt_Index")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>栏目页、专题页的扩展名：</font></strong></td>"
        Response.Write "      <td>" & arrFileExt_List(rsChannel("FileExt_List")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='200' class='tdbg5'><strong><font color=red>内容页的扩展名：</font></strong></td>"
        Response.Write "      <td>" & arrFileExt_Item(rsChannel("FileExt_Item")) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    End If
    '2005-12-23
    If IsCustom_Content = True And rsChannel("ModuleType") <> 4 And rsChannel("ModuleType") <> 6 And rsChannel("ModuleType") <> 7 And rsChannel("ModuleType") <> 8 Then
        Call EditCustom_Content("Modify", rsChannel("Custom_Content"), "Channel")
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center'><tr><td colspan='2' align='center'>"
    If rsChannel("ModuleType") <> 4 Then
        Response.Write "     <br><font color='red'>在更改频道有关参数前，你最好先删除所有以前生成的文件，更改参数后再重新生成所有文件。</font><br><br>" & vbCrLf
    End If
    Response.Write "     <input name='iChannelID' type='hidden' id='iChannelID' value='" & rsChannel("ChannelID") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input name='ModuleType' type='hidden' id='hidden' value='" & rsChannel("ModuleType") & "'>"
    Response.Write "        <input name='Submit'  type='submit' id='Submit' value='保存修改结果'> &nbsp; <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Channel.asp'"" style='cursor:hand;'></td>"
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "</form>"

    rsChannel.Close
    Set rsChannel = Nothing
    Call ShowChekcFormJS
End Sub

Sub SaveAdd()
    Dim rsChannel
    Dim OrderID, ChannelType, LinkUrl

    ChannelName = Trim(Request("ChannelName"))
    ChannelShortName = Trim(Request("ChannelShortName"))
    ChannelItemUnit = Trim(Request("ChannelItemUnit"))
    LinkUrl = Trim(Request("LinkUrl"))
    ChannelType = PE_CLng(Trim(Request("ChannelType")))
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    ChannelDir = Trim(Request("ChannelDir"))
    UploadDir = ReplaceBadChar(Trim(Request("UploadDir")))
    If ChannelName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>频道名称不能为空！</li>"
    Else
        ChannelName = ReplaceBadChar(ChannelName)
    End If
    If ChannelType = 1 Then
        If ChannelDir = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>频道目录不能为空！</li>"
        ElseIf LCase(ChannelDir) = "others" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>“Others”已被系统使用，请更换频道目录名！</li>"
        Else
            If IsValidStr(ChannelDir) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>频道目录名只能为英文字母和数字的组合，且第一个字符必须为英文字母！</li>"
            Else
                If fso.FolderExists(Server.MapPath(InstallDir & ChannelDir)) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>频道目录已经存在！请另外指定一个目录。</li>"
                End If
            End If
        End If
        If ChannelShortName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定项目名称！</li>"
        End If
        If UploadDir = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定上传目录</li>"
        End If
    Else
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址不能为空！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If ChannelItemUnit = "" Then ChannelItemUnit = "个"
    
    ChannelID = GetNewID("PE_Channel", "ChannelID")
    OrderID = GetNewID("PE_Channel", "OrderID")
    
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    rsChannel.Open "Select * from PE_Channel Where ChannelName='" & ChannelName & "'", Conn, 1, 3
    If Not (rsChannel.BOF And rsChannel.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>频道名称已经存在！</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Sub
    End If
    rsChannel.addnew
    rsChannel("ChannelID") = ChannelID
    rsChannel("OrderID") = OrderID
    rsChannel("ChannelName") = ChannelName
    rsChannel("ChannelType") = ChannelType
    If ChannelType = 1 Then
        If SiteUrlType = 0 Then
            rsChannel("LinkUrl") = ""
        Else
            rsChannel("LinkUrl") = Trim(Request("ChannelUrl"))
        End If
    Else
        rsChannel("LinkUrl") = LinkUrl
    End If
    rsChannel("ModuleType") = ModuleType
    rsChannel("ChannelDir") = ChannelDir
    rsChannel("ChannelShortName") = ChannelShortName
    rsChannel("ChannelItemUnit") = ChannelItemUnit
    rsChannel("UploadDir") = UploadDir

    Call SaveChannel(rsChannel)

    rsChannel("ItemCount") = 0
    rsChannel("ItemChecked") = 0
    rsChannel("CommentCount") = 0
    rsChannel("SpecialCount") = 0
    rsChannel("HitsCount") = 0

    '自设内容
    Dim Custom_Num, Custom_Content, i
    Custom_Num = PE_CLng(Request.Form("Custom_Num"))
    If Custom_Num <> 0 Then
        For i = 1 To Custom_Num
            If i <> 1 Then
                Custom_Content = Custom_Content & "{#$$$#}"
            End If
            Custom_Content = Custom_Content & Trim(Request("Custom_Content" & i))
        Next
    End If
    rsChannel("Custom_Content") = Custom_Content
    If ModuleType = 1 then
        Dim rsCheckChannel
        Set rsCheckChannel = Conn.Execute("Select * from PE_MailChannel where ChannelID = "&ChannelID)
        If rsCheckChannel.bof and rsCheckChannel.eof Then
            Conn.Execute ("insert into PE_MailChannel(ChannelID,UserID,arrClass,SendNum,IsUse) values("&ChannelID&",'','',10," & PE_False & ")")
        End If
        rsCheckChannel.Close
        set rsCheckChannel = nothing	
    End If
    rsChannel.Update
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteEntry(2, AdminName, "添加频道成功：" & ChannelName)

    If ChannelType = 1 Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table PE_Admin add AdminPurview_" & ChannelDir & " Int null")
        Else
            Conn.Execute ("alter table PE_Admin add COLUMN AdminPurview_" & ChannelDir & " INTEGER")	
        End If
        Call CreateChannelDir(ChannelID, ChannelDir, UploadDir, ModuleType)
        Call AddTemplate(ModuleType, ChannelID)
        Call AddJsFile(ModuleType, ChannelID)
        Call ReloadLeft("Admin_Channel.asp")
    Else
        Call CloseConn
        Response.Redirect "Admin_Channel.asp"
    End If
End Sub

Sub SaveModify()
    Dim ChannelType, LinkUrl
    Dim rsChannel, sqlChannel
    Dim CommandChannelPoint '发布信息要扣除的点数和设定频道推荐每天要扣除的点数
    ChannelID = Trim(Request("iChannelID"))
    ChannelType = PE_CLng(Trim(Request("ChannelType")))
    ChannelName = Trim(Request("ChannelName"))
    ChannelShortName = Trim(Request("ChannelShortName"))
    ChannelItemUnit = Trim(Request("ChannelItemUnit"))
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    LinkUrl = Trim(Request("LinkUrl"))
    UploadDir = Trim(Request("UploadDir"))
    CommandChannelPoint = PE_CLng(Trim(Request("CommandChannelPoint")))
 
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的频道ID</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If ChannelName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>频道名称不能为空！</li>"
    Else
        ChannelName = ReplaceBadChar(ChannelName)
    End If
    If ChannelType = 1 Then
        If ModuleType <> 4 Then
            If ChannelShortName = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请指定项目名称！</li>"
            End If
            If UploadDir = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请指定上传目录</li>"
            End If
        End If
    Else
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址不能为空！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    sqlChannel = "Select * from PE_Channel Where ChannelID=" & ChannelID
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    rsChannel.Open sqlChannel, Conn, 1, 3
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的频道！</li>"
        rsChannel.Close
        Set rsChannel = Nothing
    Else
        If ChannelType = 1 Then
            If RenameDir(InstallDir & rsChannel("ChannelDir") & "/" & rsChannel("UploadDir"), InstallDir & rsChannel("ChannelDir") & "/" & UploadDir) = True Then
                rsChannel("UploadDir") = UploadDir
            End If
        End If
        rsChannel("ChannelName") = ChannelName
        rsChannel("ChannelShortName") = ChannelShortName
        If ChannelItemUnit <> "" Then
            rsChannel("ChannelItemUnit") = ChannelItemUnit
        End If
        If ChannelType = 1 Then
            If SiteUrlType = 0 Then
                rsChannel("LinkUrl") = ""
            Else
                rsChannel("LinkUrl") = Trim(Request("ChannelUrl"))
            End If
        Else
            rsChannel("LinkUrl") = LinkUrl
        End If
        rsChannel("CommandChannelPoint") = CommandChannelPoint
        Call SaveChannel(rsChannel)

        '自设内容
        Dim Custom_Num, Custom_Content, i
        Custom_Num = PE_CLng(Request.Form("Custom_Num"))
        If Custom_Num <> 0 Then
            For i = 1 To Custom_Num
                If i <> 1 Then
                    Custom_Content = Custom_Content & "{#$$$#}"
                End If
                Custom_Content = Custom_Content & Trim(Request("Custom_Content" & i))
            Next
        End If
        rsChannel("Custom_Content") = Custom_Content

        rsChannel.Update
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    'Call ClearSiteCache(0)
    Call WriteEntry(2, AdminName, "修改频道成功：" & ChannelName)

    If ChannelType = 1 Then
        Call ReloadLeft("Admin_Channel.asp")
    Else
        Call CloseConn
        Response.Redirect "Admin_Channel.asp"
    End If
End Sub

Sub SaveChannel(rsChannel)
    Dim ChannelPurview, UseCreateHTML
    ChannelPurview = PE_CLng(Trim(Request("ChannelPurview")))
    UseCreateHTML = PE_CLng(Trim(Request("UseCreateHTML")))
    If ChannelPurview = 1 Then
        UseCreateHTML = 0
    End If
    rsChannel("ChannelPicUrl") = Trim(Request("ChannelPicUrl"))
    rsChannel("ReadMe") = Trim(Request("ReadMe"))
    rsChannel("OpenType") = PE_CLng(Trim(Request("OpenType")))
    rsChannel("ChannelPurview") = ChannelPurview
    rsChannel("arrGroupID") = Trim(Request("GroupID"))
    rsChannel("CheckLevel") = PE_CLng(Trim(Request("CheckLevel")))
    rsChannel("EmailOfReject") = Trim(Request("EmailOfReject"))
    rsChannel("EmailOfPassed") = Trim(Request("EmailOfPassed"))
    rsChannel("Meta_Keywords") = Trim(Request("Meta_Keywords"))
    rsChannel("Meta_Description") = Trim(Request("Meta_Description"))
    rsChannel("arrEnabledTabs") = ReplaceBadChar(Trim(Request("arrEnabledTabs")))
    rsChannel("MoneyPerKw") = PE_CDbl(Trim(Request("MoneyPerKw")))

    rsChannel("Disabled") = PE_CBool(Trim(Request("Disabled")))
    rsChannel("ShowName") = PE_CBool(Trim(Request("ShowName")))
    rsChannel("ShowNameOnPath") = PE_CBool(Trim(Request("ShowNameOnPath")))
    rsChannel("ShowClassTreeGuide") = PE_CBool(Trim(Request("ShowClassTreeGuide")))
    rsChannel("ShowSuspensionPoints") = PE_CBool(Trim(Request("ShowSuspensionPoints")))
    rsChannel("EnableUploadFile") = PE_CBool(Trim(Request("EnableUploadFile")))

    rsChannel("MaxFileSize") = PE_CLng(Trim(Request("MaxFileSize")))
    rsChannel("UpFileType") = Replace(Trim(Request("UpFileType")), ",", "$")
    rsChannel("HitsOfHot") = PE_CLng(Trim(Request("HitsOfHot")))
    rsChannel("DaysOfNew") = PE_CLng(Trim(Request("DaysOfNew")))
    rsChannel("MaxPerLine") = PE_CLng(Trim(Request("MaxPerLine")))
    rsChannel("AuthorInfoLen") = PE_CLng(Trim(Request("AuthorInfoLen")))
    rsChannel("JS_SpecialNum") = PE_CLng(Trim(Request("JS_SpecialNum")))
    rsChannel("MaxPerPage_Index") = PE_CLng(Trim(Request("MaxPerPage_Index")))
    rsChannel("MaxPerPage_SearchResult") = PE_CLng(Trim(Request("MaxPerPage_SearchResult")))
    rsChannel("MaxPerPage_New") = PE_CLng(Trim(Request("MaxPerPage_New")))
    rsChannel("MaxPerPage_Hot") = PE_CLng(Trim(Request("MaxPerPage_Hot")))
    rsChannel("MaxPerPage_Elite") = PE_CLng(Trim(Request("MaxPerPage_Elite")))
    rsChannel("MaxPerPage_SpecialList") = PE_CLng(Trim(Request("MaxPerPage_SpecialList")))
    rsChannel("Template_Index") = PE_CLng(Trim(Request("Template_Index")))
    rsChannel("DefaultSkinID") = PE_CLng(Trim(Request("DefaultSkinID")))
    rsChannel("TopMenuType") = PE_CLng(Trim(Request("TopMenuType")))
    rsChannel("ClassGuideType") = PE_CLng(Trim(Request("ClassGuideType")))
    rsChannel("UseCreateHTML") = UseCreateHTML
    rsChannel("StructureType") = PE_CLng(Trim(Request("StructureType")))
    rsChannel("ListFileType") = PE_CLng(Trim(Request("ListFileType")))
    rsChannel("FileNameType") = PE_CLng(Trim(Request("FileNameType")))
    rsChannel("AutoCreateType") = PE_CLng(Trim(Request("AutoCreateType")))
    rsChannel("FileExt_Index") = PE_CLng(Trim(Request("FileExt_Index")))
    rsChannel("FileExt_List") = PE_CLng(Trim(Request("FileExt_List")))
    rsChannel("FileExt_Item") = PE_CLng(Trim(Request("FileExt_Item")))
    rsChannel("UpdatePages") = PE_CLng1(Trim(Request("UpdatePages")))
    rsChannel("EnableComment") = PE_CBooL(Trim(Request("EnableComment")))
    rsChannel("CheckComment") = PE_CBooL(Trim(Request("CheckComment")))		
End Sub

Sub CreateChannelDir(iChannelID, DirName, sUploadDir, iModuleType)
    On Error Resume Next
    Dim fsfl, fl, fsfm, fm, strDir
    If Not fso.FolderExists(Server.MapPath(InstallDir & DirName)) Then
        fso.CreateFolder Server.MapPath(InstallDir & DirName)
    End If
    Select Case iModuleType
    Case 1
        strDir = "Article"
    Case 2
        strDir = "Soft"
    Case 3
        strDir = "Photo"
    Case 5
        strDir = "Shop"
    End Select
    Set fsfl = fso.GetFolder(Server.MapPath(InstallDir & strDir))
    For Each fl In fsfl.Files
        If LCase(Left(fl.name, 7)) <> LCase(strDir) And Not IsNumeric(Left(fl.name, InStr(fl.name, ".") - 1)) And GetFileExt(fl.name) = "asp" Then
            fl.Copy Server.MapPath(InstallDir & DirName & "/" & fl.name), True
        End If
    Next
    Set fsfl = Nothing
    
    Set fl = fso.CreateTextFile(Server.MapPath(InstallDir & DirName & "/Channel_Config.asp"), True)
    fl.WriteLine ("<" & "%")
    fl.WriteLine ("ChannelID = " & iChannelID)
    fl.WriteLine ("%" & ">")
    fl.Close
    Set fl = Nothing
    'Set fl = fso.CreateTextFile(Server.MapPath(InstallDir & DirName & "/Index.asp"), True)
    'fl.WriteLine ("<!" & "--#include file=""CommonCode.asp""" & "-->")
    'fl.WriteLine ("<" & "%")
    'fl.WriteLine ("Call PE_" & strDir & ".ShowIndex")
    'fl.WriteLine ("Set PE_" & strDir & " = Nothing")
    'fl.WriteLine ("%" & ">")
    'fl.Close
    'Set fl = Nothing

    If Trim(sUploadDir) <> "" Then
        If Not fso.FolderExists(Server.MapPath(InstallDir & DirName & "/" & Trim(sUploadDir))) Then
            fso.CreateFolder Server.MapPath(InstallDir & DirName & "/" & Trim(sUploadDir))
        End If
    End If
    If Not fso.FolderExists(Server.MapPath(InstallDir & DirName & "/Images")) Then
        fso.CreateFolder Server.MapPath(InstallDir & DirName & "/Images")
    End If
    Set fsfm = fso.GetFolder(Server.MapPath(InstallDir & strDir & "/Images"))
    For Each fm In fsfm.Files
        fm.Copy Server.MapPath(InstallDir & DirName & "/Images/" & fm.name), True
    Next
    fm.Close
    Set fm = Nothing
    Set fsfm = Nothing
End Sub

Function RenameDir(strFolderName, strTargetName)
    RenameDir = False
    On Error Resume Next
    If LCase(strFolderName) = LCase(strTargetName) Then Exit Function
    If Not fso.FolderExists(Server.MapPath(strFolderName)) Then Exit Function
    If fso.FolderExists(Server.MapPath(strTargetName)) Then Exit Function
    fso.MoveFolder Server.MapPath(strFolderName), Server.MapPath(strTargetName)
    If Err Then
        Err.Clear
    Else
        RenameDir = True
    End If
End Function

Sub DelChannelDir(DirName)
    On Error Resume Next
    If IsNull(DirName) Or Trim(DirName) = "" Then Exit Sub
    If fso.FolderExists(Server.MapPath(InstallDir & DirName)) Then
        fso.DeleteFolder Server.MapPath(InstallDir & DirName)
    End If
End Sub

Sub DelChannel()
    On Error Resume Next
    Dim ChannelID, rsChannel, sqlChannel
    ChannelID = Trim(Request("iChannelID"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的频道ID</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If

    sqlChannel = "Select * from PE_Channel Where ChannelID=" & ChannelID
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    rsChannel.Open sqlChannel, Conn, 1, 3
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的频道！</li>"
    Else
        If rsChannel("ChannelType") = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能删除系统频道，如果你不想用此频道，可以禁用此频道。</li>"
        Else
            If rsChannel("ChannelType") = 1 Then
                Call DelChannelDir(rsChannel("ChannelDir"))

                '删除所属频道评论
                Dim Infotable
                If rsChannel("ModuleType") = 1 Then
                    Infotable = "Article"
                ElseIf rsChannel("ModuleType") = 2 Then
                    Infotable = "Photo"
                ElseIf rsChannel("ModuleType") = 3 Then
                    Infotable = "Soft"
                ElseIf rsChannel("ModuleType") = 5 Then
                    Infotable = "Product"
                End If
                
                Dim rsComment
                Set rsComment = Conn.Execute("Select I." & Infotable & "ID,I.ChannelID,C.InfoID from PE_" & Infotable & " I inner join PE_Comment C on I." & Infotable & "ID=C.InfoID where  I.ChannelID=" & ChannelID & "")
               
                Do While Not rsComment.EOF
                    Conn.Execute "delete from PE_Comment where InfoID=" & rsComment("InfoID")
                    rsComment.MoveNext
                Loop
         
                Set rsComment = Nothing

                Dim rs
                Set rs = Conn.Execute("Select FieldName From PE_Field Where ChannelID=" & ChannelID)
                Do While Not rs.EOF
                    Conn.Execute ("alter table PE_" & Infotable & " drop COLUMN " & rs("FieldName") & "")
                    rs.MoveNext
                Loop
                Set rs = Nothing
                Conn.Execute ("alter table PE_Admin drop COLUMN AdminPurview_" & rsChannel("ChannelDir") & "")
                Conn.Execute ("delete from PE_Class where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Special where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Article where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Soft where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Photo where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_JsFile where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Template where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Announce where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Vote where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Author where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_CopyFrom where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Favorite where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Field where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_KeyLink where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_NewKeys where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_Item where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_HistrolyNews where ChannelID=" & ChannelID)
                Conn.Execute ("delete from PE_MailChannel where ChannelID=" & ChannelID)
                
            End If
            rsChannel.Delete
            rsChannel.Update
            Call ReloadLeft("Admin_Channel.asp")
        End If
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteEntry(2, AdminName, "删除频道成功，ChannelID：" & ChannelID)
    'Call ClearSiteCache(0)
End Sub

Sub DisabledChannel(ActionType)
    Dim ChannelID
    ChannelID = Trim(Request("iChannelID"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定频道ID</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If ActionType = 0 Then
        Conn.Execute ("update PE_Channel set Disabled=" & PE_True & " where ChannelID=" & ChannelID)
    Else
        Conn.Execute ("update PE_Channel set Disabled=" & PE_False & " where ChannelID=" & ChannelID)
    End If
    Call ReloadLeft("Admin_Channel.asp")
    'Call ClearSiteCache(0)
End Sub

Function getMoveNum(ByVal ChannelID, ByVal MoveNum)
    Dim sqlChannelList, rsOrder, i
    sqlChannelList = "Select OrderID,ModuleType From PE_Channel Where 1 <> 1 "
    If Not (FoundInArr(AllModules, "Supply", ",")) Then
        sqlChannelList = sqlChannelList & " Or ModuleType = 6"
    End If
    If Not (FoundInArr(AllModules, "Job", ",")) Then
        sqlChannelList = sqlChannelList & " Or  ModuleType = 8"
    End If
    If Not (FoundInArr(AllModules, "House", ",")) Then
        sqlChannelList = sqlChannelList & " Or  ModuleType = 7"
    End If
    Set rsOrder = Server.CreateObject("Adodb.Recordset")
    Dim CurrentOrderID
    CurrentOrderID = PE_CLng(Conn.Execute("Select OrderID From PE_Channel Where ChannelID = " & ChannelID & "")(0))
    rsOrder.Open sqlChannelList, Conn, 1, 1
     i = 0
    Do While Not rsOrder.EOF
        If CurrentOrderID > rsOrder("OrderID") Then
            i = i + 1
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    getMoveNum = MoveNum + i
End Function

Sub UpOrder()
    Dim ChannelID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsChannel
    ChannelID = Trim(Request("iChannelID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cOrderID = PE_CLng(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    MoveNum = getMoveNum(ChannelID, MoveNum)
    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_Channel")
    MaxOrderID = mrs(0) + 1
    '先将当前栏目移至最后，包括子栏目
    Conn.Execute ("update PE_Channel set OrderID=" & MaxOrderID & " where ChannelID=" & ChannelID)
    
    '然后将位于当前栏目以上的栏目的OrderID依次加一，范围为要提升的数字
    sqlOrder = "select * from PE_Channel where OrderID<" & cOrderID & "  order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前栏目已经在最上面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID，包括子栏目
        Conn.Execute ("update PE_Channel set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前栏目从最后移到相应位置，包括子栏目
    Conn.Execute ("update PE_Channel set OrderID=" & tOrderID & " where ChannelID=" & ChannelID)

    Call ReloadLeft("Admin_Channel.asp?Action=Order")
    'Call ClearSiteCache(0)
End Sub

Sub DownOrder()
    Dim ChannelID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsChannel, PrevID, NextID
    ChannelID = Trim(Request("iChannelID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ChannelID = PE_CLng(ChannelID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cOrderID = PE_CLng(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_Channel")
    MaxOrderID = mrs(0) + 1
    '先将当前栏目移至最后，包括子栏目
    Conn.Execute ("update PE_Channel set OrderID=" & MaxOrderID & " where ChannelID=" & ChannelID)
    
    '然后将位于当前栏目以下的栏目的OrderID依次减一，范围为要下降的数字
    sqlOrder = "select * from PE_Channel where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前栏目已经在最下面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID，包括子栏目
        Conn.Execute ("update PE_Channel set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前栏目从最后移到相应位置，包括子栏目
    Conn.Execute ("update PE_Channel set OrderID=" & tOrderID & " where ChannelID=" & ChannelID)
    
    Call ReloadLeft("Admin_Channel.asp?Action=Order")
    'Call ClearSiteCache(0)
End Sub

Sub UpdateData()
    Call UpdateChannelData(PE_CLng(Trim(Request("iChannelID"))))

    Call WriteSuccessMsg("更新频道数据成功！", ComeUrl)
    'Call ClearSiteCache(0)
End Sub

Sub UpdateChannelFiles()
    Dim iChannelID, rsChannel, sqlChannel, DirName, ModuleType, UploadDir
    Dim fsfl, fl, fsfm, fm, strDir
    
    iChannelID = PE_CLng(Trim(Request("iChannelID")))
    If iChannelID > 0 Then
        sqlChannel = "select ChannelID,ChannelDir,ModuleType,UploadDir from PE_Channel where ChannelType = 1 and ChannelID=" & iChannelID & ""
    Else
        sqlChannel = "select ChannelID,ChannelDir,ModuleType,UploadDir from PE_Channel where ChannelType = 1 And ModuleType<>4 And ModuleType<>6 And ModuleType<>7 And ModuleType<>8"
    End If
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        Call CreateChannelDir(rsChannel("ChannelID"), rsChannel("ChannelDir"), rsChannel("UploadDir"), rsChannel("ModuleType"))
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteSuccessMsg("更新频道文件成功！", ComeUrl)
End Sub

Function Admin_GetSkin_Option(iSkinID)
    Dim sqlSkin, rsSkin, strSkin
    If IsNull(iSkinID) Then iSkinID = 0
    strSkin = ""
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Conn.Execute(sqlSkin)
    If rsSkin.BOF And rsSkin.EOF Then
        strSkin = strSkin & "<option value=''>请先添加风格</option>"
    Else
        If iSkinID = 0 Then
            strSkin = strSkin & "<option value='0' selected>使用系统的默认风格</option>"
            Do While Not rsSkin.EOF
                If rsSkin("IsDefault") = True Then
                    strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'>" & rsSkin("SkinName") & "（默认）</option>"
                Else
                    strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'>" & rsSkin("SkinName") & "</option>"
                End If
                rsSkin.MoveNext
            Loop
        Else
            strSkin = strSkin & "<option value='0'>使用系统的默认风格</option>"
            Do While Not rsSkin.EOF
                strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'"
                If rsSkin("SkinID") = iSkinID Then
                    strSkin = strSkin & " selected"
                End If
                strSkin = strSkin & ">" & rsSkin("SkinName")
                If rsSkin("IsDefault") = True Then
                    strSkin = strSkin & "（默认）"
                End If
                strSkin = strSkin & "</option>"
                rsSkin.MoveNext
            Loop
        End If
    End If
    rsSkin.Close
    Set rsSkin = Nothing
    Admin_GetSkin_Option = strSkin
End Function

Sub AddTemplate(ChannelID_Source, ChannelID_Target)
    
    Dim sqlTemplate, rsTemplate, trs
    '以下代码只复制当前方案下的模板，由于在方案切换会丢失模板，暂时禁用掉
    'Dim rsProjectName, ProjectName
    'Set rsProjectName = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    'If rsProjectName.BOF And rsProjectName.EOF Then
    '    Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
    '    Exit Sub
    'Else
    '    ProjectName = rsProjectName("TemplateProjectName")
    'End If
    'Set rsProjectName = Nothing
    'Set trs = Conn.Execute("select * from PE_Template where ChannelID=" & ChannelID_Source & " and ProjectName='" & ProjectName & "'")
    
    sqlTemplate = "select top 1 * from PE_Template"
    Set rsTemplate = Server.CreateObject("adodb.recordset")
    rsTemplate.Open sqlTemplate, Conn, 1, 3

    Set trs = Conn.Execute("select * from PE_Template where ChannelID=" & ChannelID_Source)
    Do While Not trs.EOF
        rsTemplate.addnew
        rsTemplate("ChannelID") = ChannelID_Target
        rsTemplate("TemplateName") = trs("TemplateName")
        rsTemplate("TemplateType") = trs("TemplateType")
        rsTemplate("TemplateContent") = trs("TemplateContent")
        rsTemplate("IsDefault") = trs("IsDefault")
        rsTemplate("ProjectName") = trs("ProjectName")
        rsTemplate("IsDefaultInProject") = trs("IsDefaultInProject")
        rsTemplate("Deleted") = trs("Deleted")
        rsTemplate.Update
        trs.MoveNext
    Loop
    rsTemplate.Close
    Set rsTemplate = Nothing
    Set trs = Nothing
End Sub

Sub AddJsFile(ChannelID_Source, ChannelID_Target)
    Dim sqlJsFile, rsJsFile, trs
    sqlJsFile = "select top 1 * from PE_JsFile"
    Set rsJsFile = Server.CreateObject("adodb.recordset")
    rsJsFile.Open sqlJsFile, Conn, 1, 3
    Set trs = Conn.Execute("select * from PE_JsFile where ChannelID=" & ChannelID_Source)
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = "<li>找不到指定的JsFile</li>"
        Exit Sub
    End If
    Do While Not trs.EOF
        rsJsFile.addnew
        rsJsFile("ChannelID") = ChannelID_Target
        rsJsFile("JsName") = trs("JsName")
        rsJsFile("JsReadme") = trs("JsReadme")
        rsJsFile("JsFileName") = trs("JsFileName")
        rsJsFile("JsType") = trs("JsType")
        rsJsFile("Config") = trs("Config")
        rsJsFile.Update
        trs.MoveNext
    Loop
    rsJsFile.Close
    Set rsJsFile = Nothing
    Set trs = Nothing
End Sub

Sub ReloadLeft(strUrl)
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "  parent.left.location.reload();" & vbCrLf
    Response.Write "  window.location.href='" & strUrl & "';"
    Response.Write "</script>" & vbCrLf
End Sub

Function GetMenuType_Option(MenuType)
    Dim strMenuType
    strMenuType = strMenuType & "<option " & OptionValue(MenuType, 1) & ">无级分类菜单</option>"
    strMenuType = strMenuType & "<option " & OptionValue(MenuType, 2) & ">普通下拉菜单</option>"
    strMenuType = strMenuType & "<option " & OptionValue(MenuType, 3) & ">无菜单</option>"
    GetMenuType_Option = strMenuType
End Function

Function GetGuideType_Option(GuideType)
    Dim strGuideType
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 1) & ">平行式（每行2个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 2) & ">平行式（每行3个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 3) & ">平行式（每行4个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 4) & ">平行式（每行5个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 5) & ">平行式（每行6个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 6) & ">平行式（每行7个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 7) & ">平行式（每行8个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 8) & ">纵列式（一列，每列中显示2个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 9) & ">纵列式（一列，每列中显示3个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 10) & ">纵列式（一列，每列中显示4个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 11) & ">纵列式（一列，每列中显示5个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 12) & ">纵列式（一列，每列中显示6个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 13) & ">纵列式（一列，每列中显示7个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 14) & ">纵列式（一列，每列中显示8个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 15) & ">纵列式（两列，每列中显示2个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 16) & ">纵列式（两列，每列中显示3个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 17) & ">纵列式（两列，每列中显示4个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 18) & ">纵列式（两列，每列中显示5个子栏目）</option>"
    strGuideType = strGuideType & "<option " & OptionValue(GuideType, 19) & ">纵列式（两列，每列中显示6个子栏目）</option>"
    GetGuideType_Option = strGuideType
End Function


Function GetOrderType_Option(OrderType)
    Dim strOrderType
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 1) & ">" & ChannelShortName & "ID（降序）</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 2) & ">" & ChannelShortName & "ID（升序）</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 3) & ">更新时间（降序）</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 4) & ">更新时间（升序）</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 5) & ">点击次数（降序）</option>"
    strOrderType = strOrderType & "<option " & OptionValue(OrderType, 6) & ">点击次数（升序）</option>"
    GetOrderType_Option = strOrderType
End Function

Function GetUseCreateHTML(UseCreateHTML, ModuleType)
    Dim strUseCreateHTML
    strUseCreateHTML = strUseCreateHTML & "<input name='UseCreateHTML' type='radio' value='0'"
    If UseCreateHTML = 0 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " checked"
    strUseCreateHTML = strUseCreateHTML & ">不生成&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>当频道中的信息量比较少（≤1000）时，可以选用此种方式，此方式最耗费系统资源。</font><br>"
    strUseCreateHTML = strUseCreateHTML & "<input type='radio' name='UseCreateHTML' value='1'"
    If UseCreateHTML = 1 And ObjInstalled_FSO = True Then strUseCreateHTML = strUseCreateHTML & " checked"
    If ModuleType = 4 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " disabled "
    strUseCreateHTML = strUseCreateHTML & ">全部生成&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>此方式在生成后将最节省系统资源，但当信息量大时，生成过程将比较长。</font><br>"
    strUseCreateHTML = strUseCreateHTML & "<input type='radio' name='UseCreateHTML' value='2'"
    If UseCreateHTML = 2 And ObjInstalled_FSO = True Then strUseCreateHTML = strUseCreateHTML & " checked"
    If ModuleType = 4 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " disabled "
    strUseCreateHTML = strUseCreateHTML & ">首页和内容页为HTML，栏目和专题页为ASP<br>"
    strUseCreateHTML = strUseCreateHTML & "<input type='radio' name='UseCreateHTML' value='3'"
    If UseCreateHTML = 3 And ObjInstalled_FSO = True Then strUseCreateHTML = strUseCreateHTML & " checked"
    If ModuleType = 4 Or ObjInstalled_FSO = False Then strUseCreateHTML = strUseCreateHTML & " disabled "
    strUseCreateHTML = strUseCreateHTML & ">首页、内容页、栏目和专题的首页为HTML，其他页为ASP <font color='red'><b>（推荐）</b></font>"
    GetUseCreateHTML = strUseCreateHTML
End Function

Function GetAutoCreateType(AutoCreateType)
    Dim strAutoCreateType
    strAutoCreateType = strAutoCreateType & "<input name='AutoCreateType' type='radio' " & RadioValue(AutoCreateType, 0) & ">不自动生成，由管理员手工生成相关页面<br>"
    strAutoCreateType = strAutoCreateType & "<input name='AutoCreateType' type='radio' " & RadioValue(AutoCreateType, 1) & ">自动生成全部所需页面<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>当“生成HTML方式”设置为“全部生成”时，将生成所有页面；当“生成HTML方式”设置为后两种时，会根据设置的选项生成有关页面。</font><br>"
    strAutoCreateType = strAutoCreateType & "<input name='AutoCreateType' type='radio' " & RadioValue(AutoCreateType, 2) & ">自动生成部分所需页面<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>仅当“生成HTML方式”设置为“全部生成”时方有效。此方式只生成首页、内容页、栏目和专题的首页，其他页面由管理员手工生成。</font><br>"
    GetAutoCreateType = strAutoCreateType
End Function

Function GetListFileType(ListFileType)
    Dim strListFileType
    strListFileType = strListFileType & "<input name='ListFileType' type='radio' " & RadioValue(ListFileType, 0) & ">列表文件分目录保存在所属栏目的文件夹中<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/ASP/JiChu/index.html（栏目首页）<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Article/ASP/JiChu/List_2.html（第二页）</font><br>"
    strListFileType = strListFileType & "<input name='ListFileType' type='radio' " & RadioValue(ListFileType, 1) & ">列表文件统一保存在指定的“List”文件夹中<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/List/List_236.html（栏目首页）<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Article/List/List_236_2.html（第二页）</font><br>"
    strListFileType = strListFileType & "<input name='ListFileType' type='radio' " & RadioValue(ListFileType, 2) & ">列表文件统一保存在频道文件夹中<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/List_236.html（栏目首页）<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Article/List_236_2.html（第二页）</font><br>"
    GetListFileType = strListFileType
End Function

Function GetStructureType(StructureType)
    Dim strStructureType
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 0) & ">频道/大类/小类/月份/文件（栏目分级，再按月份保存）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/ASP/JiChu/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 1) & ">频道/大类/小类/日期/文件（栏目分级，再按日期分，每天一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/ASP/JiChu/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 2) & ">频道/大类/小类/文件（栏目分级，不再按月份）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/ASP/JiChu/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 3) & ">频道/栏目/月份/文件（栏目平级，再按月份保存）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/JiChu/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 4) & ">频道/栏目/日期/文件（栏目平级，再按日期分，每天一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/JiChu/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 5) & ">频道/栏目/文件（栏目平级，不再按月份）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/JiChu/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 6) & ">频道/文件（直接放在频道目录中）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 7) & ">频道/HTML/文件（直接放在指定的“HTML”文件夹中）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/HTML/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 8) & ">频道/年份/文件（直接按年份保存，每年一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/2004/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 9) & ">频道/月份/文件（直接按月份保存，每月一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 10) & ">频道/日期/文件（直接按日期保存，每天一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 11) & ">频道/年份/月份/文件（先按年份，再按月份保存，每月一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/2004/200408/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 12) & ">频道/年份/日期/文件（先按年份，再按日期分，每天一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/2004/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 13) & ">频道/月份/日期/文件（先按月份，再按日期分，每天一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/200408/2004-08-25/1368.html</font><br>"
    strStructureType = strStructureType & "<input name='StructureType' type='radio' " & RadioValue(StructureType, 14) & ">频道/年份/月份/日期/文件（先按年份，再按日期分，每天一个目录）<br>&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article/2004/200408/2004-08-25/1368.html</font>"
    GetStructureType = strStructureType
End Function

Function GetFileNameType(FileNameType)
    Dim strFileNameType
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 0) & ">文章ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：1358.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 1) & ">更新时间.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：20040828112308.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 2) & ">频道英文名_文章ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article_1358.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 3) & ">频道英文名_更新时间.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article_20040828112308.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 4) & ">更新时间_ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：20040828112308_1358.html</font><br>"
    strFileNameType = strFileNameType & "<input name='FileNameType' type='radio' " & RadioValue(FileNameType, 5) & ">频道英文名_更新时间_ID.html&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>例：Article_20040828112308_1358.html</font>"
    GetFileNameType = strFileNameType
End Function

Function arrFileExt_Index(FileExt_Index)
    Dim strFileExt_Index
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 0) & ">.html&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 1) & ">.htm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 2) & ">.shtml&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 3) & ">.shtm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Index = strFileExt_Index & "<input name='FileExt_Index' type='radio' " & RadioValue(FileExt_Index, 4) & ">.asp"
    arrFileExt_Index = strFileExt_Index
End Function

Function arrFileExt_List(FileExt_List)
    Dim strFileExt_List
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 0) & ">.html&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 1) & ">.htm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 2) & ">.shtml&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 3) & ">.shtm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_List = strFileExt_List & "<input name='FileExt_List' type='radio' " & RadioValue(FileExt_List, 4) & ">.asp"
    arrFileExt_List = strFileExt_List
End Function

Function arrFileExt_Item(FileExt_Item)
    Dim strFileExt_Item
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 0) & ">.html&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 1) & ">.htm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 2) & ">.shtml&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 3) & ">.shtm&nbsp;&nbsp;&nbsp;&nbsp;"
    strFileExt_Item = strFileExt_Item & "<input name='FileExt_Item' type='radio' " & RadioValue(FileExt_Item, 4) & ">.asp"
    arrFileExt_Item = strFileExt_Item
End Function


Function GetModuleTypeName(ModuleType)
    Dim strModuleType
    Select Case ModuleType
    Case 1
        strModuleType = "文章"
    Case 2
        strModuleType = "下载"
    Case 3
        strModuleType = "图片"
    Case 4
        strModuleType = "留言板"
    Case 5
        strModuleType = "商城"
    Case Else
        strModuleType = "其他"
    End Select
    GetModuleTypeName = strModuleType
End Function
%>
