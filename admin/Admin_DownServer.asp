<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 1   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "DeliverType"   '其他权限

Response.Write "<html><head><title>服务器管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle(ChannelName & "管理----镜像服务器管理", 10123)

Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td colspan='5'>"
Response.Write "    <a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'>服务器管理首页</a>&nbsp;|&nbsp;<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&Action=Add'>添加新服务器</a>&nbsp;|&nbsp;<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&Action=Order'>服务器排序</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveDownServer
Case "Del"
    Call Del
Case "Order"
    Call Order
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "SaveAllShowType"
    Call SaveAllShowType
Case Else
    Call Main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub Main()
    Dim rsDownServer, sqlDownServer
    sqlDownServer = "select * from PE_DownServer where ChannelID=" & ChannelID & " order by OrderID"
    Set rsDownServer = Conn.Execute(sqlDownServer)

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' align='center'>"
    Response.Write "    <td width='40'><strong>ID</strong></td>"
    Response.Write "    <td width='100'><strong>服务器名</strong></td>"
    Response.Write "    <td width='120'><strong>服务器LOGO</strong></td>"
    Response.Write "    <td width='60'><strong>显示方式</strong></td>"
    Response.Write "    <td><strong>服务器地址</strong></td>"
    Response.Write "    <td width='60'><strong>操作</strong></td>"
    Response.Write "  </tr>"
    If rsDownServer.BOF And rsDownServer.EOF Then
        Response.Write "  <tr class='tdbg' align='center' height='50'><td colspan='10'>没有任何下载服务器</td></tr>"
    Else
        Do While Not rsDownServer.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='40' align='center'>" & rsDownServer("ServerID") & "</td>"
            Response.Write "    <td width='100' align='center'><a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&action=Modify&ServerID=" & rsDownServer("ServerID") & "'>" & rsDownServer("ServerName") & "</a></td>"
            Response.Write "    <td width='120' align='center'>"
            If rsDownServer("ServerLogo") <> "" Then
                Response.Write "<img src='" & rsDownServer("ServerLogo") & "'>"
            End If
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>"
            If rsDownServer("ShowType") = 1 Then
                Response.Write "显示LOGO"
            Else
                Response.Write "显示名称"
            End If
            Response.Write "</td>"
            Response.Write "    <td>" & rsDownServer("ServerUrl") & "</td>"
            Response.Write "    <td width='60' align='center'>"
            Response.Write "<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&action=Modify&ServerID=" & rsDownServer("ServerID") & "'>修改</a> "
            Response.Write "<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "&Action=Del&ServerID=" & rsDownServer("ServerID") & "' onClick=""return confirm('确定要删除此服务器信息吗？');"">删除</a>"
            Response.Write "</td></tr>"
            rsDownServer.MoveNext
        Loop
    End If
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"
    Response.Write "  <tr>"
    Response.Write "    <td height='22' align='center'>"
    Response.Write "    <form action='Admin_DownServer.asp?Action=SaveAllShowType' method='post'>批量设置显示方式"
    Response.Write "<select name='ShowType'><option value='0'>显示名称</option><option value='1'>显示LOGO</option></select></select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "<input type='submit' name='Submit' value='修 改'>"
    Response.Write "</form>"
    Response.Write "</td></tr>"
    Response.Write "</table>"
    Response.Write "<br><b><font color=red>注意：</font></b><br>&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>删除某个镜像服务器信息时,与之相关的下载错误信息也将一起被删除掉。</font><br><br>"
    rsDownServer.Close
    Set rsDownServer = Nothing
End Sub

Sub Add()
    Call ShowJS_Soft
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'>镜像服务器管理</a>&nbsp;&gt;&gt;&nbsp;添加镜像服务器</td></tr></table>"
    Response.Write "<form method='post' action='Admin_DownServer.asp' name='form1'>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>收费选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table id='Tabs' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>服务器名称：</strong><br>在此输入在前台显示的镜像服务器名，如广东下载、上海下载等。</td>"
    Response.Write "          <td class='tdbg'><input name='ServerName' type='text' id='ServerName' size='50' maxlength='30'>&nbsp;</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>服务器LOGO：</strong><br>输入服务器LOGO的绝对地址，如http://www.powereasy.net/Soft/Images/ServerLogo.gif</td>"
    Response.Write "          <td class='tdbg'><input name='ServerLogo' type='text' id='ServerLogo' size='50' maxlength='200'>&nbsp;</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>服务器地址：</strong><br>请认真输入正确的服务器地址。<br>如http://www.powereasy.net/这样的地址</td>"
    Response.Write "          <td class='tdbg'><input name='ServerUrl' type='text' id='ServerUrl' size='50' maxlength='200'>&nbsp;</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='350' class='tdbg5'><strong>显示方式：</strong></td>"
    Response.Write "          <td class='tdbg'><select name='ShowType' size=1><option value='0'>显示名称</option><option value='1'>显示LOGO</option></select></td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "      <table id='Tabs' style='display:none' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>阅读权限：</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0' checked>继承栏目权限（当所属栏目为认证栏目时，建议选择此项）<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'>所有会员（当所属栏目为开放栏目，想单独对某些文章进行查看权限设置，可以选择此项）<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'>指定会员组（当所属栏目为开放栏目，想单独对某些文章进行查看权限设置，可以选择此项）<br>"
    Response.Write GetUserGroup("", "")
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>阅读点数：</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & Session("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'> "
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>如果点数大于0，则有权限的会员阅读此" & ChannelShortName & "时将消耗相应点数（设为9999时除外），游客将无法查看此" & ChannelShortName & "</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>重复收费：</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0' checked>不重复收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'>距离上次收费时间 <input name='PitchTime' type='text' value='24' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'>会员重复查看此文章 <input name='ReadTimes' type='text' value='10' size='8' maxlength='8' style='text-align:center'> 次后重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'>上述两者都满足时重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'>上述两者任一个满足时就重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'>每阅读一次就重复收费一次（建议不要使用）"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>分成比例：</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='0' size='5' maxlength='4' style='text-align:center'> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>如果比例大于0，则将按比例把向阅读者收取的点数支付给录入者</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "      </table>"
    Response.Write "      <br><br><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "      <input  type='submit' name='Submit' value=' 添 加 '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "      <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'""  style='cursor:hand;'>" & vbCrLf
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub ShowJS_Soft()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "    if(ID==0){" & vbCrLf
    Response.Write "      editor.yToolbarsCss();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub SaveDownServer()
    Dim ServerID, ServerName, ServerUrl, ServerLogo, OrderID
    Dim rsDownServer, sqlDownServer
    Dim ShowType
    ServerID = PE_CLng(Trim(Request("ServerID")))
    ServerName = Trim(Request.Form("ServerName"))
    ServerUrl = Trim(Request.Form("ServerUrl"))
    ServerLogo = Trim(Request.Form("ServerLogo"))
    ShowType = PE_CLng(Request.Form("ShowType"))

    If ChannelID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>频道ID丢失！</li>"
    End If
    If ServerName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>服务器名不能为空！</li>"
    End If
    If ServerUrl = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>服务器地址不能为空！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    If Action = "SaveAdd" Then
        ServerID = GetNewID("PE_DownServer", "ServerId")
        OrderID = GetNewID("PE_DownServer", "OrderID")
        
        Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
        rsDownServer.Open "Select top 1 * from PE_DownServer", Conn, 1, 3
        rsDownServer.addnew
        'rsDownServer("ServerID") = ServerID
        'rsDownServer("OrderID") = OrderID
    Else
        If ServerID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要修改的服务器ID！</li>"
            Exit Sub
        End If
        sqlDownServer = "Select * from PE_DownServer Where ServerID=" & ServerID
        Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
        rsDownServer.Open sqlDownServer, Conn, 1, 3
        If rsDownServer.BOF And rsDownServer.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的服务器，可能已经被删除！</li>"
            rsDownServer.Close
            Set rsDownServer = Nothing
            Exit Sub
        End If
    End If
    rsDownServer("ChannelID") = ChannelID
    rsDownServer("ServerName") = ServerName
    rsDownServer("ServerUrl") = ServerUrl
    rsDownServer("ServerLogo") = ServerLogo
    rsDownServer("ShowType") = ShowType

    rsDownServer("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
    rsDownServer("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
    rsDownServer("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
    rsDownServer("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
    rsDownServer("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
    rsDownServer("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
    rsDownServer("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))

    rsDownServer.Update
    rsDownServer.Close
    Set rsDownServer = Nothing
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?ChannelID=" & ChannelID & ""
End Sub


Sub Modify()
    Dim ServerID, rsDownServer, sqlDownServer
    ServerID = PE_CLng(Trim(Request("ServerID")))
    If ServerID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的服务器ID！</li>"
        Exit Sub
    End If
    sqlDownServer = "Select * from PE_DownServer Where ServerID=" & ServerID
    Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
    rsDownServer.Open sqlDownServer, Conn, 1, 3
    If rsDownServer.BOF And rsDownServer.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的服务器，可能已经被删除！</li>"
        Exit Sub
    End If

    Call ShowJS_Soft
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'>镜像服务器管理</a>&nbsp;&gt;&gt;&nbsp;修改镜像服务器设置</td></tr></table>"
    Response.Write "<form method='post' action='Admin_DownServer.asp' name='form1'>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>收费选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table id='Tabs' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>服务器名称：</strong><br>在此输入在前台显示的镜像服务器名，如佛山下载、广州下载等。</td>"
    Response.Write "      <td class='tdbg'><input name='ServerName' type='text' id='ServerName' size='50' maxlength='30' value='" & rsDownServer("ServerName") & "'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>服务器LOGO：</strong><br>输入服务器LOGO的绝对地址，如http://www.powereasy.net/Soft/Images/ServerLogo.gif</td>"
    Response.Write "      <td class='tdbg'><input name='ServerLogo' type='text' id='ServerLogo' size='50' maxlength='200' value='" & rsDownServer("ServerLogo") & "'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>服务器地址：</strong><br>请认真输入正确的服务器地址。<br>如http://www.powereasy.net/这样的地址</td>"
    Response.Write "      <td class='tdbg'><input name='ServerUrl' type='text' id='ServerUrl' size='50' maxlength='200' value='" & rsDownServer("ServerUrl") & "'>&nbsp;</td>"
    Response.Write "    </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>显示方式：</strong></td>"
    Response.Write "      <td class='tdbg'><select name='ShowType'><option value='0'"
    If rsDownServer("ShowType") = 0 Then Response.Write " selected"
    Response.Write ">显示名称</option>"
    Response.Write "<option value='1'"
    If rsDownServer("ShowType") = 1 Then Response.Write " selected"
    Response.Write ">显示LOGO</option>"
    Response.Write "</select>"
    Response.Write "</td></tr>"
    Response.Write "      </table>"
    Response.Write "      <table id='Tabs' style='display:none' width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>阅读权限：</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0'"
    If rsDownServer("InfoPurview") = 0 Then Response.Write " checked"
    Response.Write ">继承栏目权限（当所属栏目为认证栏目时，建议选择此项）<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'"
    If rsDownServer("InfoPurview") = 1 Then Response.Write " checked"
    Response.Write ">所有会员（当所属栏目为开放栏目，想单独对某些文章进行查看权限设置，可以选择此项）<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'"
    If rsDownServer("InfoPurview") = 2 Then Response.Write " checked"
    Response.Write ">指定会员组（当所属栏目为开放栏目，想单独对某些文章进行查看权限设置，可以选择此项）<br>"
    Response.Write GetUserGroup(rsDownServer("arrGroupID"), "")
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "阅读点数：</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & rsDownServer("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'" & ">&nbsp;&nbsp;&nbsp;&nbsp; <font color='#0000FF'>如果大于0，则会员阅读此" & ChannelShortName & "时将消耗相应点数（设为9999时除外），游客将无法查看此" & ChannelShortName & "。</font></td>"
    Response.Write "          </tr>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>重复收费：</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0'"
    If rsDownServer("ChargeType") = 0 Then Response.Write " checked"
    Response.Write ">不重复收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'"
    If rsDownServer("ChargeType") = 1 Then Response.Write " checked"
    Response.Write ">距离上次收费时间 <input name='PitchTime' type='text' value='" & rsDownServer("PitchTime") & "' size='8' maxlength='8' style='text-align:center'" & "> 小时后重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'"
    If rsDownServer("ChargeType") = 2 Then Response.Write " checked"
    Response.Write ">会员重复查看此文章 <input name='ReadTimes' type='text' value='" & rsDownServer("ReadTimes") & "' size='8' maxlength='8' style='text-align:center'" & "> 次后重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'"
    If rsDownServer("ChargeType") = 3 Then Response.Write " checked"
    Response.Write ">上述两者都满足时重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'"
    If rsDownServer("ChargeType") = 4 Then Response.Write " checked"
    Response.Write ">上述两者任一个满足时就重新收费<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'"
    If rsDownServer("ChargeType") = 5 Then Response.Write " checked"
    Response.Write ">每阅读一次就重复收费一次（建议不要使用）"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>分成比例：</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='" & rsDownServer("DividePercent") & "' size='5' maxlength='4' style='text-align:center'" & "> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>如果比例大于0，则将按比例把向阅读者收取的点数支付给录入者</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "      </table>"
    Response.Write "      <input name='ServerID' type='hidden' id='ServerID' value='" & rsDownServer("ServerID") & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "      <input  type='submit' name='Submit' value=' 保存修改结果 '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "      <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_DownServer.asp?ChannelID=" & ChannelID & "'""  style='cursor:hand;'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    rsDownServer.Close
    Set rsDownServer = Nothing
End Sub

Sub Del()
    Dim ServerID, iOrderID
    Dim rs, sql
    ServerID = Trim(Request("ServerID"))
    If ServerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的服务器ID！</li>"
        Exit Sub
    Else
        ServerID = PE_CLng(ServerID)
    End If
    If FoundErr = True Then
    Exit Sub
    End If
    sql = "select OrderID from PE_DownServer where ServerID=" & ServerID
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Open sql, Conn, 1, 3
    If rs.BOF Or rs.EOF Then
    FoundErr = True
        ErrMsg = ErrMsg & "<li>所指定的参数的记录不存在或已被删除！</li>"
        Exit Sub
    Else
        iOrderID = rs("OrderID")
    End If
    '删除下载错误信息表PE_DownError中属于该镜像服务器的报错信息
    Dim rsDownError, sqlDownError
    Dim UrlID
    Set rsDownError = Server.CreateObject("ADODB.Recordset")
    sqlDownError = "select D.ErrorID,S.DownloadUrl from PE_DownError D left join PE_Soft S on D.InfoID=S.SoftID where D.UrlID=" & ServerID
    rsDownError.Open sqlDownError, Conn, 1, 3
    Do While Not rsDownError.EOF
        If InStr(rsDownError("DownloadUrl"), "@@@") > 0 Then
            Conn.Execute ("delete from PE_DownError where ErrorID =" & rsDownError("ErrorID"))
        End If
        rsDownError.MoveNext
    Loop
    rsDownError.Close
    Set rsDownError = Nothing
    Conn.Execute ("update PE_DownServer set OrderID=OrderID-1 where OrderID>" & iOrderID)
    Conn.Execute ("delete from PE_DownServer where ServerID=" & ServerID)
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?ChannelID=" & ChannelID & ""
End Sub

Sub Order()
    Dim iCount, i, j
    Dim rs, sql
    Set rs = Server.CreateObject("Adodb.RecordSet")
    sql = "select * from PE_DownServer where ChannelID=" & ChannelID & " Order by OrderID"
    rs.Open sql, Conn, 1, 1
    iCount = rs.RecordCount
    j = 1
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4' align='center'><strong>服务器排序</strong></td>"
    Response.Write "  </tr>"
    Do While Not rs.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""> "
        Response.Write "    <td align='center'>" & rs("ServerName") & "</td>"
        Response.Write "    <form action='Admin_DownServer.asp?Action=UpOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If j > 1 Then
            Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=ServerID value=" & rs("ServerID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rs("OrderID") & ">&nbsp;<input type=submit name=Submit value='修改'>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form>"
        Response.Write "    <form action='Admin_DownServer.asp?Action=DownOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If iCount > j Then
            Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=ServerID value=" & rs("ServerID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rs("OrderID") & ">&nbsp;<input type=submit name=Submit value='修改'>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "      <td width='200' align='center'>&nbsp;</td>"
        Response.Write "    </form>"
        Response.Write "  </tr>"
        j = j + 1
        rs.MoveNext
    Loop
    Response.Write "</table> "
    rs.Close
    Set rs = Nothing
End Sub


Sub UpOrder()
    Dim ServerID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rs
    ServerID = Trim(Request("ServerID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ServerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ServerID = PE_CLng(ServerID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cOrderID = CInt(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        MoveNum = CInt(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_DownServer")
    MaxOrderID = mrs(0) + 1
    '先将当前服务器移至最后
    Conn.Execute ("update PE_DownServer set OrderID=" & MaxOrderID & " where ServerID=" & ServerID)
    
    '然后将位于当前服务器以上的服务器的OrderID依次加一，范围为要提升的数字
    sqlOrder = "select * from PE_DownServer where OrderID<" & cOrderID & " order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前服务器已经在最上面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID
        Conn.Execute ("update PE_DownServer set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前服务器从最后移到相应位置
    Conn.Execute ("update PE_DownServer set OrderID=" & tOrderID & " where ServerID=" & ServerID)
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?Action=Order&ChannelID=" & ChannelID
End Sub

Sub DownOrder()
    Dim ServerID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rs, PrevID, NextID
    ServerID = Trim(Request("ServerID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If ServerID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ServerID = PE_CLng(ServerID)
    End If
    If cOrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cOrderID = CInt(cOrderID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        MoveNum = CInt(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要提升的数字！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim mrs, MaxOrderID
    Set mrs = Conn.Execute("select max(OrderID) from PE_DownServer")
    MaxOrderID = mrs(0) + 1
    '先将当前服务器移至最后
    Conn.Execute ("update PE_DownServer set OrderID=" & MaxOrderID & " where ServerID=" & ServerID)
    
    '然后将位于当前服务器以下的前服务器的OrderID依次减一，范围为要下降的数字
    sqlOrder = "select * from PE_DownServer where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前服务器已经在最下面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID
        Conn.Execute ("update PE_DownServer set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前服务器从最后移到相应位置
    Conn.Execute ("update PE_DownServer set OrderID=" & tOrderID & " where ServerID=" & ServerID)
    Call CloseConn
    Response.Redirect "Admin_DownServer.asp?Action=Order&ChannelID=" & ChannelID
End Sub

'保存批量设置下载地址显示方式(logo还是名称)
Sub SaveAllShowType()
    Dim rsDownServer, sqlDownServer
    Dim ShowType, ChannelID
    ShowType = Trim(Request("ShowType"))
    ChannelID = PE_CLng(Request("ChannelID"))

    If ShowType = "" Then
       ShowType = "False"
    End If

    sqlDownServer = "Select * from PE_DownServer where ChannelID=" & ChannelID
    Set rsDownServer = Server.CreateObject("Adodb.RecordSet")
    rsDownServer.Open sqlDownServer, Conn, 1, 3
    If rsDownServer.BOF And rsDownServer.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到相关的服务器信息，可能已经被删除！</li>"
        rsDownServer.Close
        Set rsDownServer = Nothing
    Else
        Do While Not rsDownServer.EOF
            rsDownServer("ShowType") = ShowType
            rsDownServer.Update
            rsDownServer.MoveNext
        Loop
        rsDownServer.Close
        Set rsDownServer = Nothing
        Call CloseConn
        Response.Redirect "Admin_DownServer.asp?ChannelID=" & ChannelID & ""
    End If
End Sub
%>
