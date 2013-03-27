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
Const PurviewLevel_Others = "UserGroup"   '其他权限

Dim rsUserGroup, GroupSetting

GroupID = PE_CLng(Trim(Request("GroupID")))

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>会员组管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("会 员 组 管 理", 10042)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'>管理导航：</td>" & vbCrLf
Response.Write "    <td height='30'><a href='?'>会员组管理首页</a>&nbsp;|&nbsp;<a href='?Action=Add'>新增会员组</a> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveGroup
Case "Del"
    Call Del
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim strSql, i
    strSql = "SELECT GroupID,GroupName,GroupIntro,GroupType,GroupSetting,arrClass_Browse,arrClass_View,arrClass_Input FROM PE_UserGroup ORDER by GroupType asc,GroupID asc"
    Set rsUserGroup = Server.CreateObject("adodb.recordset")
    rsUserGroup.Open strSql, Conn, 1, 1
    If rsUserGroup.BOF And rsUserGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，数据库中没有找到任何会员组资料，您的数据库已经损坏，请从默认数据库中导入PE_UserGroup表。</li>"
        rsUserGroup.Close
        Set rsUserGroup = Nothing
        Exit Sub
    End If
    
    totalPut = rsUserGroup.recordcount
    CurrentPage = Trim(Request("page"))
    If CurrentPage = "" Then
        CurrentPage = 1
    Else
        CurrentPage = PE_CLng(CurrentPage)
    End If
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 20
    
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
            rsUserGroup.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' height='22' class='title'>" & vbCrLf
    Response.Write "    <td width='35'>ID</td>" & vbCrLf
    Response.Write "    <td width='120'>会员组名</td>" & vbCrLf
    Response.Write "    <td>会员组简介</td>" & vbCrLf
    Response.Write "    <td width='120'>会员组类型</td>" & vbCrLf
    Response.Write "    <td width='60'>会员数量</td>" & vbCrLf
    Response.Write "    <td width='150'>操 作</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    Dim UserGroupNum
    UserGroupNum = 0
    Do While Not rsUserGroup.EOF
        Response.Write "     <tr align='center' class='tdbg' onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbg2'"">" & vbCrLf
        Response.Write "    <td width='35'>" & rsUserGroup("GroupID") & "</td>" & vbCrLf
        Response.Write "    <td width='120'>" & rsUserGroup("GroupName") & "</td>" & vbCrLf
        Response.Write "    <td align='left'>" & rsUserGroup("GroupIntro") & "</td>" & vbCrLf
        Response.Write "    <td width='120'>" & GetGroupType(rsUserGroup("GroupType")) & "</td>" & vbCrLf
        Response.Write "    <td width='60'>" & GetGroupNum(rsUserGroup("GroupID")) & "</td>" & vbCrLf
        Response.Write "    <td width='150'><a href='Admin_UserGroup.asp?Action=Modify&GroupID=" & rsUserGroup("GroupID") & "'>修改</a>"
        If rsUserGroup("GroupType") > 2  and rsUserGroup("GroupType") <> 5 Then
            Response.Write " | <a href='?Action=Del&GroupID=" & rsUserGroup("GroupID") & "' onclick=""return confirm('确实要删除此会员组吗？');"">删除</a>" & vbCrLf
        Else
            Response.Write " | <font color='#CCCCCC'>删除</font>"
        End If
        Response.Write " | <a href='Admin_User.asp?SearchType=11&GroupID=" & rsUserGroup("GroupID") & "'>列出会员</a></td>"
        Response.Write "    </tr>" & vbCrLf
        rsUserGroup.MoveNext
        UserGroupNum = UserGroupNum + 1
        If UserGroupNum >= MaxPerPage Then Exit Do
    Loop
    rsUserGroup.Close
    Set rsUserGroup = Nothing

    Response.Write "</table>" & vbCrLf
    Response.Write ShowPage("Admin_UserGroup.asp", totalPut, MaxPerPage, CurrentPage, True, True, "个会员组", True)
End Sub

Function GetGroupType(GroupType)
    Select Case GroupType
    Case 0
        GetGroupType = "等待邮件验证会员"
    Case 1
        GetGroupType = "等待管理员审核会员"
    Case 2
        GetGroupType = "默认会员组"
    Case 3
        GetGroupType = "注册会员"
    Case 4
        GetGroupType = "代 理 商"
    Case 5
        GetGroupType = "匿名投稿"        
    Case Else
        GetGroupType = "未知会员组"
    End Select
End Function

Sub ShowJS_Check()
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function GetClassPurview(){" & vbCrLf
    Dim rsChannel, ChannelDir
    If PE_Clng(GroupID)<>-1 Then
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType<>4 And ModuleType<>5 And ModuleType<>7 And ModuleType<>8 And Disabled=" & PE_False & " ORDER BY OrderID")
    Else
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType = 1 And Disabled=" & PE_False & " ORDER BY OrderID")
    End If    
    
    Do While Not rsChannel.EOF
        ChannelDir = rsChannel(0)
        Response.Write "if(document.form1." & ChannelDir & "purview[1].checked==true){" & vbCrLf
        Response.Write "  document.form1.arrClass_Browse_" & ChannelDir & ".value='';" & vbCrLf
        Response.Write "  document.form1.arrClass_View_" & ChannelDir & ".value='';" & vbCrLf
        Response.Write "  document.form1.arrClass_Input_" & ChannelDir & ".value='';" & vbCrLf
        Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Browse.length;i++){" & vbCrLf
        Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Browse[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_Browse[i].checked==true){" & vbCrLf
        Response.Write "             if (document.form1.arrClass_Browse_" & ChannelDir & ".value=='')" & vbCrLf
        Response.Write "                 document.form1.arrClass_Browse_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Browse[i].value;" & vbCrLf
        Response.Write "             else" & vbCrLf
        Response.Write "                 document.form1.arrClass_Browse_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Browse[i].value;" & vbCrLf
        Response.Write "         }" & vbCrLf
        Response.Write "     }" & vbCrLf
        Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_View.length;i++){" & vbCrLf
        Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_View[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_View[i].checked==true){" & vbCrLf
        Response.Write "             if (document.form1.arrClass_View_" & ChannelDir & ".value=='')" & vbCrLf
        Response.Write "                 document.form1.arrClass_View_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
        Response.Write "             else" & vbCrLf
        Response.Write "                 document.form1.arrClass_View_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_View[i].value;" & vbCrLf
        Response.Write "         }" & vbCrLf
        Response.Write "     }" & vbCrLf
        Response.Write "     for(var i=0;i<frm" & ChannelDir & ".document.myform.Purview_Input.length;i++){" & vbCrLf
        Response.Write "         if (frm" & ChannelDir & ".document.myform.Purview_Input[i].disabled==false&&frm" & ChannelDir & ".document.myform.Purview_Input[i].checked==true){" & vbCrLf
        Response.Write "             if (document.form1.arrClass_Input_" & ChannelDir & ".value=='')" & vbCrLf
        Response.Write "                 document.form1.arrClass_Input_" & ChannelDir & ".value=frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
        Response.Write "             else" & vbCrLf
        Response.Write "                 document.form1.arrClass_Input_" & ChannelDir & ".value+=','+frm" & ChannelDir & ".document.myform.Purview_Input[i].value;" & vbCrLf
        Response.Write "         }" & vbCrLf
        Response.Write "     }" & vbCrLf
        Response.Write "  }" & vbCrLf
        rsChannel.MoveNext
    Loop
    Set rsChannel = Nothing
    Response.Write "}" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "  if(document.form1.GroupName.value==''){" & vbCrLf
    Response.Write "      alert('会员组名称不能为空！');" & vbCrLf
    Response.Write "   document.form1.GroupName.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    GetClassPurview();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub
Sub Add()
    Call ShowJS_Check
    Response.Write "<form method='post' action='Admin_UserGroup.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='3'><div align='center'>新 增 会 员 组</div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>会员组名称：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupName' type='text' id='GroupName' size='20' maxlength='20'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>会员组说明：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupIntro' type='text' id='GroupIntro' size='50' maxlength='200'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>组 类 型：</td>" & vbCrLf
    Response.Write "      <td><select name='GroupType' id='GroupType'>" & vbCrLf
    Response.Write "                            <option value='3'>注册会员</option>" & vbCrLf
    Response.Write "                            <option value='4'>代 理 商</option>" & vbCrLf
    Response.Write "                        </select><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>发布权限：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting1' type='checkbox' value='1'>在发布信息需要审核的频道，此组会员发布信息不需要审核<br>" & vbCrLf
    Response.Write "<input name='GroupSetting2' type='checkbox' value='1'>可以修改和删除已审核的（自己的）信息<br>" & vbCrLf
    Response.Write "<input name='GroupSetting21' type='checkbox' value='1'>发表信息时可以设置标题前缀<br>" & vbCrLf
    Response.Write "<input name='GroupSetting22' type='checkbox' value='1'>发表信息时可以设置是否显示评论链接<br>" & vbCrLf
    Response.Write "<input name='GroupSetting23' type='checkbox' value='1'>发表信息时可以设置转向链接<br>" & vbCrLf
    Response.Write "<input name='GroupSetting24' type='checkbox' value='1'>发表信息时HTML编辑器为高级模式（默认为简洁模式）<br>" & vbCrLf
    Response.Write "每天最多发布<input name='GroupSetting3' type='text' value='10' size='6' maxlength='6' style='text-align: center;'>条信息（不想限制请设置为<b>0</b>）。<br>"
    Response.Write "发布信息时获取积分为栏目设置的<input name='GroupSetting4' type='text' value='1' size='5' maxlength='5' style='text-align: center;'>倍<br>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>评论权限：</td>" & vbCrLf
    Response.Write "         <td><input name='GroupSetting5' type='checkbox' value='1'>在禁止发表评论的栏目仍然可发表评论<br><input name='GroupSetting6' type='checkbox' value='1'>在评论需要审核的栏目里发表评论不需要审核</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>短消息权限：</td>" & vbCrLf
    Response.Write "      <td> 每次最多可同时向<input name='GroupSetting7' type='text' value='1' size='4' maxlength='4' style='text-align: center;'>人发送短消息（如果为0，则不允许发送短消息）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>收藏夹权限：</td>" & vbCrLf
    Response.Write "      <td>会员收藏夹内最多可收录<input name='GroupSetting8' type='text' value='500' size='5' maxlength='5' style='text-align: center;'>条信息（如果为0，则没有收藏权限）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>上传文件权限：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting9' type='checkbox' value='1' checked>允许在开放上传的频道上传文件<br>最大允许上传<input name='GroupSetting10' type='text' value='1024' size='5' style='text-align: center;'>K的文件（当所设置值大于某一频道的设置时，以频道设置为准。）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>商城权限：</td>" & vbCrLf
    Response.Write "      <td>购物时可以享受的折扣率：<input name='GroupSetting11' type='text' value='100' size='5' maxlength='5' style='text-align: center;'> %<br>"
    Response.Write "        <input name='GroupSetting12' type='checkbox' value='1' checked>是否可以享受折上折优惠（对指定会员价的商品无效）<br>"
    Response.Write "        允许透支的最大额度：<input name='GroupSetting13' type='text' value='0' size='6' maxsize='6' style='text-align: center;'> 元人民币<br>"
    Response.Write "        <input name='GroupSetting30' type='checkbox' value='1'>是否可以批发商品<br>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>计费方式：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting14' type='radio' value='0' checked>只判断" & PointName & "：有" & PointName & "时，即使有效期已经到期，仍可以查看收费内容；" & PointName & "用完后，即使有效期没有到期，也不能查看收费内容。<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting14' type='radio' value='1'>只判断有效期：只要在有效期内，" & PointName & "用完后仍可以查看收费内容；过期后，即使会员有" & PointName & "也不能查看收费内容。<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting14' type='radio' value='2'>同时判断" & PointName & "和有效期：" & PointName & "用完或有效期到期后，就不可查看收费内容。<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting14' type='radio' value='3'>同时判断" & PointName & "和有效期：" & PointName & "用完并且有效期到期后，才不能查看收费内容。" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>扣" & PointName & "方式：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupSetting15' type='radio' value='0' checked>有效期内，查看收费内容不扣" & PointName & "数，也不做记录。<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting15' type='radio' value='1'>有效期内，查看收费内容不扣" & PointName & "数，但做记录。<br>" & vbCrLf
    Response.Write "          <input name='GroupSetting15' type='radio' value='2'>有效期内，查看收费内容也扣" & PointName & "数。<br>" & vbCrLf
    Response.Write "有效期内，总共可以看<input name='GroupSetting16' type='text' value='0' size='10' maxlength='10' style='text-align: center;'> 条收费信息（如果为0，则不限制）<br>" & vbCrLf
    Response.Write "有效期内，每天最多可以看<input name='GroupSetting17' type='text' value='100' size='10' maxlength='10' style='text-align: center;'> 条收费信息（如果为0，则不限制）" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>自助充值：</td>" & vbCrLf
    Response.Write "         <td><input name='GroupSetting18' type='checkbox' value='1' checked>允许自助兑换" & PointName & "<br><input name='GroupSetting19' type='checkbox' value='1' checked>允许自助兑换有效期<br><input name='GroupSetting20' type='checkbox' value='1' checked>允许将" & PointName & "赠送给他人</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='15%' class='tdbg5' align='right'>聚合空间：</td>" & vbCrLf
    Response.Write "         <td><input name='GroupSetting25' type='checkbox'>启用聚合空间<br>" & vbCrLf
    Response.Write "         <input name='GroupSetting26' type='checkbox'>申请时无须管理员审核<br>" & vbCrLf
    Response.Write " 聚合空间容量为:<input name='GroupSetting27' type='text' value='10' size='4' maxlength='10' style='text-align: center;'>M<br>" & vbCrLf
    Response.Write "         <input name='GroupSetting28' type='checkbox'>用户可以自主更换皮肤" & vbCrLf
    Response.Write "    </td></tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td colspan='3'>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='10' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td colspan='2' align='center'>频 道 权 限 详 细 设 置</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
 
    Dim rsChannel
    Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType<>4 And ModuleType<>5 And ModuleType<>7 And ModuleType<>8 AND Disabled=" & PE_False & " ORDER BY OrderID")
    Do While Not rsChannel.EOF
        Response.Write "          <tr valign='top'>" & vbCrLf
        Response.Write "           <td><fieldset>" & vbCrLf
        Response.Write "   <legend>此会员组在【<font color='red'>" & rsChannel("ChannelName") & "</font>】频道的权限：</legend>" & vbCrLf
        Response.Write "    <table width='100%' cellspacing='1'>" & vbCrLf
        Response.Write "        <tr class='tdbg'>" & vbCrLf
        Response.Write "                <td width='50%'><input type='radio' name='" & rsChannel("ChannelDir") & "purview' checked onClick=table" & rsChannel("ChannelID") & ".style.display='none'>无任何权限(开放栏目除外)"
        Response.Write "&nbsp;&nbsp;<input type='radio' name='" & rsChannel("ChannelDir") & "purview' onClick=table" & rsChannel("ChannelID") & ".style.display='block'>设置会员组在该频道的权限</td>" & vbCrLf
        Response.Write "             <td></td>" & vbCrLf
        Response.Write "        <tr class='tdbg' id='table" & rsChannel("ChannelID") & "' style='display:none'>" & vbCrLf
        Response.Write "         <td width='50%'>" & vbCrLf
        Response.Write "         <iframe id='frm" & rsChannel("ChannelDir") & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Group&ChannelID=" & rsChannel("ChannelID") & "'></iframe>" & vbCrLf
        Response.Write "         <input name='arrClass_Browse_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Browse_" & rsChannel("ChannelDir") & "' value=''>" & vbCrLf
        Response.Write "         <input name='arrClass_View_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_View_" & rsChannel("ChannelDir") & "' value=''>" & vbCrLf
        Response.Write "         <input name='arrClass_Input_" & rsChannel("ChannelDir") & "' type='hidden' id='arrClass_Input_" & rsChannel("ChannelDir") & "' value=''></td>" & vbCrLf
        Response.Write "         <td width='50%'><font color='#0000FF'>注：</font><br>1、栏目权限采用继承制度，即在某一栏目拥有某项权限，则在此栏目的所有子栏目中都拥有这项权限，并可在子栏目中指定更多的权限。<br>2、灰色并且选中的项目，说明该栏目为开放栏目，会员组在此栏目拥有浏览和查看权限。<br><br><font color='red'>权限含义：</font><br>浏览－－指可以浏览此栏目的信息列表<br>查看－－指可以查看此栏目中的信息的内容<br>发布－－指可以在此栏目中发布信息</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "   </table>" & vbCrLf
        Response.Write "   </fieldset></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "            <tr>" & vbCrLf
    Response.Write "                <td align='center'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='Action' value='SaveAdd'>" & vbCrLf
    Response.Write "                    <input type='submit' value='添加会员组'>" & vbCrLf
    Response.Write "                    <input type='button' name='cancel' value=' 取 消 ' onClick=""JavaScript:window.location.href='Admin_UserGroup.asp'"">" & vbCrLf
    Response.Write "                </td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
 
End Sub

Sub Modify()
    GroupID = PE_CLng(Trim(Request.QueryString("GroupID")))
    If GroupID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定GroupID</li>"
        Exit Sub
    End If
        
    Set rsUserGroup = Conn.Execute("SELECT GroupID,GroupName,GroupIntro,GroupType,GroupSetting,arrClass_Browse,arrClass_View,arrClass_Input FROM PE_UserGroup WHERE GroupID=" & GroupID & "")
    If rsUserGroup.BOF And rsUserGroup.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的会员组</li>"
        rsUserGroup.Close
        Set rsUserGroup = Nothing
        Exit Sub
    End If
    
    '防止会员手工修改数据库导致数组内容缺少的错误
    GroupSetting = rsUserGroup("GroupSetting") & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
    '完毕
    GroupSetting = Split(GroupSetting, ",")
    Dim i
    For i = 0 To UBound(GroupSetting)
        If GroupSetting(i) = "" Then GroupSetting(i) = 0
    Next
    
    Call ShowJS_Check
    
    Response.Write "<form method='post' action='Admin_UserGroup.asp' name='form1' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='3'><div align='center'>修 改 会 员 组 设 置</div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>会员组名称：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupName' type='text' id='GroupName' value='" & rsUserGroup("GroupName") & "' size='20' maxlength='20'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>会员组说明：</td>" & vbCrLf
    Response.Write "      <td><input name='GroupIntro' type='text' id='GroupIntro' value='" & rsUserGroup("GroupIntro") & "' size='50' maxlength='200'><font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='15%' class='tdbg5' align='right'>组 类 型：</td>" & vbCrLf
    Response.Write "      <td><select name='GroupType' id='GroupType'"
    If rsUserGroup("GroupType") < 3 or rsUserGroup("GroupType") = 5  Then Response.Write "disabled"
    Response.Write ">" & vbCrLf
    If rsUserGroup("GroupType") = 0 Then
        Response.Write "        <option value='0' selected>等待邮件验证</option>" & vbCrLf
    End If
    If rsUserGroup("GroupType") = 1 Then
        Response.Write "        <option value='1' selected>等待管理员审批</option>" & vbCrLf
    End If
    If rsUserGroup("GroupType") = 2 Then
        Response.Write "        <option value='2' selected>默认会员组</option>" & vbCrLf
    End If
    If rsUserGroup("GroupType") = 5 Then
        Response.Write "        <option value='5' selected>匿名投稿</option>" & vbCrLf
    End If    
    Response.Write "            <option value='3'"
    If rsUserGroup("GroupType") = 3 Then
        Response.Write " selected"
    End If
    Response.Write ">注册会员</option>" & vbCrLf
    Response.Write "            <option value='4'"
    If rsUserGroup("GroupType") = 4 Then
        Response.Write " selected"
    End If
    Response.Write ">代 理 商</option></select>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    If rsUserGroup("GroupID")<>-1 Then            
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>发布权限：</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting1' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(1)), 1) & ">在发布信息需要审核的频道，此组会员发布信息不需要审核<br>" & vbCrLf
        Response.Write "<input name='GroupSetting2' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(2)), 1) & ">可以修改和删除已审核的（自己的）信息<br>" & vbCrLf
        Response.Write "<input name='GroupSetting21' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(21)), 1) & ">发表信息时可以设置标题前缀<br>" & vbCrLf
        Response.Write "<input name='GroupSetting22' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(22)), 1) & ">发表信息时可以设置是否显示评论链接<br>" & vbCrLf
        Response.Write "<input name='GroupSetting23' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(23)), 1) & ">发表信息时可以设置转向链接<br>" & vbCrLf
        Response.Write "<input name='GroupSetting24' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(24)), 1) & ">发表信息时HTML编辑器为高级模式（默认为简洁模式）<br>" & vbCrLf
        Response.Write "每天最多发布<input name='GroupSetting3' type='text' value='" & GroupSetting(3) & "' size='6' maxlength='6' style='text-align: center;'>条信息（不想限制请设置为<b>0</b>）。<br>"
        Response.Write "发布信息时获取积分为栏目设置的<input name='GroupSetting4' type='text' value='" & GroupSetting(4) & "' size='5' maxlength='5' style='text-align: center;'>倍<br>"
        
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "    <td width='15%' class='tdbg5' align='right'>评论权限：</td>" & vbCrLf
        Response.Write "         <td><input name='GroupSetting5' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(5)), 1) & ">在禁止发表评论的栏目里仍然可发表评论<br><input name='GroupSetting6' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(6)), 1) & ">在评论需要审核的栏目里发表评论不需要审核</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>短消息权限：</td>" & vbCrLf
        Response.Write "      <td>每次最多可同时向<input name='GroupSetting7' type='text' value='" & GroupSetting(7) & "' size='4' maxlength='4' style='text-align: center;'>人发送短消息（如果为0，则不允许发送短消息）</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>收藏夹权限：</td>" & vbCrLf
        Response.Write "      <td>会员收藏夹内最多可收录<input name='GroupSetting8' type='text' value='" & GroupSetting(8) & "' size='5' maxlength='5' style='text-align: center;'>条信息（如果为0，则没有收藏权限）</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>上传文件权限：</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting9' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(9)), 1) & ">允许在开放上传的频道上传文件<br>最大允许上传<input name='GroupSetting10' type='text' value='" & GroupSetting(10) & "' size='5' style='text-align: center;'>K的文件（当所设置值大于某一频道的设置时，以频道设置为准。）</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>商城权限：</td>" & vbCrLf
        Response.Write "      <td>购物时可以享受的折扣率：<input name='GroupSetting11' type='text' value='" & GroupSetting(11) & "' size='5' maxlength='5' style='text-align: center;'>%<br>"
        Response.Write "<input name='GroupSetting12' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(12)), 1) & ">是否可以享受折上折优惠（对指定会员价的商品无效）<br> 允许透支的最大额度：<input name='GroupSetting13' type='text' value='" & GroupSetting(13) & "' size='6' maxsize='6' style='text-align: center;'>元人民币" & vbCrLf
        Response.Write "        <br><input name='GroupSetting30' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(30)), 1) & ">是否可以批发商品<br>"
        Response.Write "    </td></tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>计费方式：</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting14' type='radio' " & RadioValue(PE_CLng(GroupSetting(14)), 0) & ">只判断" & PointName & "：有" & PointName & "时，即使有效期已经到期，仍可以查看收费内容；" & PointName & "用完后，即使有效期没有到期，也不能查看收费内容。<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting14' " & RadioValue(PE_CLng(GroupSetting(14)), 1) & ">只判断有效期：只要在有效期内，" & PointName & "用完后仍可以查看收费内容；过期后，即使会员有" & PointName & "也不能查看收费内容。<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting14' " & RadioValue(PE_CLng(GroupSetting(14)), 2) & ">同时判断" & PointName & "和有效期：" & PointName & "用完或有效期到期后，就不可查看收费内容。<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting14' " & RadioValue(PE_CLng(GroupSetting(14)), 3) & ">同时判断" & PointName & "和有效期：" & PointName & "用完并且有效期到期后，才不能查看收费内容。" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>扣" & PointName & "方式：</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting15' type='radio' " & RadioValue(PE_CLng(GroupSetting(15)), 0) & ">有效期内，查看收费内容不扣" & PointName & "，也不做记录。<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting15' " & RadioValue(PE_CLng(GroupSetting(15)), 1) & ">有效期内，查看收费内容不扣" & PointName & "，但做记录。<br>" & vbCrLf
        Response.Write "          <input type='radio' name='GroupSetting15' " & RadioValue(PE_CLng(GroupSetting(15)), 2) & ">有效期内，查看收费内容也扣" & PointName & "。<br>" & vbCrLf
        Response.Write "有效期内，总共可以看<input name='GroupSetting16' type='text' value='" & GroupSetting(16) & "' size='10' maxlength='10' style='text-align: center;'> 条收费信息（如果为0，则不限制）<br>" & vbCrLf
        Response.Write "有效期内，每天最多可以看<input name='GroupSetting17' type='text' value='" & GroupSetting(17) & "' size='10' maxlength='10' style='text-align: center;'> 条收费信息（如果为0，则不限制）" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "    <td width='15%' class='tdbg5' align='right'>自助充值：</td>" & vbCrLf
        Response.Write "         <td><input name='GroupSetting18' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(18)), 1) & ">允许自助兑换" & PointName & "<br><input name='GroupSetting19' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(19)), 1) & ">允许自助兑换有效期<br><input name='GroupSetting20' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(20)), 1) & ">允许将" & PointName & "赠送给他人</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "    <td width='15%' class='tdbg5' align='right'>聚合空间：</td>" & vbCrLf
        Response.Write "         <td><input name='GroupSetting25' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(25)), 1) & ">启用聚合空间<br>" & vbCrLf
        Response.Write "         <input name='GroupSetting26' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(26)), 1) & ">聚合空间无须审核<br>" & vbCrLf
        Response.Write " 聚合空间最大容量为:<input name='GroupSetting27' type='text' value='" & GroupSetting(27) & "' size='4' maxlength='10' style='text-align: center;'>M<br>" & vbCrLf
        Response.Write "         <input name='GroupSetting28' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(28)), 1) & ">用户可以自主更换皮肤" & vbCrLf
        Response.Write "    </td></tr>" & vbCrLf
    Else    	
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>发布权限：</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting1' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(1)), 1) & ">在发布信息需要审核的频道，此组会员发布信息不需要审核<br>" & vbCrLf
        Response.Write "<input name='GroupSetting21' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(21)), 1) & ">发表信息时可以设置标题前缀<br>" & vbCrLf
        Response.Write "<input name='GroupSetting22' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(22)), 1) & ">发表信息时可以设置是否显示评论链接<br>" & vbCrLf
        Response.Write "<input name='GroupSetting23' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(23)), 1) & ">发表信息时可以设置转向链接<br>" & vbCrLf
        Response.Write "<input name='GroupSetting24' type='checkbox' " & RadioValue(PE_CLng(GroupSetting(24)), 1) & ">发表信息时HTML编辑器为高级模式（默认为简洁模式）<br>" & vbCrLf
        Response.Write "每天最多发布<input name='GroupSetting3' type='text' value='" & GroupSetting(3) & "' size='6' maxlength='6' style='text-align: center;'>条信息（不想限制请设置为<b>0</b>）。<br>"
        
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='15%' class='tdbg5' align='right'>上传文件权限：</td>" & vbCrLf
        Response.Write "      <td><input name='GroupSetting9' type='checkbox'" & RadioValue(PE_CLng(GroupSetting(9)), 1) & ">允许在开放上传的频道上传文件<br>最大允许上传<input name='GroupSetting10' type='text' value='" & GroupSetting(10) & "' size='5' style='text-align: center;'>K的文件（当所设置值大于某一频道的设置时，以频道设置为准。）</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf    
    End If        
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td colspan='3'>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='10' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td colspan='2' align='center'>频 道 权 限 详 细 设 置</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
     
    Dim rsChannel, arrPurviews, IsNoPurview
    arrPurviews = rsUserGroup("arrClass_Browse") & "," & rsUserGroup("arrClass_View") & "," & rsUserGroup("arrClass_Input")
    If rsUserGroup("GroupID")<>-1 Then            
        Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType<>4 And ModuleType<>5 and ModuleType<>7 and ModuleType<>8 AND Disabled=" & PE_False & " ORDER BY OrderID")
    Else
        Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType=1 AND Disabled=" & PE_False & " ORDER BY OrderID")    
    End If
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
        Response.Write "         <iframe id='frm" & rsChannel("ChannelDir") & "' height='200' width='100%' src='Admin_SetClassPurview.asp?ManageType=Group&Action=Modify&ChannelID=" & rsChannel("ChannelID") & "&GroupID=" & GroupID & "'></iframe>" & vbCrLf
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
    Response.Write "            <tr>" & vbCrLf
    Response.Write "                <td align='center'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='GroupID' value='" & rsUserGroup("GroupID") & "'>" & vbCrLf
    Response.Write "                    <input type='hidden' name='Action' value='SaveModify'>" & vbCrLf
    Response.Write "                    <input type='submit' value='保存修改结果'>" & vbCrLf
    Response.Write "                    <input type='button' name='cancel' value=' 取 消 ' onClick=""JavaScript:window.location.href='Admin_UserGroup.asp'"">" & vbCrLf
    Response.Write "                </td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Del()
    GroupID = PE_CLng(Trim(Request("GroupID")))
    If GroupID = 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不能删除系统默认的会员组</li>"
        Exit Sub
    End If
    Conn.Execute ("update PE_User set GroupID=1 where GroupID=" & GroupID & "")
    Conn.Execute ("delete from PE_UserGroup where GroupID=" & GroupID & " AND GroupType>=2")
    Call main
End Sub

Sub SaveGroup()
    Dim GroupType, strValue, GroupIntro, i
    Dim rsUserGroup, rsChannel, GroupPurview, GroupPurviewChannel
    Dim arrClass_Browse, arrClass_View, arrClass_Input
    GroupID = Trim(Request.Form("GroupID"))
    GroupName = Trim(Request.Form("GroupName"))
    GroupIntro = Trim(Request.Form("GroupIntro"))
    GroupType = Trim(Request.Form("GroupType"))
    FoundErr = False
    If GroupName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>会员组名称不能为空</li>"
        Exit Sub
    Else
        GroupName = ReplaceBadChar(GroupName)
    End If
    If GroupType = "" Then
        GroupType = 0
    Else
        GroupType = CLng(GroupType)
    End If
    GroupSetting = ""
    For i = 0 To 30
        strValue = Trim(Request.Form("GroupSetting" & i & ""))
        If strValue = "" Or (Not IsNumeric(strValue)) Then
            strValue = "0"
        End If
        If GroupSetting = "" Then
            GroupSetting = strValue
        Else
            GroupSetting = GroupSetting & "," & strValue
        End If
    Next

    arrClass_Browse = ""
    arrClass_View = ""
    arrClass_Input = ""
    
    Dim tBrowse, tView, tInput, ChannelDir
    If PE_Clng(GroupID)<>-1 then
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType<>4 And ModuleType<>5 And Disabled=" & PE_False & " ORDER BY OrderID")
    Else
        Set rsChannel = Conn.Execute("SELECT ChannelDir FROM PE_Channel WHERE ChannelType<=1 And ModuleType=1 And Disabled=" & PE_False & " ORDER BY OrderID")
    End If    
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

    Set rsUserGroup = Server.CreateObject("Adodb.Recordset")
    If Action = "SaveAdd" Then
        rsUserGroup.Open "SELECT * FROM PE_UserGroup WHERE GroupName='" & GroupName & "'", Conn, 1, 3
        If Not (rsUserGroup.BOF And rsUserGroup.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库中已有同名会员组！</li>"
        Else
            rsUserGroup.addnew
            rsUserGroup("GroupID") = PE_CLng(Conn.Execute("select max(GroupID) from PE_UserGroup")(0)) + 1
        End If
    Else
        rsUserGroup.Open "SELECT * FROM PE_UserGroup WHERE GroupID=" & GroupID, Conn, 1, 3
        If rsUserGroup.BOF And rsUserGroup.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>数据库没有发现此会员组！</li>"
        End If
    End If
    If FoundErr = True Then
        rsUserGroup.Close
        Set rsUserGroup = Nothing
        Exit Sub
    End If

    rsUserGroup("GroupName") = GroupName
    rsUserGroup("GroupIntro") = GroupIntro
    If GroupType > 0 Then
        rsUserGroup("GroupType") = GroupType
    End If
    rsUserGroup("GroupSetting") = GroupSetting
    rsUserGroup("arrClass_Browse") = arrClass_Browse
    rsUserGroup("arrClass_View") = arrClass_View
    rsUserGroup("arrClass_Input") = arrClass_Input
    rsUserGroup.Update
    rsUserGroup.Close
    Set rsUserGroup = Nothing
    Call main

End Sub

Function GetGroupNum(iGroupID)
    If Not IsNumeric(iGroupID) Then Exit Function
    Dim rsUserGroup
    Set rsUserGroup = Conn.Execute("SELECT Count(UserID) FROM PE_User WHERE GroupID=" & iGroupID & "")
    If IsNull(rsUserGroup(0)) Then
        GetGroupNum = 0
    Else
        GetGroupNum = rsUserGroup(0)
    End If
    Set rsUserGroup = Nothing
End Function
%>

