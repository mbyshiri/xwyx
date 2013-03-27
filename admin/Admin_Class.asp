<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.Class.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 1   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim arrInvalidDir
Dim pNum, pNum2, OpenType_Class, iOrderID
Dim ClassLink
Dim HtmlDir



HtmlDir = InstallDir & ChannelDir
ParentID = PE_CLng(Trim(Request("ParentID")))
arrInvalidDir = "HTML,JS,Special,List,Images,UploadFiles,UploadSoft,UploadSoftPic,UploadThumbs,UploadPhotos,UploadFlash,UploadVideo,UploadMusic"

Response.Write "<html><head><title>" & ChannelName & "管理----栏目管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle(ChannelName & "管理----栏目管理", 10003)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>"
Response.Write "    <td height='30'>"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "栏目管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Add'>添加" & ChannelShortName & "栏目</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Order'>一级栏目排序</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=OrderN'>N级栏目排序</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Reset'>复位所有" & ChannelShortName & "栏目</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Unite'>" & ChannelShortName & "栏目合并</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Batch'>批量设置</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Patch'>修复栏目结构</a>"
Response.Write "    </td></tr></table>" & vbCrLf

Select Case Action
Case "Add"
    Call AddClass
Case "SaveAdd"
    Call SaveAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Move"
    Call MoveClass
Case "SaveMove"
    Call SaveMove
Case "Del"
    Call DeleteClass
Case "Clear"
    Call ClearClass
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "Order"
    Call order
Case "UpOrderN"
    Call UpOrderN
Case "DownOrderN"
    Call DownOrderN
Case "OrderN"
    Call OrderN
Case "Reset"
    Call Reset
Case "SaveReset"
    Call SaveReset
Case "Unite"
    Call Unite
Case "SaveUnite"
    Call SaveUnite
Case "Batch"
    Call ShowBatch
Case "DoBatch"
    Call DoBatch
Case "Patch"
    Call Patch
Case "DoPatch"
    Call DoPatch
Case "ResetChildClass"
    Call ResetChildClass
Case "CreateJS"
    Call CreateJS_Class
    Call WriteSuccessMsg("已经成功生成栏目JS文件。", ComeUrl)
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "栏目管理操作失败，失败原因：" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim arrShowLine(20), i
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    Dim sqlClass, rsClass, iDepth, ClassDir, ClassItemDir
    sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'> "
    Response.Write "    <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "    <td align='center'><strong>栏目名称及目录</strong></td>"
    Response.Write "    <td width='60' align='center'><strong>栏目权限</strong></td>"
    Response.Write "    <td width='100' align='center'><strong>栏目属性</strong></td>"
    Response.Write "    <td width='380' align='center'><strong>操作选项</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    If rsClass.BOF And rsClass.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>没有任何栏目</td></tr>"
    Else
        Do While Not rsClass.EOF
            Response.Write "<tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='30' align='center'>" & rsClass("ClassID") & "</td>"
            Response.Write "    <td>"
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
                Response.Write " [" & rsClass("ClassDir") & "]"
            End If

            'Response.Write "&nbsp;&nbsp;" & rsClass("ClassID") & "," & rsClass("PrevID") & "," & rsClass("NextID") & "," & rsClass("ParentID") & "," & rsClass("RootID")
            Response.Write "</td><td align='center' width='60'>"
            Select Case rsClass("ClassPurview")
            Case 0
                Response.Write "<font color='green'>开放</font>"
            Case 1
                Response.Write "<font color='blue'>半开放</font>"
            Case 2
                Response.Write "<font color='red'>认证</font>"
            End Select
            Response.Write "</td><td align='left' width='100'>"
            If rsClass("OpenType") = 0 Then
                Response.Write "&nbsp;原 "
            Else
                Response.Write "&nbsp;新 "
            End If
            If rsClass("ClassType") = 1 Then
                If rsClass("ParentID") = 0 And rsClass("ShowOnIndex") = True Then
                    Response.Write "首 "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsClass("ParentID") > 0 And rsClass("IsElite") = True Then
                    Response.Write "列 "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsClass("Child") > 0 Then
                    If rsClass("EnableAdd") = True Then
                        Response.Write "开 "
                    Else
                        Response.Write "锁 "
                    End If
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsClass("EnableProtect") = True Then
                    Response.Write "保"
                Else
                    Response.Write "&nbsp;&nbsp;"
                End If
            End If
            Response.Write "</td><td align='left' width='380'>&nbsp;"
            If ModuleType = 6 And iDepth > 3 Then
                Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;"
            Else
                If rsClass("ClassType") = 1 Then
                    Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Add&ParentID=" & rsClass("ClassID") & "'>添加子栏目</a>&nbsp;|&nbsp;"
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;"
                End If
            End If
            Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Modify&ClassID=" & rsClass("ClassID") & "'>修改设置</a>&nbsp;|&nbsp;"
            Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Move&ClassID=" & rsClass("ClassID") & "'>移动栏目</a>&nbsp;|&nbsp;"
            If rsClass("ParentID") = 0 And rsClass("ClassType") = 1 Then
                Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=ResetChildClass&ClassID=" & rsClass("ClassID") & "' onclick=""return confirm('“复位子栏目”将把此栏目的所有子栏目都复位成二级子栏目！请慎重操作！确定要复位子栏目吗？')"">复位子栏目</a>&nbsp;|&nbsp;"
            Else
                Response.Write "<font color='gray'>复位子栏目</font>&nbsp;|&nbsp;"
            End If
            Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Clear&ClassID=" & rsClass("ClassID") & "' onClick='return ConfirmDel3();'>清空</a>&nbsp;|&nbsp;"
            Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Del&ClassID=" & rsClass("ClassID") & "' onClick='return ConfirmDel2();'>删除</a>"
            Response.Write "</td></tr>" & vbCrLf
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    Response.Write "</table>" & vbCrLf
    
    Response.Write ""
    Response.Write "<table width='100%'><tr><form name='form1' action='Admin_Class.asp' method='post'><td align='center'>"
    Response.Write "<input name='Action' type='hidden' id='Action' value='CreateJS'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='刷新栏目JS' style='cursor:hand;'"
    If ObjInstalled_FSO = False Then
        Response.Write " disabled"
    End If
    Response.Write "></td></form></tr></table>"
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function ConfirmDel1(){" & vbCrLf
    Response.Write "   alert('此栏目下还有子栏目，必须先删除下属子栏目后才能删除此栏目！');" & vbCrLf
    Response.Write "   return false;}" & vbCrLf
    Response.Write "function ConfirmDel2(){" & vbCrLf
    Response.Write "   if(confirm('删除栏目操作将删除此栏目中的所有子栏目和" & ChannelShortName & "，并且不能恢复！确定要删除此栏目吗？'))" & vbCrLf
    Response.Write "     return true;" & vbCrLf
    Response.Write "   else" & vbCrLf
    Response.Write "     return false;}" & vbCrLf
    Response.Write "function ConfirmDel3(){" & vbCrLf
    Response.Write "   if(confirm('清空栏目将把栏目（包括子栏目）的所有" & ChannelShortName & "放入回收站中！确定要清空此栏目吗？'))" & vbCrLf
    Response.Write "     return true;" & vbCrLf
    Response.Write "   else" & vbCrLf
    Response.Write "     return false;}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br><table width='100%'><tr><td colspan='5'><b>栏目属性中各项的含义：</b></td></tr>"
    Response.Write "<tr><td>原----在原窗口打开</td><td>新----在新窗口打开</td><td>首----在首页分类列表处显示，只对一级栏目有效</td></tr>"
    Response.Write "<tr><td>列----在父栏目分类列表处显示</td><td>锁----有子栏目时不允许添加" & ChannelShortName & "</td><td>开----有子栏目时可以添加" & ChannelShortName & "</td></tr>"
    Response.Write "<tr><td>保----启用防复制/下载功能</td></tr></table>"
End Sub


Sub AddClass()
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Class.asp?ChannelID=" & ChannelID & "'>栏目管理</a>&nbsp;&gt;&gt;&nbsp;添加栏目</td></tr></table>"
    Response.Write "<form name='form1' method='post' action='Admin_Class.asp' onsubmit='return check()'>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>权限设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'"
    If ModuleType = 5 Then Response.Write " style='display:none'"
    Response.Write ">收费设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>栏目选项</td>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>自设内容</td>" & vbCrLf
    End If
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>所属栏目：</strong>"
    If ModuleType = 6 Then
        Response.Write "<br>&nbsp;&nbsp;<font color=red>在供求频道下,子栏目的层数不可以超过4层</font>"
    End If
    Response.Write "       </td>"
    Response.Write "      <td><select name='ParentID'><option value='0'>无（做为一级栏目）</option>" & GetClass_Option(ChannelID, ParentID) & "</select> <font color=blue>不能指定为外部栏目</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目名称：</strong></td>"
    Response.Write "      <td><input name='ClassName' type='text' size='20' maxlength='50'> <font color=red>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目类型：</strong><br><font color=red>请慎重选择，栏目一旦添加后就不能再更改栏目类型。</font></td>" & vbCrLf
    Response.Write "      <td>"
    If ModuleType = 5 Then
        Response.Write "        <input name='ClassType' type='radio' value='1' checked onclick=""HideTabTitle('',0)"">"
    Else
        Response.Write "        <input name='ClassType' type='radio' value='1' checked onclick=""HideTabTitle('',1)"">"
    End If
    Response.Write "        <font color=blue><b>内部栏目</b></font>&nbsp;&nbsp;内部栏目具有详细的参数设置。可以添加子栏目和文章。<br>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;内部栏目的目录名：<input name='ClassDir' type='text' size='20' maxlength='20'> <font color='#FF0000'>注意，目录名只能是英文</font><br><br>"
    Response.Write "        <input name='ClassType' type='radio' value='2' onclick=""HideTabTitle('none')"">"
    Response.Write "        <font color=blue><b>外部栏目</b></font>&nbsp;&nbsp;外部栏目指链接到本系统以外的地址中。当此栏目准备链接到网站中的其他系统时，请使用这种方式。不能在外部栏目中添加文章，也不能添加子栏目。<br>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;外部栏目的链接地址：<input name='LinkUrl' type='text' id='LinkUrl' value='" & SiteUrl & "' size='40' maxlength='200'>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目图片地址：</strong><br>用于在栏目页显示指定的图片</td>"
    Response.Write "      <td><input name='ClassPicUrl' type='text' id='ClassPicUrl' size='60' maxlength='255'></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目提示：</strong><br>鼠标移至栏目名称上时将显示设定的提示文字（不支持HTML）</td>"
    Response.Write "      <td><textarea name='Tips' cols='60' rows='2' id='Tips'></textarea></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目说明：</strong><br>用于在栏目页详细介绍栏目信息，支持HTML</td>"
    Response.Write "      <td><textarea name='Readme' cols='60' rows='3' id='Readme'></textarea></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>栏目META关键词：</strong><br>针对搜索引擎设置的关键词<br>多个关键词请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>栏目META网页描述：</strong><br>针对搜索引擎设置的网页描述<br>多个描述请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    '刘永涛 屏蔽了供求信息中不需要的选项
    If ModuleType = 6 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>发表权限：</strong><br>请在此设置允许在此栏目发表信息的会员组。<br>游客没有发表信息权限。</td>"
        Response.Write "      <td>" & GetUserGroup("Input", 0, 5) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>评论权限：</strong></td>"
        Response.Write "      <td>"
        Response.Write "        <input name='EnableComment' type='checkbox' value='True' checked>允许在此栏目发表评论<br>"
        Response.Write "        <input name='CheckComment' type='checkbox' value='True' checked>评论需要审核"
        Response.Write "      </td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
    Response.Write "  <tbody id='Tabs' style='display:none'>"
    If ModuleType <> 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>浏览/查看权限：</strong><br><font color='red'>栏目权限为继承关系。例如：当父栏目设为“认证栏目”时，子栏目的权限设置将继承父栏目设置，即使子栏目设为“开放栏目”也无效。相反，如果父栏目设为“开放栏目”，子栏目可以设为“半开放栏目”或“认证栏目”。</font></td>"
        Response.Write "      <td>"
        Response.Write "        <table>"
        Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' value='0' checked>开放栏目</td><td>任何人（包括游客）可以浏览和查看此栏目下的信息。</td></tr>"
        Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' value='1'>半开放栏目</td><td>任何人（包括游客）都可以浏览。游客不可查看，其他会员根据会员组的栏目权限设置决定是否可以查看。</td></tr>"
        Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' value='2'>认证栏目</td><td>游客不能浏览和查看，其他会员根据会员组的栏目权限设置决定是否可以浏览和查看。</td></tr>"
        Response.Write "        </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>允许浏览此栏目的会员组：</strong><br>如果栏目是“半开放栏目”或“认证栏目”，请在此设置允许浏览此栏目的会员组</td>"
        Response.Write "      <td>" & GetUserGroup("Browse", 0, 5) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>允许查看此栏目下信息的会员组：</strong><br>如果栏目是“认证栏目”，请在此设置允许查看此栏目下信息的会员组<br>如果在信息中设置了查看权限，则以信息中的权限设置优先</td>"
        Response.Write "      <td>" & GetUserGroup("View", 0, 5) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>发表权限：</strong><br>请在此设置允许在此栏目发表信息的会员组。<br>游客没有发表信息权限。</td>"
        Response.Write "      <td>" & GetUserGroup("Input", 0, 5) & "</td>"
        Response.Write "    </tr>" & vbCrLf
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>评论权限：</strong></td>"
    Response.Write "      <td>"
    Response.Write "        <input name='EnableComment' type='checkbox' value='True' checked>允许在此栏目发表评论<br>"
    Response.Write "        <input name='CheckComment' type='checkbox' value='True' checked>评论需要审核"
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    End If
    '2005-12-23
    '刘永涛 屏蔽掉一些不必要的信息
    If ModuleType = 6 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>发布信息扣除的点数：</strong><br>会员在此栏目下发布信息时，该扣除其会员点数</td>"
        Response.Write "      <td><input name='ReleaseClassPoint' type='text' value='0' size='4' maxlength='4' style='text-align:center'>&nbsp;点数</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>设置栏目推荐每天扣除的点数：</strong></td>"
        Response.Write "      <td><input name='CommandClassPoint' type='text' value='0' size='4' maxlength='4' style='text-align:center'>&nbsp;点/天</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>积分奖励：</strong><br>会员在此栏目发表信息时可以得到的积分</td>"
    Response.Write "      <td>会员在此栏目每发表一条信息，可以得到 <input name='PresentExp' type='text' value='1' size='4' maxlength='4' style='text-align:center'> 分积分</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>默认收费" & PointName & "数：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，该" & ChannelShortName & "默认的收费" & PointName & "数</td>"
    Response.Write "      <td><input name='DefaultItemPoint' type='text' value='0' size='4' maxlength='4' style='text-align:center'> " & PointUnit & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>默认重复收费：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，该" & ChannelShortName & "默认的重复收费方式</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' value='0' checked>不重复收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' value='1'>距离上次收费时间 <input name='DefaultItemPitchTime' type='text' value='24' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' value='2'>会员重复查看此文章 <input name='DefaultItemReadTimes' type='text' value='10' size='8' maxlength='8' style='text-align:center'> 次后重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' value='3'>上述两者都满足时重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' value='4'>上述两者任一个满足时就重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' value='5'>每阅读一次就重复收费一次（建议不要使用）"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>默认分成比例：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，该" & ChannelShortName & "默认的分成比例</td>"
    Response.Write "      <td><input name='DefaultItemDividePercent' type='text' value='0' size='4' maxlength='4' style='text-align:center'> %</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    End If
    '2005-12-23

    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>打开方式：</strong></td>"
    Response.Write "      <td><input name='OpenType' type='radio' value='0' checked>在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1'>在新窗口打开</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在顶部导航栏显示：</strong><br>此选项只对一级栏目有效。</td>"
    Response.Write "      <td><input name='ShowOnTop' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='ShowOnTop' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在频道首页分类列表处显示：</strong><br>此选项只对一级栏目有效。如果一级栏目比较多，但首页不想显示太多的分类列表，这个选项就非常有用。</td>"
    Response.Write "      <td><input name='ShowOnIndex' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='ShowOnIndex' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在父栏目的分类列表处显示：</strong><br>如果某栏目下有几十个子栏目，但只想显示其中几个子栏目的" & ChannelShortName & "列表，这个选项就非常有用。</td>"
    Response.Write "      <td><input name='IsElite' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>有子栏目时是否可以在此栏目添加" & ChannelShortName & "：</strong></td>"
    Response.Write "      <td><input name='EnableAdd' type='radio' value='True'>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='EnableAdd' type='radio' value='False' checked>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否启用此栏目的防止复制、防盗链功能：</strong></td>"
    Response.Write "      <td><input name='EnableProtect' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='EnableProtect' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='SkinID' id='SkinID'>" & GetSkin_Option(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目模板：</strong><br>相关模板中包含了栏目设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色风格”失效。 </td>"
    Response.Write "      <td><select name='TemplateID' id='TemplateID'>" & GetTemplate_Option(ChannelID, 2, 0) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>每页显示的" & ChannelShortName & "数：</strong><br>当此栏目为最下一级栏目时，则会分页显示此栏目中的" & ChannelShortName & "，这里指定的是每页显示的" & ChannelShortName & "数。</td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, 20) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>此栏目下的" & ChannelShortName & "的默认配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='DefaultItemSkin' id='DefaultItemSkin'>" & GetSkin_Option(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>此栏目下的" & ChannelShortName & "的默认模板：</strong></td>"
    Response.Write "      <td><select name='DefaultItemTemplate' id='DefaultItemTemplate'>" & GetTemplate_Option(ChannelID, 3, 1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><b>此栏目下的" & ChannelShortName & "列表的排序方式：</b></td>"
    Response.Write "      <td><select name='ItemListOrderType'>" & GetOrderType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><b>此栏目下的" & ChannelShortName & "打开方式：</b></td>"
    Response.Write "      <td><select name='ItemOpenType'>" & GetOpenType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Call EditCustom_Content("Add", "", "Class")
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='2' align='center'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "      <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      <input name='Add' type='submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Call WriteJS
End Sub

Sub WriteJS()
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function check(){" & vbCrLf
    Response.Write "  if (document.form1.ClassName.value==''){" & vbCrLf
    Response.Write "   ShowTabs(0);" & vbCrLf
    Response.Write "   alert('栏目名称不能为空！');" & vbCrLf
    Response.Write "   document.form1.ClassName.focus();" & vbCrLf
    Response.Write "   return false;}" & vbCrLf
    Response.Write "  if(document.form1.ClassType[0].checked==true){" & vbCrLf
    Response.Write "    if(document.form1.ClassDir.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('栏目目录不能为空！');" & vbCrLf
    Response.Write "      document.form1.ClassDir.focus();" & vbCrLf
    Response.Write "      return false;}" & vbCrLf
    Response.Write "  }else{" & vbCrLf
    Response.Write "    if(document.form1.LinkUrl.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('栏目链接地址不能为空！');" & vbCrLf
    Response.Write "      document.form1.LinkUrl.focus();" & vbCrLf
    Response.Write "      return false;}" & vbCrLf
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

End Sub

Sub Modify()
    Dim ClassID, sql, rsClass, i
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    sql = "select * from PE_Class where ClassID=" & ClassID
    Set rsClass = Server.CreateObject("Adodb.recordset")
    rsClass.Open sql, Conn, 1, 1
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Class.asp?ChannelID=" & ChannelID & "'>栏目管理</a>&nbsp;&gt;&gt;&nbsp;修改栏目设置：<font color='red'>" & rsClass("ClassName") & "</td></tr></table>"
    Response.Write "<form name='form1' method='post' action='Admin_Class.asp' onsubmit='return check()'>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    If rsClass("ClassType") = 1 Then
        Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>权限设置</td>" & vbCrLf
        Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'"
        If ModuleType = 5 Then Response.Write " style='display:none'"
        Response.Write ">收费设置</td>" & vbCrLf
        Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>栏目选项</td>" & vbCrLf
        If IsCustom_Content = True And ModuleType <> 6 Then
            Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>自设内容</td>" & vbCrLf
        End If
    End If

    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>所属栏目：</strong><br>如果你想改变所属栏目，请<a href='Admin_Class.asp?Action=Move&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>点此移动栏目</a></td>"
    Response.Write "      <td>" & GetPath(rsClass("ParentID"), rsClass("ParentPath")) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目名称：</strong></td>"
    Response.Write "      <td><input name='ClassName' type='text' value='" & rsClass("ClassName") & "' size='20' maxlength='20'> <font color=red>*</font></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目类型：</strong><br><font color=red>请慎重选择，栏目一旦添加后就不能再更改栏目类型。</font></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "        <input name='ClassType' type='radio' value='1'"
    If rsClass("ClassType") = 1 Then
        Response.Write " checked"
    Else
        Response.Write " disabled"
    End If
    Response.Write ">"
    Response.Write "        <font color=blue><b>内部栏目</b></font>&nbsp;&nbsp;内部栏目具有详细的参数设置。可以添加子栏目和文章。<br>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;内部栏目的目录名：<input name='ClassDir' type='text' size='20' maxlength='20' value='" & rsClass("ClassDir") & "' disabled> <br><br>"
    Response.Write "        <input name='ClassType' type='radio' value='2'"
    If rsClass("ClassType") = 2 Then
        Response.Write " checked"
    Else
        Response.Write " disabled"
    End If
    Response.Write ">"
    Response.Write "        <font color=blue><b>外部栏目</b></font>&nbsp;&nbsp;外部栏目指链接到本系统以外的地址中。当此栏目准备链接到网站中的其他系统时，请使用这种方式。不能在外部栏目中添加文章，也不能添加子栏目。<br>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;外部栏目的链接地址：<input name='LinkUrl' type='text' id='LinkUrl' value='" & rsClass("LinkUrl") & "' size='40' maxlength='200'"
    If rsClass("ClassType") = 1 Then Response.Write " disabled"
    Response.Write ">"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目图片地址：</strong><br>用于在栏目页显示指定的图片</td>"
    Response.Write "      <td><input name='ClassPicUrl' type='text' id='ClassPicUrl' value='" & rsClass("ClassPicUrl") & "' size='60' maxlength='255'></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目提示：</strong><br>鼠标移至栏目名称上时将显示设定的提示文字（不支持HTML）</td>"
    Response.Write "      <td><textarea name='Tips' cols='60' rows='2' id='Tips'>" & rsClass("Tips") & "</textarea></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目说明：</strong><br>用于在栏目页详细介绍栏目信息，支持HTML</td>"
    Response.Write "      <td><textarea name='Readme' cols='60' rows='3' id='Readme'>" & rsClass("ReadMe") & "</textarea></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>栏目META关键词：</strong><br>针对搜索引擎设置的关键词<br>多个关键词请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsClass("Meta_Keywords") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>栏目META网页描述：</strong><br>针对搜索引擎设置的网页描述<br>多个描述请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsClass("Meta_Description") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf


    '刘永涛 屏蔽了供求信息中不需要的选项
    If ModuleType = 6 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>发表权限：</strong><br>请在此设置允许在此栏目发表信息的会员组。<br>游客没有发表信息权限。</td>"
        Response.Write "      <td>" & GetUserGroup("Input", ClassID, 3) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>评论权限：</strong></td>"
        Response.Write "      <td>"
        Response.Write "        <input name='EnableComment' type='checkbox' " & RadioValue(rsClass("EnableComment"), True) & "> 允许在此栏目发表评论<br>"
        Response.Write "        <input name='CheckComment' type='checkbox' " & RadioValue(rsClass("CheckComment"), True) & ">评论需要审核"
        Response.Write "      </td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        If ModuleType <> 5 Then
            Response.Write "  <tbody id='Tabs' style='display:none'>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='300' class='tdbg5'><strong>浏览/查看权限：</strong><br><font color='red'>栏目权限为继承关系。例如：当父栏目设为“认证栏目”时，子栏目的权限设置将继承父栏目设置，即使子栏目设为“开放栏目”也无效。相反，如果父栏目设为“开放栏目”，子栏目可以设为“半开放栏目”或“认证栏目”。</font></td>"
            Response.Write "      <td>"
            Response.Write "        <table>"
            Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' " & RadioValue(rsClass("ClassPurview"), 0) & ">开放栏目</td><td>任何人（包括游客）可以浏览和查看此栏目下的信息。</td></tr>"
            Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' " & RadioValue(rsClass("ClassPurview"), 1) & ">半开放栏目</td><td>任何人（包括游客）都可以浏览。游客不可查看，其他会员根据会员组的栏目权限设置决定是否可以查看。</td></tr>"
            Response.Write "          <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' " & RadioValue(rsClass("ClassPurview"), 2) & ">认证栏目</td><td>游客不能浏览和查看，其他会员根据会员组的栏目权限设置决定是否可以浏览和查看。</td></tr>"
            Response.Write "        </table>"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='300' class='tdbg5'><strong>允许浏览此栏目的会员组：</strong><br>如果栏目是“半开放栏目”或“认证栏目”，请在此设置允许浏览此栏目的会员组</td>"
            Response.Write "      <td>" & GetUserGroup("Browse", ClassID, 5) & "</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='300' class='tdbg5'><strong>允许查看此栏目下信息的会员组：</strong><br>如果栏目是“认证栏目”，请在此设置允许查看此栏目下信息的会员组<br>如果在信息中设置了查看权限，则以信息中的权限设置优先</td>"
            Response.Write "      <td>" & GetUserGroup("View", ClassID, 5) & "</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='300' class='tdbg5'><strong>发表权限：</strong><br>请在此设置允许在此栏目发表信息的会员组。<br>游客没有发表信息权限。</td>"
            Response.Write "      <td>" & GetUserGroup("Input", ClassID, 5) & "</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='300' class='tdbg5'><strong>评论权限：</strong></td>"
            Response.Write "      <td>"
            Response.Write "        <input name='EnableComment' type='checkbox' " & RadioValue(rsClass("EnableComment"), True) & "> 允许在此栏目发表评论<br>"
            Response.Write "        <input name='CheckComment' type='checkbox' " & RadioValue(rsClass("CheckComment"), True) & ">评论需要审核"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "  </tbody>" & vbCrLf
        Else
            Response.Write "  <tbody id='Tabs' style='display:none'>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='300' class='tdbg5'><strong>评论权限：</strong></td>"
            Response.Write "      <td>"
            Response.Write "        <input name='EnableComment' type='checkbox' " & RadioValue(rsClass("EnableComment"), True) & "> 允许在此栏目发表评论<br>"
            Response.Write "        <input name='CheckComment' type='checkbox' " & RadioValue(rsClass("CheckComment"), True) & ">评论需要审核"
            Response.Write "      </td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "  </tbody>" & vbCrLf
        End If
    End If

    '2005-12-23
    '刘永涛 屏蔽掉一些不必要的信息
    If ModuleType = 6 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>发布信息扣除的点数：</strong><br>会员在此栏目下发布信息时，该扣除其会员点数</td>"
        Response.Write "      <td><input name='ReleaseClassPoint' type='text' value='" & rsClass("ReleaseClassPoint") & "' size='4' maxlength='4' style='text-align:center'>&nbsp;点数</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='300' class='tdbg5'><strong>设置栏目推荐每天扣除的点数：</strong></td>"
        Response.Write "      <td><input name='CommandClassPoint' type='text' value='" & rsClass("CommandClassPoint") & "' size='4' maxlength='4' style='text-align:center'>&nbsp;点/天</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>积分奖励：</strong><br>会员在此栏目发表信息时可以得到的积分</td>"
    Response.Write "      <td>会员在此栏目每发表一条信息，可以得到 <input name='PresentExp' type='text' value='" & rsClass("PresentExp") & "' size='4' maxlength='4' style='text-align:center'> 分积分</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>默认收费" & PointName & "数：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，该" & ChannelShortName & "默认的收费" & PointName & "数</td>"
    Response.Write "      <td><input name='DefaultItemPoint' type='text' value='" & rsClass("DefaultItemPoint") & "' size='4' maxlength='4' style='text-align:center'> " & PointUnit & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>默认重复收费：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，该" & ChannelShortName & "默认的重复收费方式</td>"
    Response.Write "      <td>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' " & RadioValue(rsClass("DefaultItemChargeType"), 0) & ">不重复收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' " & RadioValue(rsClass("DefaultItemChargeType"), 1) & ">距离上次收费时间 <input name='DefaultItemPitchTime' type='text' value='" & rsClass("DefaultItemPitchTime") & "' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' " & RadioValue(rsClass("DefaultItemChargeType"), 2) & ">会员重复查看此文章 <input name='DefaultItemReadTimes' type='text' value='" & rsClass("DefaultItemReadTimes") & "' size='8' maxlength='8' style='text-align:center'> 次后重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' " & RadioValue(rsClass("DefaultItemChargeType"), 3) & ">上述两者都满足时重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' " & RadioValue(rsClass("DefaultItemChargeType"), 4) & ">上述两者任一个满足时就重新收费<br>"
    Response.Write "        <input name='DefaultItemChargeType' type='radio' " & RadioValue(rsClass("DefaultItemChargeType"), 5) & ">每阅读一次就重复收费一次（建议不要使用）"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>默认分成比例：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，该" & ChannelShortName & "默认的分成比例</td>"
    Response.Write "      <td><input name='DefaultItemDividePercent' type='text' value='" & rsClass("DefaultItemDividePercent") & "' size='4' maxlength='4' style='text-align:center'> %</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    End If


    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>打开方式：</strong></td>"
    Response.Write "      <td><input name='OpenType' type='radio' " & RadioValue(rsClass("OpenType"), 0) & ">在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' " & RadioValue(rsClass("OpenType"), 1) & ">在新窗口打开</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在顶部导航栏显示：</strong><br>此选项只对一级栏目有效。</td>"
    Response.Write "      <td><input name='ShowOnTop' type='radio' " & RadioValue(rsClass("ShowOnTop"), True) & ">是&nbsp;&nbsp;&nbsp;&nbsp; <input name='ShowOnTop' type='radio' " & RadioValue(rsClass("ShowOnTop"), False) & ">否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在频道首页分类列表处显示：</strong><br>此选项只对一级栏目有效。如果一级栏目比较多，但首页不想显示太多的分类列表，这个选项就非常有用。</td>"
    Response.Write "      <td><input name='ShowOnIndex' type='radio' " & RadioValue(rsClass("ShowOnIndex"), True) & ">是&nbsp;&nbsp;&nbsp;&nbsp; <input name='ShowOnIndex' type='radio' " & RadioValue(rsClass("ShowOnIndex"), False) & ">否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在父栏目的分类列表处显示：</strong><br>如果某栏目下有几十个子栏目，但只想显示其中几个子栏目的" & ChannelShortName & "列表，这个选项就非常有用。</td>"
    Response.Write "      <td><input name='IsElite' type='radio' " & RadioValue(rsClass("IsElite"), True) & ">是&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' " & RadioValue(rsClass("IsElite"), False) & ">否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>有子栏目时是否可以在此栏目添加" & ChannelShortName & "：</strong></td>"
    Response.Write "      <td><input name='EnableAdd' type='radio' " & RadioValue(rsClass("EnableAdd"), True) & ">是&nbsp;&nbsp;&nbsp;&nbsp; <input name='EnableAdd' type='radio' " & RadioValue(rsClass("EnableAdd"), False) & ">否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否启用此栏目的防止复制、防盗链功能：</strong></td>"
    Response.Write "      <td><input name='EnableProtect' type='radio' " & RadioValue(rsClass("EnableProtect"), True) & ">是&nbsp;&nbsp;&nbsp;&nbsp; <input name='EnableProtect' type='radio' " & RadioValue(rsClass("EnableProtect"), False) & ">否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='SkinID' id='SkinID'>" & GetSkin_Option(rsClass("SkinID")) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目模板：</strong><br>相关模板中包含了栏目设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色风格”失效。 </td>"
    Response.Write "      <td><select name='TemplateID' id='TemplateID'>" & GetTemplate_Option(ChannelID, 2, rsClass("TemplateID")) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>每页显示的" & ChannelShortName & "数：</strong><br>当此栏目为最下一级栏目时，则会分页显示此栏目中的" & ChannelShortName & "，这里指定的是每页显示的" & ChannelShortName & "数。</td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, rsClass("MaxPerPage")) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>此栏目下的" & ChannelShortName & "的默认配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='DefaultItemSkin' id='DefaultItemSkin'>" & GetSkin_Option(rsClass("DefaultItemSkin")) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><strong>此栏目下的" & ChannelShortName & "的默认模板：</strong></td>"
    Response.Write "      <td><select name='DefaultItemTemplate' id='DefaultItemTemplate'>" & GetTemplate_Option(ChannelID, 3, rsClass("DefaultItemTemplate")) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><b>此栏目下的" & ChannelShortName & "列表的排序方式：</b></td>"
    Response.Write "      <td><select name='ItemListOrderType'>" & GetOrderType_Option(rsClass("ItemListOrderType")) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300' class='tdbg5'><b>此栏目下的" & ChannelShortName & "打开方式：</b></td>"
    Response.Write "      <td><select name='ItemOpenType'>" & GetOpenType_Option(rsClass("ItemOpenType")) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Call EditCustom_Content("Modify", rsClass("Custom_Content"), "Class")
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='2' align='center'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "      <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      <input name='ClassID' type='hidden' id='ClassID' value='" & rsClass("ClassID") & "'>"
    Response.Write "      <input name='Modify' type='submit' value=' 保存修改结果 ' style='cursor:hand;'>&nbsp;&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Call WriteJS
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub MoveClass()
    Dim tChannelID
    Dim ClassID, sql, rsClass, i
    Dim SkinID, TemplateID
    tChannelID = Trim(Request("tChannelID"))
    ClassID = Trim(Request("ClassID"))
    If tChannelID = "" Then
        tChannelID = ChannelID
    Else
        tChannelID = PE_CLng(tChannelID)
    End If
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    
    sql = "select * from PE_Class where ClassID=" & ClassID
    Set rsClass = Server.CreateObject("Adodb.recordset")
    rsClass.Open sql, Conn, 1, 3
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
    Else
        Response.Write "<form name='myform' method='post' action='Admin_Class.asp'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='3' align='center'><strong>移动" & ChannelShortName & "栏目</strong></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='left' valign='top' width='260'><strong>当前栏目：</strong><br><select name='ClassID2' size='2' style='height:330px;width:260px;' disabled>" & GetClass_Option(ChannelID, ClassID) & "</select></td>"
        Response.Write "      <td align='center' width='70'><strong>移动到&gt;&gt;&gt;</strong></td>"
        Response.Write "      <td align='left'>"
        Response.Write "        <strong>目标频道：</strong><select name='tChannelID' onChange='document.myform.submit();'>" & GetChannel_Option(ModuleType, tChannelID) & "</select><br>"
        Response.Write "        <strong>目标栏目：</strong><font color=red>（不能指定为当前栏目的下属子栏目或外部栏目）</font><br><select name='ParentID' size='2' style='height:300px;width:260px;'><option value='0'>无（做为一级栏目）</option>" & GetClass_Option(tChannelID, rsClass("ParentID")) & "</select>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='40' colspan='3' align='center'>"
        Response.Write "        <input name='Action' type='hidden' id='Action' value='Move'>"
        Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "        <input name='ClassID' type='hidden' id='ClassID' value='" & ClassID & "'>"
        Response.Write "        <input name='Submit' type='submit' value=' 保存移动结果 ' style='cursor:hand;' onClick=""document.myform.Action.value='SaveMove';"">&nbsp;&nbsp;"
        Response.Write "        <input name='Cancel' type='button' value=' 取 消 ' style='cursor:hand;' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"">"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub order()
    Dim sqlClass, rsClass, i, iCount, j
    sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 order by RootID"
    Set rsClass = Server.CreateObject("adodb.recordset")
    rsClass.Open sqlClass, Conn, 1, 1
    iCount = rsClass.RecordCount

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4' align='center'><strong>一 级 栏 目 排 序</strong></td>"
    Response.Write "  </tr>"
    j = 1
    Do While Not rsClass.EOF

        Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "      <td width='200'>" & rsClass("ClassName") & "</td>"
     
        If j > 1 Then
            Response.Write "<form action='Admin_Class.asp?Action=UpOrder' method='post'><td width='150'>"
            Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select>"
            Response.Write "<input type=hidden name=ClassID value=" & rsClass("ClassID") & "><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=cRootID value=" & rsClass("RootID") & ">&nbsp;<input type=submit name=Submit value=修改>"
            Response.Write "</td></form>"
        Else
            Response.Write "<td width='150'>&nbsp;</td>"
        End If
        If iCount > j Then
            Response.Write "<form action='Admin_Class.asp?Action=DownOrder' method='post'><td width='150'>"
            Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select>"
            Response.Write "<input type=hidden name=ClassID value=" & rsClass("ClassID") & "><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=cRootID value=" & rsClass("RootID") & ">&nbsp;<input type=submit name=Submit value=修改>"
            Response.Write "</td></form>"
        Else
            Response.Write "<td width='150'>&nbsp;</td>"
        End If
        Response.Write "      <td>&nbsp;</td>"
        Response.Write " </tr>"
        j = j + 1
        rsClass.MoveNext
    Loop
    Response.Write "</table>"
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub OrderN()
    Dim sqlClass, rsClass, i, iCount, trs, UpMoveNum, DownMoveNum
    sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Server.CreateObject("adodb.recordset")
    rsClass.Open sqlClass, Conn, 1, 1
    
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4' align='center'><strong>N 级 栏 目 排 序</strong></td>"
    Response.Write "  </tr>"

    Do While Not rsClass.EOF
        Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "      <td width='300'>"
        For i = 1 To rsClass("Depth")
            Response.Write "&nbsp;&nbsp;&nbsp;"
        Next
        If rsClass("Child") > 0 Then
            Response.Write "<img src='../images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
        Else
            Response.Write "<img src='../images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
        End If
        If rsClass("ParentID") = 0 Then
            Response.Write "<b>"
        End If
        Response.Write rsClass("ClassName")
        If rsClass("Child") > 0 Then
            Response.Write "(" & rsClass("Child") & ")"
        End If
        Response.Write "</td>"
        If rsClass("ParentID") > 0 Then '如果不是一级栏目，则算出相同深度的栏目数目，得到该栏目在相同深度的栏目中所处位置（之上或者之下的栏目数）
            '所能提升最大幅度应为For i=1 to 该版之上的版面数
            Set trs = Conn.Execute("select count(ClassID) from PE_Class where ParentID=" & rsClass("ParentID") & " and OrderID<" & rsClass("OrderID") & "")
            UpMoveNum = trs(0)
            If IsNull(UpMoveNum) Then UpMoveNum = 0
            If UpMoveNum > 0 Then
                Response.Write "<form action='Admin_Class.asp?Action=UpOrderN' method='post'><td width='150'>"
                Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>"
                For i = 1 To UpMoveNum
                    Response.Write "<option value=" & i & ">" & i & "</option>"
                Next
                Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
                Response.Write "<input type=hidden name=ClassID value=" & rsClass("ClassID") & ">&nbsp;<input type=submit name=Submit value=修改>"
                Response.Write "</td></form>"
            Else
                Response.Write "<td width='150'>&nbsp;</td>"
            End If
            trs.Close
            '所能降低最大幅度应为For i=1 to 该版之下的版面数
            Set trs = Conn.Execute("select count(ClassID) from PE_Class where ParentID=" & rsClass("ParentID") & " and orderID>" & rsClass("orderID") & "")
            DownMoveNum = trs(0)
            If IsNull(DownMoveNum) Then DownMoveNum = 0
            If DownMoveNum > 0 Then
                Response.Write "<form action='Admin_Class.asp?Action=DownOrderN' method='post'><td width='150'>"
                Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>"
                For i = 1 To DownMoveNum
                    Response.Write "<option value=" & i & ">" & i & "</option>"
                Next
                Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
                Response.Write "<input type=hidden name=ClassID value=" & rsClass("ClassID") & ">&nbsp;<input type=submit name=Submit value=修改>"
                Response.Write "</td></form>"
            Else
                Response.Write "<td width='150'>&nbsp;</td>"
            End If
            trs.Close
        Else
            Response.Write "<td colspan=2>&nbsp;</td>"
        End If
        Response.Write "      <td>&nbsp;</td>"
        Response.Write " </tr>"

        UpMoveNum = 0
        DownMoveNum = 0
        rsClass.MoveNext
    Loop
    Response.Write "</table>"
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub Reset()
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='3' align='center'><strong>复位所有" & ChannelShortName & "栏目</strong></td> "
    Response.Write "  </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "    <td align='center'>"
    Response.Write "      <form name='form1' method='post' action='Admin_Class.asp?Action=SaveReset'>"
    Response.Write "        <table width='80%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "            <td height='150'><font color='#FF0000'><strong>注意：</strong></font><br>&nbsp;&nbsp;&nbsp;&nbsp;如果选择复位所有栏目，则所有栏目都将作为一级栏目，这时您需要重新对各个栏目进行归属的基本设置。不要轻易使用该功能，仅在做出了错误的设置而无法复原栏目之间的关系和排序的时候使用。<br><br>&nbsp;&nbsp;&nbsp;&nbsp;如果复位时存在着同名栏目，则系统会自动将目录名进行重命名。<br><br>&nbsp;&nbsp;&nbsp;&nbsp;复位成功后，请记得一定要重新生成所有HTML的内容。"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "        <input type='submit' name='Submit' value='复位所有栏目'> &nbsp; <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "     <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      </form></td>"
    Response.Write "    </tr>"
    Response.Write "</table>"
End Sub


Sub Unite()
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='3' align='center'><strong>" & ChannelShortName & "栏目合并</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='100'><form name='myform' method='post' action='Admin_Class.asp' onSubmit='return ConfirmUnite();'>"
    Response.Write "        &nbsp;&nbsp;将栏目 <select name='ClassID' id='ClassID'>" & GetClass_Option(ChannelID, 0) & "</select> &nbsp;&nbsp;合并到 <select name='TargetClassID' id='TargetClassID'>" & GetClass_Option(ChannelID, 0) & "</select><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "     <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveUnite'>"
    Response.Write "        <input type='submit' name='Submit' value=' 合并栏目 ' style='cursor:hand;'>"
    Response.Write "        &nbsp;&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "      </form>"
    Response.Write " </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='60'><strong>注意事项：</strong><br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;所有操作不可逆，请慎重操作！！！<br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;不能在同一个栏目内进行操作，不能将一个栏目合并到其下属栏目中。目标栏目中不能含有子栏目。<br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;合并后您所指定的栏目（或者包括其下属栏目）将被删除，所有" & ChannelShortName & "将转移到目标栏目中。</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<script language='JavaScript' type='text/JavaScript'>"
    Response.Write "function ConfirmUnite(){"
    Response.Write "  if (document.myform.ClassID.value==document.myform.TargetClassID.value){"
    Response.Write "    alert('请不要在相同栏目内进行操作！');"
    Response.Write " document.myform.TargetClassID.focus();"
    Response.Write " return false;}"
    Response.Write "  if (document.myform.TargetClassID.value==''){"
    Response.Write "    alert('目标栏目不能指定为含有子栏目的栏目！');"
    Response.Write " document.myform.TargetClassID.focus();"
    Response.Write " return false;}"
    Response.Write "}"
    Response.Write "</script>"
End Sub

Sub ShowBatch()
    Response.Write "<form name='form1' method='post' action='Admin_Class.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='3' align='center'><strong>批量设置" & ChannelShortName & "栏目属性</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' valign='top'><font color='red'>提示：</font>可以按住“Shift”<br>或“Ctrl”键进行多个栏目的选择<br>"
    Response.Write "      <select name='ClassID' size='2' multiple style='height:380px;width:200px;'>" & GetClass_Option(ChannelID, 0) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  选定所有栏目  ' onclick='SelectAll()'><br>"
    Response.Write "      <input type='button' name='Submit' value='取消选定所有栏目' onclick='UnSelectAll()'></div></td>"
    Response.Write "      <td valign='top'><br>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>权限设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'"
    If ModuleType = 5 Then Response.Write " style='display:none'"
    Response.Write ">收费设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>栏目选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "<table width='99%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
 
    If ModuleType = 6 Then
        Response.Write "  <tbody id='Tabs' style='display:'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyGroupPurview_Input' value='Yes'></td>"
        Response.Write "      <td width='150' class='tdbg5'><strong>发表权限：</strong><br>请在此设置允许在此栏目发表信息的会员组。<br>游客没有发表信息权限。</td>"
        Response.Write "      <td>" & GetUserGroup("Input", 0, 3) & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyEnableComment' value='Yes'><br><input type='checkbox' name='ModifyCheckComment' value='Yes'></td>"
        Response.Write "      <td width='150' class='tdbg5'><strong>评论权限：</strong></td>"
        Response.Write "      <td><input name='EnableComment' type='checkbox' value='True' checked>允许在此栏目发表评论<br><input name='CheckComment' type='checkbox' value='True' checked>评论需要审核</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
        If ModuleType <> 5 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyClassPurview' value='Yes'></td>"
            Response.Write "      <td width='300' class='tdbg5'><strong>浏览/查看权限：</strong><br><font color='red'>栏目权限为继承关系。例如：当父栏目设为“认证栏目”时，子栏目的权限设置将继承父栏目设置，即使子栏目设为“开放栏目”也无效。相反，如果父栏目设为“开放栏目”，子栏目可以设为“半开放栏目”或“认证栏目”。</font></td>"
            Response.Write "      <td><table><tr><td width='80' valign='top'><input type='radio' name='ClassPurview' value='0' checked>开放栏目</td><td>任何人（包括游客）可以浏览和查看此栏目下的信息。</td></tr>"
            Response.Write "        <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' value='1'>半开放栏目</td><td>任何人（包括游客）都可以浏览。游客不可查看，其他会员根据会员组的栏目权限设置决定是否可以查看。</td></tr>"
            Response.Write "        <tr><td width='80' valign='top'><input type='radio' name='ClassPurview' value='2'>认证栏目</td><td>游客不能浏览和查看，其他用户根据会员组的栏目权限设置决定是否可以浏览和查看。</td></tr></table></td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyGroupPurview_Browse' value='Yes'></td>"
            Response.Write "      <td width='300' class='tdbg5'><strong>允许浏览此栏目的会员组：</strong><br>如果栏目是“半开放栏目”或“认证栏目”，请在此设置允许浏览此栏目的会员组</td>"
            Response.Write "      <td>" & GetUserGroup("Browse", 0, 3) & "</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyGroupPurview_View' value='Yes'></td>"
            Response.Write "      <td width='300' class='tdbg5'><strong>允许查看此栏目下信息的会员组：</strong><br>如果栏目是“认证栏目”，请在此设置允许查看此栏目下信息的会员组<br>如果在信息中设置了查看权限，则以信息中的权限设置优先</td>"
            Response.Write "      <td>" & GetUserGroup("View", 0, 3) & "</td>"
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyGroupPurview_Input' value='Yes'></td>"
            Response.Write "      <td width='300' class='tdbg5'><strong>发表权限：</strong><br>请在此设置允许在此栏目发表信息的会员组。<br>游客没有发表信息权限。</td>"
            Response.Write "      <td>" & GetUserGroup("Input", 0, 3) & "</td>"
            Response.Write "    </tr>" & vbCrLf
        End If
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyEnableComment' value='Yes'><br><input type='checkbox' name='ModifyCheckComment' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>评论权限：</strong></td>"
        Response.Write "      <td><input name='EnableComment' type='checkbox' value='True' checked>允许在此栏目发表评论<br><input name='CheckComment' type='checkbox' value='True' checked>评论需要审核</td>"
        Response.Write "    </tr>" & vbCrLf
    End If
    Response.Write "  </tbody>" & vbCrLf
    If ModuleType = 6 Then
        Response.Write "  <tbody id='Tabs' style='display:none'>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyReleasePoint' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>发布信息扣除的点数：</strong><br>会员在此栏目下发布信息时，该扣除其会员点数</td>"
        Response.Write "      <td><input name='ReleaseClassPoint' type='text' value='0' size='4' maxlength='4' style='text-align:center'>&nbsp;点数</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyCommandClassPoint' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>设置栏目推荐每天扣除的点数：</strong></td>"
        Response.Write "      <td><input name='CommandClassPoint' type='text' value='0' size='4' maxlength='4' style='text-align:center'>&nbsp;点/天</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    Else
        Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyPresentExp' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>积分奖励：</strong><br>会员在此栏目发表信息时可以得到的积分</td>"
        Response.Write "      <td>在此栏目每发表一条信息，可以得到 <input name='PresentExp' type='text' value='1' size='4' maxlength='4' style='text-align:center'> 分积分</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyDefaultItemPoint' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>默认收费" & PointName & "数：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，系统默认的收费" & PointName & "数</td>"
        Response.Write "      <td><input name='DefaultItemPoint' type='text' value='0' size='4' maxlength='4' style='text-align:center'> " & PointUnit & "</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyDefaultItemChargeType' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>默认重复收费：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，系统默认的重复收费方式</td>"
        Response.Write "      <td><input name='DefaultItemChargeType' type='radio' value='0' checked>不重复收费<br>"
        Response.Write "      <input name='DefaultItemChargeType' type='radio' value='1'>距离上次收费时间 <input name='DefaultItemPitchTime' type='text' value='24' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>"
        Response.Write "      <input name='DefaultItemChargeType' type='radio' value='2'>会员重复查看此文章 <input name='DefaultItemReadTimes' type='text' value='10' size='8' maxlength='8' style='text-align:center'> 次后重新收费<br>"
        Response.Write "      <input name='DefaultItemChargeType' type='radio' value='3'>上述两者都满足时重新收费<br>"
        Response.Write "      <input name='DefaultItemChargeType' type='radio' value='4'>上述两者任一个满足时就重新收费<br>"
        Response.Write "      <input name='DefaultItemChargeType' type='radio' value='5'>每阅读一次就重复收费一次（建议不要使用）"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyDefaultItemDividePercent' value='Yes'></td>"
        Response.Write "      <td width='300' class='tdbg5'><strong>默认分成比例：</strong><br>会员在此栏目下添加" & ChannelShortName & "时，系统默认的分成比例</td>"
        Response.Write "      <td><input name='DefaultItemDividePercent' type='text' value='0' size='4' maxlength='4' style='text-align:center'> %</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </tbody>" & vbCrLf
    End If
    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyOpenType' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>打开方式：</strong></td>"
    Response.Write "      <td><input type='radio' name='OpenType' value='0' checked>在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1'>在新窗口打开</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyShowOnTop' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在顶部导航栏显示：</strong>此选项只对一级栏目有效。</td>"
    Response.Write "      <td><input name='ShowOnTop' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input type='radio' name='ShowOnTop' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyShowOnIndex' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在首页分类列表处显示：</strong><br>此选项只对一级栏目有效。如果一级栏目比较多，但首页不想显示太多的分类列表，这个选项就非常有用。</td>"
    Response.Write "      <td><input name='ShowOnIndex' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input type='radio' name='ShowOnIndex' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyIsElite' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否在父栏目的分类列表处显示：</strong><br>如果某栏目下有几十个子栏目，但只想显示其中几个子栏目的" & ChannelShortName & "列表，这个选项就非常有用。</td>"
    Response.Write "      <td><input name='IsElite' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input type='radio' name='IsElite' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyEnableAdd' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>有子栏目时是否可以在此栏目添加" & ChannelShortName & "：</strong></td>"
    Response.Write "      <td><input name='EnableAdd' type='radio' value='True'>是&nbsp;&nbsp;&nbsp;&nbsp; <input type='radio' name='EnableAdd' value='False' checked>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyEnableProtect' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否启用此栏目的防止复制、防盗链功能：</strong></td>"
    Response.Write "      <td><input name='EnableProtect' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input type='radio' name='EnableProtect' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifySkinID' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='SkinID' id='SkinID'>" & GetSkin_Option(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyTemplateID' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>栏目模板：</strong><br>相关模板中包含了栏目设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色风格”失效。 </td>"
    Response.Write "      <td><select name='TemplateID' id='TemplateID'>" & GetTemplate_Option(ChannelID, 2, 0) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyMaxPerPage' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>每页显示的" & ChannelShortName & "数：</strong><br>当此栏目为最下一级栏目时，则会分页显示此栏目中的" & ChannelShortName & "，这里指定的是每页显示的" & ChannelShortName & "数。</td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, 20) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyDefaultItemSkin' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>此栏目下的" & ChannelShortName & "的默认配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='DefaultItemSkin' id='DefaultItemSkin'>" & GetSkin_Option(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyDefaultItemTemplate' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>此栏目下的" & ChannelShortName & "的默认模板：</strong></td>"
    Response.Write "      <td><select name='DefaultItemTemplate' id='DefaultItemTemplate'>" & GetTemplate_Option(ChannelID, 3, 1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyItemListOrderType' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><b>此栏目下的" & ChannelShortName & "列表的排序方式：</b></td>"
    Response.Write "      <td><select name='ItemListOrderType'>" & GetOrderType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyItemOpenType' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><b>此栏目下的" & ChannelShortName & "打开方式：</b></td>"
    Response.Write "      <td><select name='ItemOpenType'>" & GetOpenType_Option(1) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "  </td></tr></table>"
    Response.Write "  <br><b>说明：</b><br>1、若要批量修改某个属性的值，请先选中其左侧的复选框，然后再设定属性值。<br>2、这里显示的属性值都是系统默认值，与所选栏目的已有属性无关<br>"
    Response.Write "  <p align='center'><input name='Action' type='hidden' id='Action' value='DoBatch'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Submit' type='submit' value=' 执行批处理 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'></p>"
    Response.Write "  </td></tr></table>"
    Response.Write "</form>" & vbCrLf
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.form1.ClassID.length;i++){" & vbCrLf
    Response.Write "    document.form1.ClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.form1.ClassID.length;i++){" & vbCrLf
    Response.Write "    document.form1.ClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Call WriteJS
End Sub

Sub Patch()
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='3' align='center'><strong>修复栏目结构</strong></td> "
    Response.Write "  </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "    <td align='center'>"
    Response.Write "      <form name='form1' method='post' action='Admin_Class.asp?Action=DoPatch'>"
    Response.Write "        <table width='80%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "            <td height='150'><br>当栏目出现排序错误或串位的情况时，使用此功能可以修复。本操作相当安全，不会给系统带来任何负面影响。<br><br>修复过程中请勿刷新页面！"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "        <input type='submit' name='Submit' value='开始修复'> &nbsp; <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Class.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "     <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      </form></td>"
    Response.Write "    </tr>"
    Response.Write "</table>"
End Sub

Sub DoPatch()
    Dim rsClass, sql, PrevID, trs
    Set rsClass = Server.CreateObject("ADODB.Recordset")
    sql = "select ClassID,RootID,OrderID,Depth,ParentID,ParentPath,Child,arrChildID,PrevID,NextID,ClassType,ParentDir,ClassDir,ClassPurview,ItemCount from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 order by RootID"
    rsClass.Open sql, Conn, 1, 3
    If rsClass.BOF And rsClass.EOF Then
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    End If
    PrevID = 0
    Do While Not rsClass.EOF
        rsClass("OrderID") = 0
        rsClass("Depth") = 0
        rsClass("ParentPath") = "0"
        rsClass("PrevID") = PrevID
        rsClass("NextID") = 0
        rsClass("arrChildID") = CStr(rsClass("ClassID"))
        If rsClass("ClassType") = 1 Then
            rsClass("ParentDir") = "/"
        End If
        If PrevID <> rsClass("ClassID") And PrevID > 0 Then
            Conn.Execute ("update PE_Class set NextID=" & rsClass("ClassID") & " where ClassID=" & PrevID & "")
        End If
        PrevID = rsClass("ClassID")
        
        If ModuleType = 5 Then
            Set trs = Conn.Execute("select count(0) from " & SheetName & " where ClassID=" & rsClass("ClassID") & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & "")
        Else
            Set trs = Conn.Execute("select count(0) from " & SheetName & " where ClassID=" & rsClass("ClassID") & " and Status=3 and Deleted=" & PE_False & "")
        End If
        If IsNull(trs(0)) Then
            rsClass("ItemCount") = 0
        Else
            rsClass("ItemCount") = trs(0)
        End If
        Set trs = Nothing
        
        
        rsClass.Update
        iOrderID = 1
        Call UpdateClass(rsClass("ClassID"), 1, "0", "/" & rsClass("ClassDir") & "/", rsClass("ClassPurview"))
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    Call WriteSuccessMsg("修复栏目结构成功！", ComeUrl)
    
End Sub

Sub UpdateClass(iParentID, iDepth, sParentPath, sParentDir, ClassPurview)
    Dim rsClass, sql, PrevID, ParentPath, trs, rsChild
    ParentPath = sParentPath & "," & iParentID
    
    sql = "select ClassID,RootID,OrderID,Depth,ParentID,ParentPath,Child,arrChildID,PrevID,NextID,ClassType,ParentDir,ClassDir,ClassPurview,ItemCount from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & iParentID & " order by OrderID"
    Set rsClass = Server.CreateObject("ADODB.Recordset")
    rsClass.Open sql, Conn, 1, 3
    If rsClass.BOF And rsClass.EOF Then
        Conn.Execute ("update PE_Class set Child=0 where ClassID=" & iParentID & "")
    Else
        Conn.Execute ("update PE_Class set Child=" & rsClass.RecordCount & " where ClassID=" & iParentID & "")
        
        PrevID = 0
        Do While Not rsClass.EOF
            Set rsChild = Server.CreateObject("adodb.recordset")
            rsChild.Open "select arrChildID from PE_Class where ClassID in (" & ParentPath & ")", Conn, 1, 3
            Do While Not rsChild.EOF
                rsChild(0) = rsChild(0) & "," & rsClass("ClassID")
                rsChild.Update
                rsChild.MoveNext
            Loop
            rsChild.Close
            Set rsChild = Nothing
            
            rsClass("OrderID") = iOrderID
            rsClass("Depth") = iDepth
            rsClass("ParentPath") = ParentPath
            rsClass("PrevID") = PrevID
            rsClass("NextID") = 0
            rsClass("arrChildID") = CStr(rsClass("ClassID"))
            If rsClass("ClassType") = 1 Then
                rsClass("ParentDir") = sParentDir
            End If
            If rsClass("ClassPurview") < ClassPurview Then
                rsClass("ClassPurview") = ClassPurview
            End If
            
            If PrevID <> rsClass("ClassID") And PrevID > 0 Then
                Conn.Execute ("update PE_Class set NextID=" & rsClass("ClassID") & " where ClassID=" & PrevID & "")
            End If
            PrevID = rsClass("ClassID")
            
            If ModuleType = 5 Then
                Set trs = Conn.Execute("select count(0) from " & SheetName & " where ClassID=" & rsClass("ClassID") & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & "")
            Else
                Set trs = Conn.Execute("select count(0) from " & SheetName & " where ClassID=" & rsClass("ClassID") & " and Status=3 and Deleted=" & PE_False & "")
            End If
            If IsNull(trs(0)) Then
                rsClass("ItemCount") = 0
            Else
                rsClass("ItemCount") = trs(0)
            End If
            Set trs = Nothing
            
            rsClass.Update
            
            iOrderID = iOrderID + 1
            
            Call UpdateClass(rsClass("ClassID"), iDepth + 1, ParentPath, sParentDir & rsClass("ClassDir") & "/", rsClass("ClassPurview"))
            
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
End Sub

Sub CheckClassDepth()
    Dim strSql
    strSql = "Select Depth From PE_Class Where ClassId=" & ParentID & ""
    If PE_CLng(Conn.Execute(strSql)(0)) > 3 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>供求模块下,子栏目层数不可以超过4层！</li>"
    End If
End Sub
Sub SaveAdd()
    Dim ClassID, ClassName, ClassType, LinkUrl, ClassDir, ClassPicUrl, Tips, ReadMe, Meta_Keywords, Meta_Description
    Dim ClassPurview, arrGroupID_Browse, arrGroupID_View, arrGroupID_Input, EnableComment, CheckComment
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent
    Dim OpenType, ShowOnTop, ShowOnIndex, IsElite, EnableAdd, EnableProtect, SkinID, TemplateID
    Dim MaxPerPage, DefaultItemSkin, DefaultItemTemplate, ItemListOrderType, ItemOpenType
    Dim sql, rs, trs, rsClass
    Dim RootID, ParentDepth, ParentPath, ParentStr, ParentName, MaxClassID, MaxRootID, arrChildID, ParentDir, PrevOrderID
    Dim PrevID, NextID, Child, strClassDir
    Dim ReleaseClassPoint, CommandClassPoint '在栏目下发布信息要扣除的会员点数和设置栏目推荐要扣除的会员点数

    ClassName = Trim(Request("ClassName"))
    ClassType = PE_CLng(Trim(Request("ClassType")))
    LinkUrl = Trim(Request("LinkUrl"))
    ClassDir = Trim(Request("ClassDir"))
    ClassPicUrl = Trim(Request("ClassPicUrl"))
    Tips = Trim(Request("Tips"))
    ReadMe = Trim(Request("Readme"))
    Meta_Keywords = Trim(Request("Meta_Keywords"))
    Meta_Description = Trim(Request("Meta_Description"))

    ClassPurview = PE_CLng(Trim(Request("ClassPurview")))
    arrGroupID_Browse = ReplaceBadChar(Trim(Request("arrGroupID_Browse")))
    arrGroupID_View = ReplaceBadChar(Trim(Request("arrGroupID_View")))
    arrGroupID_Input = ReplaceBadChar(Trim(Request("arrGroupID_Input")))
    EnableComment = PE_CBool(Trim(Request("EnableComment")))
    CheckComment = PE_CBool(Trim(Request("CheckComment")))

    PresentExp = PE_CDbl(Trim(Request("PresentExp")))
    DefaultItemPoint = PE_CDbl(Trim(Request("DefaultItemPoint")))
    DefaultItemChargeType = PE_CLng(Trim(Request.Form("DefaultItemChargeType")))
    DefaultItemPitchTime = PE_CLng(Trim(Request.Form("DefaultItemPitchTime")))
    DefaultItemReadTimes = PE_CLng(Trim(Request.Form("DefaultItemReadTimes")))
    DefaultItemDividePercent = PE_CLng(Trim(Request.Form("DefaultItemDividePercent")))

    OpenType = PE_CLng(Trim(Request("OpenType")))
    ShowOnTop = PE_CBool(Trim(Request("ShowOnTop")))
    ShowOnIndex = PE_CBool(Trim(Request("ShowOnIndex")))
    IsElite = PE_CBool(Trim(Request("IsElite")))
    EnableAdd = PE_CBool(Trim(Request("EnableAdd")))
    EnableProtect = PE_CBool(Trim(Request("EnableProtect")))
    SkinID = PE_CLng(Trim(Request("SkinID")))
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    DefaultItemSkin = PE_CLng(Trim(Request("DefaultItemSkin")))
    DefaultItemTemplate = PE_CLng(Trim(Request("DefaultItemTemplate")))
    ItemListOrderType = PE_CLng(Trim(Request("ItemListOrderType")))
    ItemOpenType = PE_CLng(Trim(Request("ItemOpenType")))
    
    ReleaseClassPoint = PE_CLng(Trim(Request("ReleaseClassPoint")))
    CommandClassPoint = PE_CLng(Trim(Request("CommandClassPoint")))
    If ModuleType = 6 Then
        Call CheckClassDepth
    End If
    If ClassName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>栏目名称不能为空！</li>"
    Else
        ClassName = ReplaceBadChar(ClassName)
    End If
    If ClassType > 1 Then
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址不能为空！</li>"
        End If
    Else
        If ClassDir = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>栏目目录不能为空！</li>"
        Else
            If IsValidStr(ClassDir) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>栏目目录名只能为英文字母和数字的组合，且第一个字符必须为英文字母！</li>"
            Else
                If CheckValidStr(arrInvalidDir, ClassDir) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>栏目目录名不能为“" & arrInvalidDir & "”等系统目录！</li>"
                Else
                End If
                If IsNumeric(ClassDir) = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>栏目目录名不能为数字！"
                End If
            End If
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    Set trs = Conn.Execute("Select * from PE_Class Where ChannelID=" & ChannelID & " and ParentID=" & ParentID & " AND ClassName='" & ClassName & "'")
    If Not (trs.BOF And trs.EOF) Then
        FoundErr = True
        If ParentID = 0 Then
            ErrMsg = ErrMsg & "<li>已经存在一级栏目：" & ClassName & "</li>"
        Else
            ErrMsg = ErrMsg & "<li>“" & ParentName & "”中已经存在子栏目“" & ClassName & "”！</li>"
        End If
    End If
    trs.Close
    Set trs = Nothing

    If ClassType = 1 Then
        Select Case StructureType
        Case 0, 1, 2
            Set trs = Conn.Execute("select ClassID from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & ParentID & " and ClassDir='" & ClassDir & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>指定的栏目目录已经存在！</li>"
            End If
            trs.Close
            Set trs = Nothing
        Case 3, 4, 5
            Set trs = Conn.Execute("select ClassID from PE_Class where ChannelID=" & ChannelID & " and ClassDir='" & ClassDir & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>指定的栏目目录已经存在！</li>"
            End If
            trs.Close
            Set trs = Nothing
        Case Else
            '不用判断
        End Select
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Set rs = Conn.Execute("select Max(ClassID) from PE_Class")
    MaxClassID = rs(0)
    If IsNull(MaxClassID) Then
        MaxClassID = 0
    End If
    rs.Close
    Set rs = Nothing
    ClassID = MaxClassID + 1
    
    Set rs = Conn.Execute("select max(rootid) from PE_Class where ChannelID=" & ChannelID & "")
    MaxRootID = rs(0)
    If IsNull(MaxRootID) Then
        MaxRootID = 0
    End If
    rs.Close
    Set rs = Nothing
    RootID = MaxRootID + 1
    

    If ParentID > 0 Then
        Set rs = Conn.Execute("select * from PE_Class where ClassID=" & ParentID & "")
        If rs.BOF And rs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>所属栏目已经被删除！</li>"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        If rs("ClassType") = 2 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能指定外部栏目为所属栏目！</li>"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If

        RootID = rs("RootID")
        ParentName = rs("ClassName")
        ParentDepth = rs("Depth")
        ParentPath = rs("ParentPath") & "," & rs("ClassID")   '得到此栏目的父级栏目路径
        Child = rs("Child")
        arrChildID = rs("arrChildID") & "," & ClassID
        ParentDir = rs("ParentDir") & rs("ClassDir") & "/"

        '更新本栏目的所有上级栏目的子栏目ID数组
        Set trs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & ParentPath & ")")
        Do While Not trs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & trs(1) & "," & ClassID & "' where ClassID=" & trs(0))
            trs.MoveNext
        Loop
        trs.Close


        If Child > 0 Then
            Dim rsPrevOrderID
            '得到父栏目的所有子栏目中最后一个栏目的OrderID
            Set rsPrevOrderID = Conn.Execute("select Max(OrderID) from PE_Class where ClassID in ( " & arrChildID & ")")
            PrevOrderID = rsPrevOrderID(0)
            Set rsPrevOrderID = Nothing
            
            '得到本栏目的上一个栏目ID
            Set trs = Conn.Execute("select top 1 ClassID from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & ParentID & " order by OrderID desc")
            PrevID = trs(0)
            trs.Close
        Else
            PrevOrderID = rs("OrderID")
            PrevID = 0
        End If

        rs.Close
        Set rs = Nothing
    Else
        If MaxRootID > 0 Then
            Set trs = Conn.Execute("select ClassID from PE_Class where ChannelID=" & ChannelID & " and RootID=" & MaxRootID & " and Depth=0")
            PrevID = trs(0)
            trs.Close
        Else
            PrevID = 0
        End If
        PrevOrderID = 0
        ParentPath = "0"
        If ClassType = 1 Then
            ParentDir = "/"
        Else
            ParentDir = ""
        End If
    End If

    sql = "Select top 1 * from PE_Class where ChannelID=" & ChannelID & " order by ClassID desc"
    Set rsClass = Server.CreateObject("adodb.recordset")
    rsClass.Open sql, Conn, 1, 3
    rsClass.addnew
    rsClass("ChannelID") = ChannelID
    rsClass("ClassID") = ClassID
    rsClass("RootID") = RootID
    rsClass("ParentID") = ParentID
    If ParentID > 0 Then
        rsClass("Depth") = ParentDepth + 1
    Else
        rsClass("Depth") = 0
    End If
    rsClass("ParentPath") = ParentPath
    rsClass("OrderID") = PrevOrderID
    rsClass("Child") = 0
    rsClass("PrevID") = PrevID
    rsClass("NextID") = 0
    rsClass("arrChildID") = ClassID
    rsClass("ItemCount") = 0
    rsClass("ClassName") = ClassName
    rsClass("ClassType") = ClassType
    If ClassType > 1 Then
        rsClass("LinkUrl") = LinkUrl
        rsClass("ClassDir") = ""
        rsClass("ParentDir") = ""
    Else
        rsClass("LinkUrl") = ""
        rsClass("ClassDir") = ClassDir
        rsClass("ParentDir") = ParentDir
    End If
    rsClass("ClassPicUrl") = ClassPicUrl
    rsClass("Tips") = Tips
    rsClass("Readme") = ReadMe
    rsClass("Meta_Keywords") = Meta_Keywords
    rsClass("Meta_Description") = Meta_Description

    rsClass("ClassPurview") = ClassPurview
    rsClass("EnableComment") = EnableComment
    rsClass("CheckComment") = CheckComment

    rsClass("PresentExp") = PresentExp
    rsClass("DefaultItemPoint") = DefaultItemPoint
    rsClass("DefaultItemChargeType") = DefaultItemChargeType
    rsClass("DefaultItemPitchTime") = DefaultItemPitchTime
    rsClass("DefaultItemReadTimes") = DefaultItemReadTimes
    rsClass("DefaultItemDividePercent") = DefaultItemDividePercent
    
    rsClass("OpenType") = OpenType
    rsClass("ShowOnTop") = ShowOnTop
    rsClass("ShowOnIndex") = ShowOnIndex
    rsClass("IsElite") = IsElite
    rsClass("EnableAdd") = EnableAdd
    rsClass("EnableProtect") = EnableProtect
    rsClass("SkinID") = SkinID
    rsClass("TemplateID") = TemplateID
    rsClass("MaxPerPage") = MaxPerPage
    rsClass("DefaultItemSkin") = DefaultItemSkin
    rsClass("DefaultItemTemplate") = DefaultItemTemplate
    rsClass("ItemListOrderType") = ItemListOrderType
    rsClass("ItemOpenType") = ItemOpenType
    
    rsClass("CommandClassPoint") = CommandClassPoint
    rsClass("ReleaseClassPoint") = ReleaseClassPoint
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
    rsClass("Custom_Content") = Custom_Content
    
    rsClass.Update
    rsClass.Close
    Set rsClass = Nothing
    
    '更新与本栏目同一父栏目的上一个栏目的“NextID”字段值
    If PrevID > 0 Then
        Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & PrevID)
    End If
    
    If ParentID > 0 Then
        '更新其父类的子栏目数
        Conn.Execute ("update PE_Class set child=child+1 where ClassID=" & ParentID)
        
        '更新该栏目排序以及大于本需要和同在本分类下的栏目排序序号
        Conn.Execute ("update PE_Class set OrderID=OrderID+1 where ChannelID=" & ChannelID & " and RootID=" & RootID & " and OrderID>" & PrevOrderID)
        Conn.Execute ("update PE_Class set OrderID=" & PrevOrderID & "+1 where ClassID=" & ClassID)
    End If
    
    '从父路径中继承栏目权限并更新本栏目及其子栏目的权限
    Call UpdateClassPurview(ClassID)
    
    Call AddGroupPurview("Browse", arrGroupID_Browse, ClassID)
    Call AddGroupPurview("View", arrGroupID_View, ClassID)
    Call AddGroupPurview("Input", arrGroupID_Input, ClassID)

    Call CreateJS_Class
    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID
End Sub

Sub SaveModify()
    Dim ClassID, ClassName, ClassType, LinkUrl, ClassPicUrl, Tips, ReadMe, Meta_Keywords, Meta_Description
    Dim ClassPurview, arrGroupID_Browse, arrGroupID_View, arrGroupID_Input, EnableComment, CheckComment
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent
    Dim OpenType, ShowOnTop, ShowOnIndex, IsElite, EnableAdd, EnableProtect, SkinID, TemplateID
    Dim MaxPerPage, DefaultItemSkin, DefaultItemTemplate, ItemListOrderType, ItemOpenType
    Dim sql, rsClass, i, trs
    Dim ReleaseClassPoint, CommandClassPoint '在栏目下发布信息要扣除的会员点数和设置栏目推荐要扣除的会员点数
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If

    ClassName = Trim(Request("ClassName"))
    ClassType = PE_CLng(Trim(Request("ClassType")))
    LinkUrl = Trim(Request("LinkUrl"))
    ClassPicUrl = Trim(Request("ClassPicUrl"))
    Tips = Trim(Request("Tips"))
    ReadMe = Trim(Request("Readme"))
    Meta_Keywords = Trim(Request("Meta_Keywords"))
    Meta_Description = Trim(Request("Meta_Description"))
    
    ClassPurview = PE_CLng(Trim(Request("ClassPurview")))
    arrGroupID_Browse = ReplaceBadChar(Trim(Request("arrGroupID_Browse")))
    arrGroupID_View = ReplaceBadChar(Trim(Request("arrGroupID_View")))
    arrGroupID_Input = ReplaceBadChar(Trim(Request("arrGroupID_Input")))
    EnableComment = PE_CBool(Trim(Request("EnableComment")))
    CheckComment = PE_CBool(Trim(Request("CheckComment")))
    
    PresentExp = PE_CDbl(Trim(Request("PresentExp")))
    DefaultItemPoint = PE_CDbl(Trim(Request("DefaultItemPoint")))
    DefaultItemChargeType = PE_CLng(Trim(Request.Form("DefaultItemChargeType")))
    DefaultItemPitchTime = PE_CLng(Trim(Request.Form("DefaultItemPitchTime")))
    DefaultItemReadTimes = PE_CLng(Trim(Request.Form("DefaultItemReadTimes")))
    DefaultItemDividePercent = PE_CLng(Trim(Request.Form("DefaultItemDividePercent")))

    OpenType = PE_CLng(Trim(Request("OpenType")))
    ShowOnTop = PE_CBool(Trim(Request("ShowOnTop")))
    ShowOnIndex = PE_CBool(Trim(Request("ShowOnIndex")))
    IsElite = PE_CBool(Trim(Request("IsElite")))
    EnableAdd = PE_CBool(Trim(Request("EnableAdd")))
    EnableProtect = PE_CBool(Trim(Request("EnableProtect")))
    SkinID = PE_CLng(Trim(Request("SkinID")))
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    DefaultItemSkin = PE_CLng(Trim(Request("DefaultItemSkin")))
    DefaultItemTemplate = PE_CLng(Trim(Request("DefaultItemTemplate")))
    ItemListOrderType = PE_CLng(Trim(Request("ItemListOrderType")))
    ItemOpenType = PE_CLng(Trim(Request("ItemOpenType")))
    ReleaseClassPoint = PE_CLng(Trim(Request("ReleaseClassPoint")))
    CommandClassPoint = PE_CLng(Trim(Request("CommandClassPoint")))

    If ClassName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>栏目名称不能为空！</li>"
    Else
        ClassName = ReplaceBadChar(ClassName)
    End If
    If ClassType > 1 Then
        If LinkUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址不能为空！</li>"
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    
    sql = "select * from PE_Class where ClassID=" & ClassID
    Set rsClass = Server.CreateObject("Adodb.recordset")
    rsClass.Open sql, Conn, 1, 3
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    End If

    rsClass("ClassName") = ClassName
    rsClass("LinkUrl") = LinkUrl
    rsClass("ClassPicUrl") = ClassPicUrl
    rsClass("Tips") = Tips
    rsClass("Readme") = ReadMe
    rsClass("Meta_Keywords") = Meta_Keywords
    rsClass("Meta_Description") = Meta_Description
    rsClass("ClassPurview") = ClassPurview
    rsClass("EnableComment") = EnableComment
    rsClass("CheckComment") = CheckComment

    rsClass("PresentExp") = PresentExp
    rsClass("DefaultItemPoint") = DefaultItemPoint
    rsClass("DefaultItemChargeType") = DefaultItemChargeType
    rsClass("DefaultItemPitchTime") = DefaultItemPitchTime
    rsClass("DefaultItemReadTimes") = DefaultItemReadTimes
    rsClass("DefaultItemDividePercent") = DefaultItemDividePercent

    rsClass("OpenType") = OpenType
    rsClass("ShowOnTop") = ShowOnTop
    rsClass("ShowOnIndex") = ShowOnIndex
    rsClass("IsElite") = IsElite
    rsClass("EnableAdd") = EnableAdd
    rsClass("EnableProtect") = EnableProtect
    rsClass("SkinID") = SkinID
    rsClass("TemplateID") = TemplateID
    rsClass("MaxPerPage") = MaxPerPage
    rsClass("DefaultItemTemplate") = DefaultItemTemplate
    rsClass("DefaultItemSkin") = DefaultItemSkin
    rsClass("ItemListOrderType") = ItemListOrderType
    rsClass("ItemOpenType") = ItemOpenType
    rsClass("CommandClassPoint") = CommandClassPoint
    rsClass("ReleaseClassPoint") = ReleaseClassPoint

    '自设内容
    Dim Custom_Num, Custom_Content
    Custom_Num = PE_CLng(Request.Form("Custom_Num"))
    If Custom_Num <> 0 Then
        For i = 1 To Custom_Num
            If i <> 1 Then
                Custom_Content = Custom_Content & "{#$$$#}"
            End If
            Custom_Content = Custom_Content & Trim(Request("Custom_Content" & i))
        Next
    End If
    rsClass("Custom_Content") = Custom_Content

    rsClass.Update
    rsClass.Close
    Set rsClass = Nothing

    '从父路径中继承栏目权限并更新本栏目及其子栏目的权限
    Call UpdateClassPurview(ClassID)
    Call ModifyGroupPurview("Browse", arrGroupID_Browse, ClassID)
    Call ModifyGroupPurview("View", arrGroupID_View, ClassID)
    Call ModifyGroupPurview("Input", arrGroupID_Input, ClassID)

    If FoundErr = True Then Exit Sub

    If ClassType > 1 Then
        Call CreateJS_Class
    End If

    Call ClearSiteCache(0)
    If UseCreateHTML > 0 Then
        Call WriteSuccessMsg("修改栏目属性成功！记得重新生成相关文件哦！", ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID
    End If
End Sub


Sub DeleteClass()
    Dim sql, rsClass, trs, PrevID, NextID, ClassID, arrChildID, RootID, OrderID, strMsg, strListPath
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    
    sql = "select ClassID,RootID,Depth,ParentID,arrChildID,Child,PrevID,NextID,OrderID,ClassType,ParentDir,ParentPath,ClassDir from PE_Class where ClassID=" & ClassID
    Set rsClass = Conn.Execute(sql)
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>栏目不存在，或者已经被删除</li>"
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    End If
    PrevID = rsClass("PrevID")
    NextID = rsClass("NextID")
    arrChildID = rsClass("arrChildID")
    RootID = rsClass("RootID")
    OrderID = rsClass("OrderID")
    If rsClass("Depth") > 0 Then
        Conn.Execute ("update PE_Class set child=child-1 where ClassID=" & rsClass("ParentID"))

        '更新此栏目的原来所有上级栏目的子栏目ID数组
        Set trs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & rsClass("ParentPath") & ")")
        Do While Not trs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & RemoveClassID(trs(1), arrChildID) & "' where ClassID=" & trs(0))
            trs.MoveNext
        Loop
        trs.Close
        
        '更新与此栏目同根且排序在其之下的栏目
        Conn.Execute ("update PE_Class set OrderID=OrderID-" & UBound(Split(arrChildID, ",")) + 1 & " where ChannelID=" & ChannelID & " and RootID=" & RootID & " and OrderID>" & OrderID)

    End If
    
    '修改上一栏目的NextID和下一栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If
    
    
    If rsClass("ClassType") = 1 And UseCreateHTML > 0 Then
        '删除此栏目下的内容
        Select Case StructureType
        Case 0, 1, 2
            Call DelClassDir(HtmlDir & rsClass("ParentDir") & rsClass("ClassDir"))
        Case 3, 4, 5
            Call DelClassDir(HtmlDir & "/" & rsClass("ClassDir"))
        Case Else
            Call DelInfo(arrChildID)
        End Select
        '删除此栏目列表页
        If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
            strListPath = HtmlDir & GetListPath(StructureType, ListFileType, rsClass("ParentDir"), rsClass("ClassDir")) & GetListFileName(ListFileType, rsClass("ClassID"), 1, 1)
            If fso.FileExists(Server.MapPath(strListPath & FileExt_List)) Then
                fso.DeleteFile Server.MapPath(strListPath & FileExt_List)
                DelSerialFiles (Server.MapPath(strListPath) & "_*" & FileExt_Item)
            End If
        End If

    End If
    rsClass.Close
    Set rsClass = Nothing
    
    '删除本栏目的所有内容的相关评论
    Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (select " & ModuleName & "ID from " & SheetName & " where ClassID in (" & arrChildID & "))")
    '删除本栏目（包括子栏目）
    Conn.Execute ("delete from PE_Class where ChannelID=" & ChannelID & " and ClassID in (" & arrChildID & ")")
    '删除本栏目（包括子栏目）的所有内容和评论
    Conn.Execute ("delete from " & SheetName & " where ChannelID=" & ChannelID & " and ClassID in (" & arrChildID & ")")
    
    Call UpdateChannelData(ChannelID)
    Call CreateJS_Class
    Call ClearSiteCache(0)
    If FoundErr <> True Then
        If UseCreateHTML > 0 Then
            strMsg = strMsg & "删除栏目成功！请记得重新生成相关栏目的文件呀！"
            Call WriteSuccessMsg(strMsg, ComeUrl)
        Else
            Call CloseConn
            Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID
        End If
    End If
End Sub

Sub DelClassDir(DirName)
    On Error Resume Next
    If ObjInstalled_FSO = False Or Trim(DirName) = "" Then Exit Sub
    If fso.FolderExists(Server.MapPath(DirName)) Then
        fso.DeleteFolder Server.MapPath(DirName)
        If Err Then
            Err.Clear
            FoundErr = True
            ErrMsg = ErrMsg & "<li>栏目目录无法自动删除！可能此目录中的文件正在使用中！请稍后使用FTP手动删除此目录。</li>"
        End If
    End If
End Sub

Sub ClearClass()
    Dim rsClass, SuccessMsg, ClassID
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    Set rsClass = Conn.Execute("select arrChildID,ParentDir,ClassDir,ClassType from PE_Class where ClassID=" & ClassID)
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>栏目不存在，或者已经被删除</li>"
    Else
        Conn.Execute ("update " & SheetName & " set Deleted=" & PE_True & " where ClassID in (" & rsClass(0) & ")")
        SuccessMsg = "此栏目（包括子栏目）的所有" & ChannelShortName & "已经被移到回收站中！"
        If rsClass(3) = 1 And UseCreateHTML > 0 Then
            Select Case StructureType
            Case 0, 1, 2
                Call ClearDir(HtmlDir & rsClass(1) & rsClass(2))
            Case 3, 4, 5
                Call ClearDir(HtmlDir & "/" & rsClass(2))
            Case Else
                Call DelInfo(rsClass(0))
            End Select
        End If
    End If
    rsClass.Close
    Set rsClass = Nothing
    
    If FoundErr = True Then Exit Sub
    
    Call UpdateChannelData(ChannelID)
    Call ClearSiteCache(0)
    
    If UseCreateHTML > 0 Then
        SuccessMsg = SuccessMsg & "<br>本栏目（包括子栏目）下的所有HTML文件已经被删除！你需要重新生成相关文件。"
        Call WriteSuccessMsg(SuccessMsg, ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID
    End If
End Sub

Sub ClearDir(DirName)
    On Error Resume Next
    Dim tmpDir, theFolder, theSubFolder
    tmpDir = Server.MapPath(DirName)
    If Not fso.FolderExists(tmpDir) Then
        Exit Sub
    End If
    fso.DeleteFile tmpDir & "/*.*"
    Set theFolder = fso.GetFolder(tmpDir)
    For Each theSubFolder In theFolder.SubFolders
        fso.DeleteFile tmpDir & "/" & theSubFolder.Name & "/*.*"
    Next
End Sub

Sub SaveMove()
    Dim tChannelID, ClassID, sql, rsClass, i, rsPrevOrderID
    Dim rParentID
    Dim trs, rs, strMsg
    Dim ParentID, RootID, Depth, Child, ParentPath, ParentName, iParentPath, PrevOrderID, PrevID, NextID, ClassCount
    Dim ClassName, ClassType, ParentDir, tParentDir, cParentDir, arrChildID, ClassDir, CurrentDir, TargetDir
    tChannelID = PE_CLng(Trim(Request("tChannelID")))
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    
    sql = "select * from PE_Class where ClassID=" & ClassID
    Set rsClass = Server.CreateObject("Adodb.recordset")
    rsClass.Open sql, Conn, 1, 3
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
    Else
        Depth = rsClass("Depth")
        Child = rsClass("Child")
        RootID = rsClass("RootID")
        ParentID = rsClass("ParentID")
        ParentPath = rsClass("ParentPath")
        PrevID = rsClass("PrevID")
        NextID = rsClass("NextID")
        ClassName = rsClass("ClassName")
        arrChildID = rsClass("arrChildID")
        ParentDir = rsClass("ParentDir")
        ClassDir = rsClass("ClassDir")
        ClassType = rsClass("ClassType")
    End If
    rsClass.Close
    Set rsClass = Nothing
    


    rParentID = PE_CLng(Trim(Request("ParentID")))
    If tChannelID = ChannelID Then
        If rParentID = ClassID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>所属栏目不能为自己！</li>"
        Else
            If rParentID = ParentID Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>目标栏目与当前父栏目相同，无需移动！</li>"
            End If
        End If
    End If

    If FoundErr = True Then Exit Sub
    
    If rParentID > 0 Then
        Set trs = Conn.Execute("select ClassID from PE_Class where ChannelID=" & tChannelID & " and ClassType=1 and ClassID=" & rParentID)
        If trs.BOF And trs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能指定外部栏目为所属栏目</li>"
        End If
        trs.Close
        Set trs = Nothing
        If FoundInArr(arrChildID, rParentID, ",") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能指定该栏目的下属栏目作为所属栏目</li>"
        End If
    End If

    '检查目标栏目的子栏目中是否已经存在与此栏目名称相同的栏目
    Set trs = Conn.Execute("select ClassID,ClassDir from PE_Class where ChannelID=" & tChannelID & " and ParentID=" & rParentID & " and ClassName='" & ClassName & "'")
    If Not (trs.BOF And trs.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>目标栏目的子栏目中已经存在与此栏目名称相同的栏目。"
    End If
    Set trs = Nothing

   If StructureType <= 1 Then
        '检查目标栏目的子栏目中是否已经存在与此栏目目录相同的栏目
        If ClassType = 1 Then
            Set trs = Conn.Execute("select ClassID,ParentDir from PE_Class where ChannelID=" & tChannelID & " and ParentID=" & rParentID & " and ClassDir='" & ClassDir & "'")
            If Not (trs.BOF And trs.EOF) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>目标栏目的子栏目中已经存在与此栏目目录相同的栏目。"
            End If
            Set trs = Nothing
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    ClassCount = UBound(Split(arrChildID, ",")) + 1    '得到要移动的栏目数
    CurrentDir = HtmlDir & ParentDir & ClassDir '得到当前目录
    
    '需要更新其原来所属栏目信息，包括深度、父级ID、栏目数、排序等数据
    '需要更新当前所属栏目信息
    Dim mrs, MaxRootID
    Set mrs = Conn.Execute("select max(rootid) from PE_Class where ChannelID=" & tChannelID & "")
    MaxRootID = mrs(0)
    Set mrs = Nothing
    If IsNull(MaxRootID) Then
        MaxRootID = 0
    End If

    If UseCreateHTML > 0 And StructureType <= 1 And ClassType = 1 And ObjInstalled_FSO = True Then
        If rParentID = 0 Then
            TargetDir = HtmlDir & "/"
        Else
            Set trs = Conn.Execute("select ParentDir,ClassDir from PE_Class where ClassID=" & rParentID)
            TargetDir = HtmlDir & trs("ParentDir") & trs("ClassDir") & "/"
            Set trs = Nothing
        End If
        
        If fso.FolderExists(Server.MapPath(TargetDir & ClassDir)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>目录栏目中已经存在与此栏目目录相同的子栏目。这可能是因为此栏目的目录名为“JS”、“UploadFiles”等系统栏目名造成的。</li>"
            Exit Sub
        End If
        If FoundErr = True Then Exit Sub

        If fso.FolderExists(Server.MapPath(CurrentDir)) Then
            Call DelClassDir(CurrentDir)
        End If
    End If

    '更新原来同一父栏目的上一个栏目的NextID和下一个栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If

    If ParentID = 0 And rParentID = 0 Then  '如果原来是一级分类跨频道移到另一频道一级分类
        '得到上一个一级分类栏目
        sql = "select ClassID,NextID from PE_Class where ChannelID=" & tChannelID & " and RootID=" & MaxRootID & " and Depth=0"
        Set rs = Server.CreateObject("Adodb.recordset")
        rs.Open sql, Conn, 1, 3
        If rs.BOF And rs.EOF Then
            PrevID = 0
        Else
            PrevID = rs(0)    '得到新的PrevID
            rs(1) = ClassID   '更新上一个一级分类栏目的NextID的值
            rs.Update
        End If
        rs.Close
        Set rs = Nothing

        MaxRootID = MaxRootID + 1

        '更新当前栏目数据
        Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",rootid=" & MaxRootID & ",PrevID=" & PrevID & ",NextID=0 where ClassID=" & ClassID)
        
        '如果有下属栏目，则更新其下属栏目数据。下属栏目的排序不需考虑，只需更新下属栏目深度和一级排序ID(rootid)数据
        If Child > 0 Then
            Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",rootid=" & MaxRootID & " where ClassID in (" & arrChildID & ")")
        End If

    ElseIf ParentID > 0 And rParentID = 0 Then  '如果原来不是一级分类改成一级分类

        '更新其原来所属栏目的栏目数，排序相当于剪枝而不需考虑
        Conn.Execute ("update PE_Class set child=child-1 where ClassID=" & ParentID)

        '更新此栏目的原来所有上级栏目的子栏目ID数组
        Set trs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & ParentPath & ")")
        Do While Not trs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & RemoveClassID(trs(1), arrChildID) & "' where ClassID=" & trs(0))
            trs.MoveNext
        Loop
        trs.Close

        '得到上一个一级分类栏目
        sql = "select ClassID,NextID from PE_Class where ChannelID=" & tChannelID & " and RootID=" & MaxRootID & " and Depth=0"
        Set rs = Server.CreateObject("Adodb.recordset")
        rs.Open sql, Conn, 1, 3
        If rs.BOF And rs.EOF Then
            PrevID = 0
        Else
            PrevID = rs(0)    '得到新的PrevID
            rs(1) = ClassID   '更新上一个一级分类栏目的NextID的值
            rs.Update
        End If
        rs.Close
        Set rs = Nothing

        MaxRootID = MaxRootID + 1

        tParentDir = "/"
        '更新当前栏目数据
        Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",depth=0,OrderID=0,rootid=" & MaxRootID & ",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0,ParentDir='" & tParentDir & "' where ClassID=" & ClassID)

        '如果有下属栏目，则更新其下属栏目数据。下属栏目的排序不需考虑，只需更新下属栏目深度和一级排序ID(rootid)数据
        If Child > 0 Then
            ParentPath = ParentPath & ","
            arrChildID = RemoveClassID(arrChildID, ClassID) '从子栏目数组中去掉当前栏目的ID
            Set rs = Conn.Execute("select * from PE_Class where ClassID in (" & arrChildID & ")")
            Do While Not rs.EOF
                iParentPath = Replace(rs("ParentPath"), ParentPath, "")
                cParentDir = tParentDir & Right(rs("ParentDir"), Len(rs("ParentDir")) - Len(ParentDir))
                Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",depth=depth-" & Depth & ",rootid=" & MaxRootID & ",ParentPath='0," & iParentPath & "',ParentDir='" & cParentDir & "' where ClassID=" & rs("ClassID"))
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
        
    ElseIf ParentID > 0 And rParentID > 0 Then '如果是将一个分栏目移动到其他分栏目下

        '更新其原父类的子栏目数
        Conn.Execute ("update PE_Class set child=child-1 where ClassID=" & ParentID)

        '更新此栏目的原来所有上级栏目的子栏目ID数组
        Set trs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & ParentPath & ")")
        Do While Not trs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & RemoveClassID(trs(1), arrChildID) & "' where ClassID=" & trs(0))
            trs.MoveNext
        Loop
        trs.Close

        '获得目标栏目的相关信息
        Set trs = Conn.Execute("select * from PE_Class where ClassID=" & rParentID)
        tParentDir = trs("ParentDir") & trs("ClassDir") & "/"

        If trs("Child") > 0 Then
            '得到在目标栏目中与本栏目同级的最后一个栏目的ClassID，并更新其NextID的指向
            Set rs = Conn.Execute("select ClassID from PE_Class where ParentID=" & trs("ClassID") & " order by OrderID desc")
            PrevID = rs(0)  '得到新的PrevID
            Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & rs(0) & "")
            Set rs = Nothing

            '得到目标栏目的子栏目的最大OrderID
            Set rsPrevOrderID = Conn.Execute("select Max(OrderID) from PE_Class where ClassID in (" & trs("arrChildID") & ")")
            PrevOrderID = rsPrevOrderID(0)
            Set rsPrevOrderID = Nothing
        Else
            PrevID = 0
            PrevOrderID = trs("OrderID")
        End If

        '更新目标栏目的子栏目数
        Conn.Execute ("update PE_Class set child=child+1 where ClassID=" & rParentID)

        '更新目标栏目及目标栏目的所有上级栏目的子栏目ID数组
        Set rs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & trs("ParentPath") & "," & trs("ClassID") & ")")
        Do While Not rs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & rs(1) & "," & arrChildID & "' where ClassID=" & rs(0))
            rs.MoveNext
        Loop
        rs.Close


        '在获得移动过来的栏目数后更新排序在指定栏目之后的栏目排序数据
        Conn.Execute ("update PE_Class set OrderID=OrderID+" & ClassCount & "+1 where ChannelID=" & tChannelID & " and rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
        
        '更新当前栏目数据
        Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",depth=" & trs("depth") & "+1,OrderID=" & PrevOrderID & "+1,rootid=" & trs("rootid") & ",ParentID=" & rParentID & ",ParentPath='" & trs("ParentPath") & "," & trs("ClassID") & "',PrevID=" & PrevID & ",NextID=0,ParentDir='" & tParentDir & "' where ClassID=" & ClassID)

        '如果当前栏目有子栏目则更新子栏目数据，深度为原来的相对深度加上当前所属栏目的深度
        If Child > 0 Then
            i = 1
            arrChildID = RemoveClassID(arrChildID, ClassID) '从子栏目数组中去掉当前栏目的ID
            ParentPath = ParentPath & ","
            Set rs = Conn.Execute("select * from PE_Class where ClassID in (" & arrChildID & ") order by OrderID")
            Do While Not rs.EOF
                i = i + 1
                iParentPath = trs("ParentPath") & "," & trs("ClassID") & "," & Replace(rs("ParentPath"), ParentPath, "")
                cParentDir = tParentDir & Right(rs("ParentDir"), Len(rs("ParentDir")) - Len(ParentDir))
                Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",depth=depth-" & Depth & "+" & trs("depth") & "+1,OrderID=" & PrevOrderID & "+" & i & ",rootid=" & trs("rootid") & ",ParentPath='" & iParentPath & "',ParentDir='" & cParentDir & "' where ClassID=" & rs("ClassID"))
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
        trs.Close
        Set trs = Nothing
        
        
    Else    '如果原来是一级栏目改成其他栏目的下属栏目
        '获得目标栏目的相关信息
        Set trs = Conn.Execute("select * from PE_Class where ClassID=" & rParentID)
        tParentDir = trs("ParentDir") & trs("ClassDir") & "/"

        If trs("Child") > 0 Then
            '得到在目标栏目中与本栏目同级的最后一个栏目的ClassID，并更新其NextID的指向
            Set rs = Conn.Execute("select ClassID from PE_Class where ParentID=" & trs("ClassID") & " order by OrderID desc")
            PrevID = rs(0)  '得到新的PrevID
            Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & rs(0) & "")
            Set rs = Nothing

            '得到目标栏目的子栏目的最大OrderID
            Set rsPrevOrderID = Conn.Execute("select Max(OrderID) from PE_Class where ClassID in (" & trs("arrChildID") & ")")
            PrevOrderID = rsPrevOrderID(0)
            Set rsPrevOrderID = Nothing
        Else
            PrevID = 0
            PrevOrderID = trs("OrderID")
        End If

        '更新目标栏目的子栏目数
        Conn.Execute ("update PE_Class set child=child+1 where ClassID=" & rParentID)

        '更新目标栏目及目标栏目的所有上级栏目的子栏目ID数组
        Set rs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & trs("ParentPath") & "," & trs("ClassID") & ")")
        Do While Not rs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & rs(1) & "," & arrChildID & "' where ClassID=" & rs(0))
            rs.MoveNext
        Loop
        rs.Close
    
        '在获得移动过来的栏目数后更新排序在指定栏目之后的栏目排序数据
        Conn.Execute ("update PE_Class set OrderID=OrderID+" & ClassCount & "+1 where ChannelID=" & tChannelID & " and rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
        
        '更新当前栏目数据
        Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",depth=depth+" & trs("depth") & "+1,OrderID=" & PrevOrderID + 1 & ",rootid=" & trs("rootid") & ",ParentPath='" & trs("ParentPath") & "," & trs("ClassID") & "',parentid=" & rParentID & ", PrevID=" & PrevID & ",NextID=0,ParentDir='" & tParentDir & "' where ClassID=" & ClassID & "")

        '如果当前栏目有子栏目则更新子栏目数据，深度为原来的相对深度加上当前所属栏目的深度
        Set rs = Conn.Execute("select * from PE_Class where ChannelID=" & ChannelID & " and rootid=" & RootID & " and ParentID>0 order by OrderID")
        i = 1
        Do While Not rs.EOF
            i = i + 1
            iParentPath = trs("ParentPath") & "," & trs("ClassID") & "," & Replace(rs("ParentPath"), "0,", "")
            cParentDir = tParentDir & Right(rs("ParentDir"), Len(rs("ParentDir")) - Len(ParentDir))
            Conn.Execute ("update PE_Class set ChannelID=" & tChannelID & ",depth=depth+" & trs("depth") & "+1,OrderID=" & PrevOrderID & "+" & i & ",rootid=" & trs("rootid") & ",ParentPath='" & iParentPath & "',ParentDir='" & cParentDir & "' where ClassID=" & rs("ClassID"))
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        trs.Close
        Set trs = Nothing
    End If
    
    If tChannelID <> ChannelID Then '如果是跨频道移动栏目
        Conn.Execute ("update " & SheetName & " set ChannelID=" & tChannelID & " where ChannelID=" & ChannelID & " and ClassID in (" & arrChildID & ")")
        Call MoveUpFilesToOtherChannel(tChannelID, arrChildID)
    End If
    '从父路径中继承栏目权限并更新本栏目及其子栏目的权限
    Call UpdateClassPurview(ClassID)
    
    Call CreateJS_Class
    Call ClearSiteCache(0)
    If UseCreateHTML > 0 Then
        Call WriteSuccessMsg("移动栏目成功！请记到重新生成相关文件！", ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID
    End If
End Sub

Sub MoveUpFilesToOtherChannel(tChannelID, tClassID)
    Dim rsBatchMove, sqlBatchMove, ArticlePath
    Dim rsChannel, tChannelDir, tUploadDir
    Set rsChannel = Conn.Execute("select ChannelDir,UploadDir from PE_Channel where ChannelID=" & tChannelID & "")
    If rsChannel.BOF And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到目标频道！</li>"
    Else
        tChannelDir = rsChannel("ChannelDir")
        tUploadDir = rsChannel("UploadDir")
    End If
    Set rsChannel = Nothing
    If FoundErr = True Then Exit Sub
    Select Case ModuleType
    Case 1
        sqlBatchMove = "select UploadFiles from PE_Article where ClassID in (" & tClassID & ")"
    Case 2
        sqlBatchMove = "select SoftPicUrl,DownloadUrl from PE_Soft where ClassID in (" & tClassID & ")"
    Case 3
        sqlBatchMove = "select PhotoThumb,PhotoUrl from PE_Photo where ClassID in (" & tClassID & ")"
    End Select
    Set rsBatchMove = Conn.Execute(sqlBatchMove)
    Do While Not rsBatchMove.EOF
        Select Case ModuleType
        Case 1
            Call MoveUpFiles(rsBatchMove("UploadFiles") & "", tChannelDir & "/" & tUploadDir)    '移动上传文件
        Case 2
            Call MoveUpPic(rsBatchMove("SoftPicUrl"), tChannelDir)
            Call MoveSoftUpFiles(rsBatchMove("DownloadUrl"), tChannelDir & "/" & tUploadDir)    '移动上传文件
        Case 3
            Call MovePhotoUpFiles("缩略图|" & rsBatchMove("PhotoThumb") & "$$$" & rsBatchMove("PhotoUrl"), tChannelDir & "/" & tUploadDir)    '移动上传文件
        End Select
        rsBatchMove.MoveNext
    Loop
    rsBatchMove.Close
    Set rsBatchMove = Nothing
End Sub


Sub MoveUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim strTrueFile, arrFiles, strTrueDir, i
    If IsNull(strFiles) Or strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    arrFiles = Split(strFiles, "|")
    For i = 0 To UBound(arrFiles)
        strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(arrFiles(i), InStr(arrFiles(i), "/")))
        If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
        strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & arrFiles(i))
        If fso.FileExists(strTrueFile) Then
            fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & arrFiles(i))
        End If
    Next
End Sub

Sub MoveSoftUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim arrSoftUrls, strTrueFile, arrUrls, strTrueDir, iTemp
    If strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    arrSoftUrls = Split(strFiles, "$$$")
    For iTemp = 0 To UBound(arrSoftUrls)
        arrUrls = Split(arrSoftUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(arrUrls(1), InStr(arrUrls(1), "/")))
                If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
                strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1))
                If fso.FileExists(strTrueFile) Then
                    fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & arrUrls(1))
                End If
            End If
        End If
    Next
    
End Sub

Sub MoveUpPic(strFile, strTargetDir)
    On Error Resume Next
    Dim strTrueFile, strTrueDir
    If strFile = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    If Left(strFile, 1) <> "/" And InStr(strFile, "://") <= 0 Then
        strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(strFile, InStrRev(strFile, "/")))
        If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
        strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & strFile)
        If fso.FileExists(strTrueFile) Then
            fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & strFile)
        End If
    End If
End Sub

Sub MovePhotoUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim arrPhotoUrls, strTrueFile, arrUrls, strTrueDir, iTemp
    If strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    arrPhotoUrls = Split(strFiles, "$$$")
    For iTemp = 0 To UBound(arrPhotoUrls)
        arrUrls = Split(arrPhotoUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(arrUrls(1), InStr(arrUrls(1), "/")))
                If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
                strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1))
                If fso.FileExists(strTrueFile) Then
                    fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & arrUrls(1))
                End If
            End If
        End If
    Next
    
End Sub

Sub UpOrder()
    Dim ClassID, sqlOrder, rsOrder, MoveNum, cRootID, i, rs, PrevID, NextID
    ClassID = Trim(Request("ClassID"))
    cRootID = Trim(Request("cRootID"))
    MoveNum = Trim(Request("MoveNum"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If cRootID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cRootID = PE_CLng(cRootID)
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

    Dim mrs, MaxRootID, tRootID, tClassID, tOrderID, tPrevID
    
    '得到本栏目的PrevID,NextID
    Set rs = Conn.Execute("select PrevID,NextID from PE_Class where ClassID=" & ClassID)
    PrevID = rs(0)
    NextID = rs(1)
    rs.Close
    Set rs = Nothing
    '先修改上一栏目的NextID和下一栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If

    '得到本频道最大RootID值
    Set mrs = Conn.Execute("select max(rootid) from PE_Class where ChannelID=" & ChannelID & "")
    MaxRootID = mrs(0) + 1
    '先将当前栏目移至最后，包括子栏目
    Conn.Execute ("update PE_Class set RootID=" & MaxRootID & " where ChannelID=" & ChannelID & " and RootID=" & cRootID)
    
    '然后将位于当前栏目以上的栏目的RootID依次加一，范围为要提升的数字
    sqlOrder = "select * from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and RootID<" & cRootID & " order by RootID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前栏目已经在最上面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tRootID = rsOrder("RootID")     '得到要提升位置的RootID，包括子栏目
        Conn.Execute ("update PE_Class set RootID=RootID+1 where ChannelID=" & ChannelID & " and RootID=" & tRootID)
        i = i + 1
        If i > MoveNum Then
            tClassID = rsOrder("ClassID")
            tPrevID = rsOrder("PrevID")
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
        
    '更新移动后本栏目的的PrevID和NextID，以及上一栏目的NextID和下一栏目的PrevID
    Conn.Execute ("update PE_Class set PrevID=" & tPrevID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set NextID=" & tClassID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set PrevID=" & ClassID & " where ClassID=" & tClassID)
    If tPrevID > 0 Then
        Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & tPrevID)
    End If
    
    '然后再将当前栏目从最后移到相应位置，包括子栏目
    Conn.Execute ("update PE_Class set RootID=" & tRootID & " where ChannelID=" & ChannelID & " and RootID=" & MaxRootID)
    Call CreateJS_Class
    Call ClearSiteCache(0)
    If UseCreateHTML > 0 Then
        Call WriteSuccessMsg("上移栏目成功！请记到重新生成相关文件！", ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Order"
    End If
End Sub

Sub DownOrder()
    Dim ClassID, sqlOrder, rsOrder, MoveNum, cRootID, i, rs, PrevID, NextID
    ClassID = Trim(Request("ClassID"))
    cRootID = Trim(Request("cRootID"))
    MoveNum = Trim(Request("MoveNum"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If cRootID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        cRootID = PE_CLng(cRootID)
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

    Dim mrs, MaxRootID, tRootID, tClassID, tOrderID, tNextID
    
    '得到本栏目的PrevID,NextID
    Set rs = Conn.Execute("select PrevID,NextID from PE_Class where ClassID=" & ClassID)
    PrevID = rs(0)
    NextID = rs(1)
    rs.Close
    Set rs = Nothing
    '先修改上一栏目的NextID和下一栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If

    '得到本频道最大RootID值
    Set mrs = Conn.Execute("select max(rootid) from PE_Class where ChannelID=" & ChannelID & "")
    MaxRootID = mrs(0) + 1
    '先将当前栏目移至最后，包括子栏目
    Conn.Execute ("update PE_Class set RootID=" & MaxRootID & " where ChannelID=" & ChannelID & " and RootID=" & cRootID)
    
    '然后将位于当前栏目以下的栏目的RootID依次减一，范围为要下降的数字
    sqlOrder = "select * from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and RootID>" & cRootID & " order by RootID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前栏目已经在最下面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tRootID = rsOrder("RootID")     '得到要提升位置的RootID，包括子栏目
        Conn.Execute ("update PE_Class set RootID=RootID-1 where ChannelID=" & ChannelID & " and RootID=" & tRootID)
        i = i + 1
        If i > MoveNum Then
            tClassID = rsOrder("ClassID")
            tNextID = rsOrder("NextID")
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '更新移动后本栏目的的PrevID和NextID，以及上一栏目的NextID和下一栏目的PrevID
    Conn.Execute ("update PE_Class set PrevID=" & tClassID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set NextID=" & tNextID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & tClassID)
    If tNextID > 0 Then
        Conn.Execute ("update PE_Class set PrevID=" & ClassID & " where ClassID=" & tNextID)
    End If
    
    '然后再将当前栏目从最后移到相应位置，包括子栏目
    Conn.Execute ("update PE_Class set RootID=" & tRootID & " where ChannelID=" & ChannelID & " and RootID=" & MaxRootID)
    Call CreateJS_Class
    Call ClearSiteCache(0)
    If UseCreateHTML > 0 Then
        Call WriteSuccessMsg("下移栏目成功！请记到重新生成相关文件！", ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID & "&Action=Order"
    End If
End Sub

Sub UpOrderN()
    Dim sqlOrder, rsOrder, MoveNum, ClassID, i
    Dim ParentID, OrderID, ParentPath, Child, PrevID, NextID
    ClassID = Trim(Request("ClassID"))
    MoveNum = Trim(Request("MoveNum"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
    Else
        ClassID = PE_CLng(ClassID)
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

    Dim sql, rs, trs, AddOrderNum, tClassID, tOrderID, tPrevID
    
    '要移动的栏目信息
    Set rs = Conn.Execute("select ParentID,OrderID,ParentPath,Child,PrevID,NextID from PE_Class where ClassID=" & ClassID)
    ParentID = rs(0)
    OrderID = rs(1)
    ParentPath = rs(2) & "," & ClassID
    Child = rs(3)
    PrevID = rs(4)
    NextID = rs(5)
    rs.Close
    Set rs = Nothing
    
    '获得要移动的栏目的所有子栏目数，然后加1（栏目本身），得到排序增加数（即其上栏目的OrderID增加数AddOrderNum）
    If Child > 0 Then
        Set rs = Conn.Execute("select count(*) from PE_Class where ParentPath like '%" & ParentPath & "%'")
        AddOrderNum = rs(0) + 1
        rs.Close
        Set rs = Nothing
    Else
        AddOrderNum = 1
    End If
    
    '先修改上一栏目的NextID和下一栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If
    
    '和该栏目同级且排序在其之上的栏目------更新其排序，范围为要提升的数字AddOrderNum
    sql = "select ClassID,OrderID,Child,ParentPath,PrevID,NextID from PE_Class where ParentID=" & ParentID & " and OrderID<" & OrderID & " order by OrderID desc"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 3
    i = 0
    Do While Not rs.EOF
        tOrderID = rs(1)
        Conn.Execute ("update PE_Class set OrderID=OrderID+" & AddOrderNum & " where ClassID=" & rs(0))
        If rs(2) > 0 Then
            Set trs = Conn.Execute("select ClassID,OrderID from PE_Class where ParentPath like '%" & rs(3) & "," & rs(0) & "%' order by OrderID")
            If Not (trs.BOF And trs.EOF) Then
                Do While Not trs.EOF
                    Conn.Execute ("update PE_Class set OrderID=OrderID+" & AddOrderNum & " where ClassID=" & trs(0))
                    trs.MoveNext
                Loop
            End If
            trs.Close
            Set trs = Nothing
        End If
        i = i + 1
        If i >= MoveNum Then
            '获得最后一个提升序号的同级栏目信息
            tClassID = rs(0)
            tPrevID = rs(4)
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    '更新移动后本栏目的的PrevID和NextID，以及上一栏目的NextID和下一栏目的PrevID
    Conn.Execute ("update PE_Class set PrevID=" & tPrevID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set NextID=" & tClassID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set PrevID=" & ClassID & " where ClassID=" & tClassID)
    If tPrevID > 0 Then
        Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & tPrevID)
    End If
        
    '更新所要排序的栏目的序号
    Conn.Execute ("update PE_Class set OrderID=" & tOrderID & " where ClassID=" & ClassID)
    '如果有下属栏目，则更新其下属栏目排序
    If Child > 0 Then
        i = 1
        Set rs = Conn.Execute("select ClassID from PE_Class where ParentPath like '%" & ParentPath & "%' order by OrderID")
        Do While Not rs.EOF
            Conn.Execute ("update PE_Class set OrderID=" & tOrderID + i & " where ClassID=" & rs(0))
            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    
    
    Call CreateJS_Class
    Call ClearSiteCache(0)
    If UseCreateHTML > 0 Then
        Call WriteSuccessMsg("上移栏目成功！请记到重新生成相关文件！", ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID & "&Action=OrderN"
    End If
End Sub

Sub DownOrderN()
    Dim sqlOrder, rsOrder, MoveNum, ClassID, i
    Dim ParentID, OrderID, ParentPath, Child, PrevID, NextID
    ClassID = Trim(Request("ClassID"))
    MoveNum = Trim(Request("MoveNum"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If MoveNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>错误参数！</li>"
        Exit Sub
    Else
        MoveNum = PE_CLng(MoveNum)
        If MoveNum = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请选择要下降的数字！</li>"
            Exit Sub
        End If
    End If

    Dim sql, rs, trs, ii, tClassID, tNextID
    
    '要移动的栏目信息
    Set rs = Conn.Execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID from PE_Class where ClassID=" & ClassID)
    ParentID = rs(0)
    OrderID = rs(1)
    ParentPath = rs(2) & "," & ClassID
    Child = rs(3)
    PrevID = rs(4)
    NextID = rs(5)
    rs.Close
    Set rs = Nothing

    '先修改上一栏目的NextID和下一栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If
    
    '和该栏目同级且排序在其之下的栏目------更新其排序，范围为要下降的数字
    sql = "select ClassID,OrderID,child,ParentPath,PrevID,NextID from PE_Class where ParentID=" & ParentID & " and OrderID>" & OrderID & " order by OrderID"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 3
    i = 0    '同级栏目
    ii = 0   '同级栏目和子栏目
    Do While Not rs.EOF
        Conn.Execute ("update PE_Class set OrderID=" & OrderID + ii & " where ClassID=" & rs(0))
        If rs(2) > 0 Then
            Set trs = Conn.Execute("select ClassID,OrderID from PE_Class where ParentPath like '%" & rs(3) & "," & rs(0) & "%' order by OrderID")
            If Not (trs.BOF And trs.EOF) Then
                Do While Not trs.EOF
                    ii = ii + 1
                    Conn.Execute ("update PE_Class set OrderID=" & OrderID + ii & " where ClassID=" & trs(0))
                    trs.MoveNext
                Loop
            End If
            trs.Close
            Set trs = Nothing
        End If
        ii = ii + 1
        i = i + 1
        If i >= MoveNum Then
            '获得移动后本栏目的上一栏目的信息
            tClassID = rs(0)
            tNextID = rs(5)
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
            
    '更新移动后本栏目的的PrevID和NextID，以及上一栏目的NextID和下一栏目的PrevID
    Conn.Execute ("update PE_Class set PrevID=" & tClassID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set NextID=" & tNextID & " where ClassID=" & ClassID)
    Conn.Execute ("update PE_Class set NextID=" & ClassID & " where ClassID=" & tClassID)
    If tNextID > 0 Then
        Conn.Execute ("update PE_Class set PrevID=" & ClassID & " where ClassID=" & tNextID)
    End If
    
    '更新所要排序的栏目的序号
    Conn.Execute ("update PE_Class set OrderID=" & OrderID + ii & " where ClassID=" & ClassID)
    '如果有下属栏目，则更新其下属栏目排序
    If Child > 0 Then
        i = 1
        Set rs = Conn.Execute("select ClassID from PE_Class where ParentPath like '%" & ParentPath & "%' order by OrderID")
        Do While Not rs.EOF
            Conn.Execute ("update PE_Class set OrderID=" & OrderID + ii + i & " where ClassID=" & rs(0))
            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    
    Call CreateJS_Class
    Call ClearSiteCache(0)
    If UseCreateHTML > 0 Then
        Call WriteSuccessMsg("上移栏目成功！请记到重新生成相关文件！", ComeUrl)
    Else
        Call CloseConn
        Response.Redirect "Admin_Class.asp?ChannelID=" & ChannelID & "&Action=OrderN"
    End If
End Sub

Sub SaveReset()
    Dim i, sql, rsClass, SuccessMsg, iCount, PrevID, NextID, ClassDir, trs
    sql = "select ClassID,ParentID,ClassType,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Server.CreateObject("adodb.recordset")
    rsClass.Open sql, Conn, 1, 1
    iCount = rsClass.RecordCount
    i = 1
    PrevID = 0
    Do While Not rsClass.EOF
        rsClass.MoveNext
        If rsClass.EOF Then
            NextID = 0
        Else
            NextID = rsClass(0)
        End If
        rsClass.moveprevious
        Set trs = Conn.Execute("select Count(ClassID) from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and ClassID<>" & rsClass(0) & " and ClassDir='" & rsClass(4) & "'")
        If trs(0) > 1 Then
            ClassDir = rsClass(4) & rsClass(0)
        Else
            ClassDir = rsClass(4)
        End If
        Set trs = Nothing
        If rsClass(2) = 1 And StructureType <= 1 And ObjInstalled_FSO = True And (rsClass(3) & rsClass(4) <> "/" & ClassDir) Then
            Call DelClassDir(HtmlDir & rsClass(3) & rsClass(4))
        End If
        Conn.Execute ("update PE_Class set RootID=" & i & ",OrderID=0,ParentID=0,Child=0,ParentPath='0',Depth=0,PrevID=" & PrevID & ",NextID=" & NextID & ",arrChildID='" & rsClass(0) & "',ParentDir='/',ClassDir='" & ClassDir & "' where ClassID=" & rsClass(0))
        PrevID = rsClass(0)
        i = i + 1
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    Else
        SuccessMsg = "复位成功！请返回<a href='Admin_Class.asp'>栏目管理首页</a>做栏目的归属设置。"
        Call WriteSuccessMsg(SuccessMsg, ComeUrl)
    End If
    Call CreateJS_Class
    Call ClearSiteCache(0)
End Sub

Sub ResetChildClass()
    Dim ClassID, RootID, ParentPath, ParentDir, ClassDir
    Dim sql, rsClass, SuccessMsg, iCount, PrevID, NextID, i, trs
    ClassID = Trim(Request("ClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    Else
        ClassID = PE_CLng(ClassID)
    End If
    Set rsClass = Conn.Execute("select ClassID,RootID,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and ClassID=" & ClassID)
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
    Else
        RootID = rsClass(1)
        ParentPath = "0," & rsClass(0)
        ParentDir = "/" & rsClass(2) & "/"
    End If
    Set rsClass = Nothing
    If FoundErr = True Then Exit Sub

    sql = "select ClassID,ParentID,ClassType,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and RootID=" & RootID & " and ParentID>0 order by OrderID"
    Set rsClass = Server.CreateObject("adodb.recordset")
    rsClass.Open sql, Conn, 1, 1
    iCount = rsClass.RecordCount
    i = 1
    PrevID = 0
    Do While Not rsClass.EOF
        rsClass.MoveNext
        If rsClass.EOF Then
            NextID = 0
        Else
            NextID = rsClass(0)
        End If
        rsClass.moveprevious
        Set trs = Conn.Execute("select Count(ClassID) from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & ClassID & " and ClassID<>" & rsClass(0) & " and ClassDir='" & rsClass(4) & "'")
        If trs(0) > 1 Then
            ClassDir = rsClass(4) & rsClass(0)
        Else
            ClassDir = rsClass(4)
        End If
        Set trs = Nothing
        If rsClass(2) = 1 And StructureType <= 1 And ObjInstalled_FSO = True And (rsClass(3) & rsClass(4) <> ParentDir & ClassDir) Then
            Call DelClassDir(HtmlDir & rsClass(3) & rsClass(4))
        End If
        Conn.Execute ("update PE_Class set OrderID=" & i & ",ParentID=" & ClassID & ",Child=0,ParentPath='" & ParentPath & "',Depth=1,PrevID=" & PrevID & ",NextID=" & NextID & ",arrChildID='" & rsClass(0) & "',ParentDir='" & ParentDir & "',ClassDir='" & ClassDir & "' where ClassID=" & rsClass(0))
        PrevID = rsClass(0)
        i = i + 1
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    Conn.Execute ("update PE_Class set Child=" & i - 1 & " where ClassID=" & ClassID)
    
    SuccessMsg = "复位成功！请返回<a href='Admin_Class.asp'>栏目管理首页</a>做栏目的归属设置。"
    Call CreateJS_Class
    Call WriteSuccessMsg(SuccessMsg, ComeUrl)
    Call ClearSiteCache(0)
End Sub

Sub SaveUnite()
    Dim ClassID, TargetClassID, ParentID, ParentPath, Depth, Child, PrevID, NextID, arrChildID
    Dim rsClass, trs, i, SuccessMsg
    ClassID = Trim(Request("ClassID"))
    TargetClassID = Trim(Request("TargetClassID"))
    If ClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要合并的栏目！</li>"
    Else
        ClassID = PE_CLng(ClassID)
    End If
    If TargetClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定目标栏目！</li>"
    Else
        TargetClassID = PE_CLng(TargetClassID)
    End If
    If ClassID = TargetClassID Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请不要在相同栏目内进行操作</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    '判断目标栏目是否为外部栏目及是否有子栏目
    Set rsClass = Conn.Execute("select ClassID,Child,ClassType from PE_Class where ClassID=" & TargetClassID)
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>目标栏目不存在，可能已经被删除！</li>"
    Else
        If rsClass(1) > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>目标栏目中含有子栏目，不能合并！</li>"
        End If
        If rsClass(2) = 2 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>目标栏目是外部栏目，不能合并！</li>"
        End If
    End If
    Set rsClass = Nothing
    If FoundErr = True Then
        Exit Sub
    End If
    '得到当前栏目信息
    Set rsClass = Conn.Execute("select ClassID,ParentID,ParentPath,Depth,PrevID,NextID,arrChildID,ParentDir,ClassDir,ClassType from PE_Class where ClassID=" & ClassID)
    If rsClass.BOF And rsClass.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的栏目，可能已经被删除！</li>"
        rsClass.Close
        Set rsClass = Nothing
        Exit Sub
    End If
    ParentID = rsClass(1)
    ParentPath = rsClass(2)
    Depth = rsClass(3)
    PrevID = rsClass(4)
    NextID = rsClass(5)
    arrChildID = rsClass(6)

    '判断是否是合并到其下属栏目中
    Set trs = Conn.Execute("select ClassID from PE_Class where ClassID=" & TargetClassID & " and ClassID in (" & arrChildID & ")")
    If Not (trs.BOF And trs.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不能将一个栏目合并到其下属子栏目中</li>"
    End If
    Set trs = Nothing
    
    If FoundErr = True Then
        Set rsClass = Nothing
        Exit Sub
    End If
    If rsClass("ClassType") = 1 And UseCreateHTML > 0 Then
        '删除栏目目录
        Select Case StructureType
        Case 0, 1, 2
            Call DelClassDir(HtmlDir & rsClass("ParentDir") & rsClass("ClassDir"))
        Case 3, 4, 5
            Call DelClassDir(HtmlDir & "/" & rsClass("ClassDir"))
        Case Else
            '不用做任何处理
        End Select
    End If
    Set rsClass = Nothing

    
    '先修改上一栏目的NextID和下一栏目的PrevID
    If PrevID > 0 Then
        Conn.Execute "update PE_Class set NextID=" & NextID & " where ClassID=" & PrevID
    End If
    If NextID > 0 Then
        Conn.Execute "update PE_Class set PrevID=" & PrevID & " where ClassID=" & NextID
    End If
    
    '更新文章及评论所属栏目
    Conn.Execute ("update " & SheetName & " set ClassID=" & TargetClassID & " where ChannelID=" & ChannelID & " and ClassID in (" & arrChildID & ")")
    
    '删除被合并栏目及其下属栏目
    Conn.Execute ("delete from PE_Class where ChannelID=" & ChannelID & " and  ClassID in (" & arrChildID & ")")
    
    '更新其原来所属栏目的子栏目数，排序相当于剪枝而不需考虑
    If ParentID > 0 Then
        Conn.Execute ("update PE_Class set Child=Child-1 where ClassID=" & ParentID)

        '更新此栏目的原来所有上级栏目的子栏目ID数组
        Set trs = Conn.Execute("select ClassID,arrChildID from PE_Class where ClassID in (" & ParentPath & ")")
        Do While Not trs.EOF
            Conn.Execute ("update PE_Class set arrChildID='" & RemoveClassID(trs(1), arrChildID) & "' where ClassID=" & trs(0))
            trs.MoveNext
        Loop
        trs.Close
        Set trs = Nothing
    End If


    Call CreateJS_Class
    
    SuccessMsg = "栏目合并成功！已经将被合并栏目及其下属子栏目的所有数据转入目标栏目中。<br><br>同时删除了被合并的栏目及其子栏目。"
    If UseCreateHTML > 0 Then
        SuccessMsg = SuccessMsg & "<br><br>请重新生成目标栏目的所有文章！"
    End If
    Call WriteSuccessMsg(SuccessMsg, ComeUrl)
    Call ClearSiteCache(0)
End Sub

Sub DoBatch()
    Dim ClassID, ClassPurview, arrGroupID_Browse, arrGroupID_View, arrGroupID_Input, EnableComment, CheckComment
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent
    Dim OpenType, ShowOnTop, ShowOnIndex, IsElite, EnableAdd, EnableProtect, SkinID, TemplateID
    Dim MaxPerPage, DefaultItemSkin, DefaultItemTemplate, ItemListOrderType, ItemOpenType
    Dim sql, rsClass, i
    Dim CommandClassPoint, ReleaseClassPoint
    ClassID = Trim(Request("ClassID"))
    ClassPurview = PE_CLng(Trim(Request("ClassPurview")))
    arrGroupID_Browse = ReplaceBadChar(Trim(Request("arrGroupID_Browse")))
    arrGroupID_View = ReplaceBadChar(Trim(Request("arrGroupID_View")))
    arrGroupID_Input = ReplaceBadChar(Trim(Request("arrGroupID_Input")))
    EnableComment = PE_CBool(Trim(Request("EnableComment")))
    CheckComment = PE_CBool(Trim(Request("CheckComment")))

    PresentExp = PE_CDbl(Trim(Request("PresentExp")))
    DefaultItemPoint = PE_CDbl(Trim(Request("DefaultItemPoint")))
    DefaultItemChargeType = PE_CLng(Trim(Request.Form("DefaultItemChargeType")))
    DefaultItemPitchTime = PE_CLng(Trim(Request.Form("DefaultItemPitchTime")))
    DefaultItemReadTimes = PE_CLng(Trim(Request.Form("DefaultItemReadTimes")))
    DefaultItemDividePercent = PE_CLng(Trim(Request.Form("DefaultItemDividePercent")))

    OpenType = PE_CLng(Trim(Request("OpenType")))
    ShowOnTop = PE_CBool(Trim(Request("ShowOnTop")))
    ShowOnIndex = PE_CBool(Trim(Request("ShowOnIndex")))
    IsElite = PE_CBool(Trim(Request("IsElite")))
    EnableAdd = PE_CBool(Trim(Request("EnableAdd")))
    EnableProtect = PE_CBool(Trim(Request("EnableProtect")))
    SkinID = PE_CLng(Trim(Request("SkinID")))
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    DefaultItemSkin = PE_CLng(Trim(Request("DefaultItemSkin")))
    DefaultItemTemplate = PE_CLng(Trim(Request("DefaultItemTemplate")))
    ItemListOrderType = PE_CLng(Trim(Request("ItemListOrderType")))
    ItemOpenType = PE_CLng(Trim(Request("ItemOpenType")))
    CommandClassPoint = PE_CLng(Trim(Request.Form("CommandClassPoint")))
    ReleaseClassPoint = PE_CLng(Trim(Request.Form("ReleaseClassPoint")))

    If IsValidID(ClassID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定要批量修改设置的栏目！</li>"
    Else
        ClassID = ReplaceBadChar(ClassID)
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    sql = "select * from PE_Class where ChannelID=" & ChannelID & " and ClassID in (" & ClassID & ")"
    Set rsClass = Server.CreateObject("Adodb.recordset")
    rsClass.Open sql, Conn, 1, 3
    Do While Not rsClass.EOF
        If Trim(Request("ModifyClassPurview")) = "Yes" Then rsClass("ClassPurview") = ClassPurview
        If Trim(Request("ModifyEnableComment")) = "Yes" Then rsClass("EnableComment") = EnableComment
        If Trim(Request("ModifyCheckComment")) = "Yes" Then rsClass("CheckComment") = CheckComment

        If Trim(Request("ModifyPresentExp")) = "Yes" Then rsClass("PresentExp") = PresentExp
        If Trim(Request("ModifyDefaultItemPoint")) = "Yes" Then rsClass("DefaultItemPoint") = DefaultItemPoint
        If Trim(Request("ModifyDefaultItemChargeType")) = "Yes" Then
            rsClass("DefaultItemChargeType") = DefaultItemChargeType
            rsClass("DefaultItemPitchTime") = DefaultItemPitchTime
            rsClass("DefaultItemReadTimes") = DefaultItemReadTimes
            rsClass("DefaultItemDividePercent") = DefaultItemDividePercent
        End If
        If Trim(Request("ModifyReleasePoint")) = "Yes" Then rsClass("ReleaseClassPoint") = ReleaseClassPoint
        If Trim(Request("ModifyCommandClassPoint")) = "Yes" Then rsClass("CommandClassPoint") = CommandClassPoint
        If Trim(Request("ModifyOpenType")) = "Yes" Then rsClass("OpenType") = OpenType
        If Trim(Request("ModifyShowOnTop")) = "Yes" Then rsClass("ShowOnTop") = ShowOnTop
        If Trim(Request("ModifyShowOnIndex")) = "Yes" Then rsClass("ShowOnIndex") = ShowOnIndex
        If Trim(Request("ModifyIsElite")) = "Yes" Then rsClass("IsElite") = IsElite
        If Trim(Request("ModifyEnableAdd")) = "Yes" Then rsClass("EnableAdd") = EnableAdd
        If Trim(Request("ModifyEnableProtect")) = "Yes" Then rsClass("EnableProtect") = EnableProtect
        If Trim(Request("ModifySkinID")) = "Yes" Then rsClass("SkinID") = SkinID
        If Trim(Request("ModifyTemplateID")) = "Yes" Then rsClass("TemplateID") = TemplateID
        If Trim(Request("ModifyMaxPerPage")) = "Yes" Then rsClass("MaxPerPage") = MaxPerPage
        If Trim(Request("ModifyDefaultItemSkin")) = "Yes" Then rsClass("DefaultItemSkin") = DefaultItemSkin
        If Trim(Request("ModifyDefaultItemTemplate")) = "Yes" Then rsClass("DefaultItemTemplate") = DefaultItemTemplate
        If Trim(Request("ModifyItemListOrderType")) = "Yes" Then rsClass("ItemListOrderType") = ItemListOrderType
        If Trim(Request("ModifyItemOpenType")) = "Yes" Then rsClass("ItemOpenType") = ItemOpenType
        rsClass.Update

        If Trim(Request("ModifyGroupPurview_Browse")) = "Yes" Then
            Call ModifyGroupPurview("Browse", arrGroupID_Browse, rsClass("ClassID"))
        End If
        If Trim(Request("ModifyGroupPurview_View")) = "Yes" Then
            Call ModifyGroupPurview("View", arrGroupID_View, rsClass("ClassID"))
        End If
        If Trim(Request("ModifyGroupPurview_Input")) = "Yes" Then
            Call ModifyGroupPurview("Input", arrGroupID_Input, rsClass("ClassID"))
        End If
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    
    '从父路径中继承栏目权限并更新本栏目及其子栏目的权限
    Call UpdateClassPurview(ClassID)
    
    Call CreateJS_Class
    Dim msg
    msg = "批量设置栏目属性成功！"
    If UseCreateHTML > 0 Then
        msg = msg & "请记得重新生成有关栏目的页面。"
    End If
    Call WriteSuccessMsg(msg, ComeUrl)
    Call ClearSiteCache(0)
End Sub

Function RemoveClassID(ByVal arrClassID_Parent, ByVal arrClassID_Child)
    Dim arrClassID, arrClassID2, arrClassID3, i, j, bFound
    If IsNull(arrClassID_Parent) Then
        RemoveClassID = ""
        Exit Function
    End If
    If IsNull(arrClassID_Parent) Then
        RemoveClassID = arrClassID_Parent
        Exit Function
    End If
    If Trim(arrClassID_Parent) = Trim(arrClassID_Child) Then
        RemoveClassID = ""
        Exit Function
    End If
    arrClassID = Split(arrClassID_Parent, ",")
    arrClassID3 = ""
    If InStr(arrClassID_Child, ",") > 0 Then
        arrClassID2 = Split(arrClassID_Child, ",")
        For i = 0 To UBound(arrClassID)
            bFound = False
            For j = 0 To UBound(arrClassID2)
                If PE_CLng(arrClassID(i)) = PE_CLng(arrClassID2(j)) Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound = False Then
                If arrClassID3 = "" Then
                    arrClassID3 = arrClassID(i)
                Else
                    arrClassID3 = arrClassID3 & "," & arrClassID(i)
                End If
            End If
        Next
    Else
        For i = 0 To UBound(arrClassID)
            If PE_CLng(arrClassID(i)) <> PE_CLng(arrClassID_Child) Then
                If arrClassID3 = "" Then
                    arrClassID3 = arrClassID(i)
                Else
                    arrClassID3 = arrClassID3 & "," & arrClassID(i)
                End If
            End If
        Next
    End If
    RemoveClassID = arrClassID3
End Function

Sub CreateJS_Class()
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    
    Dim hf, strTopMenu, strClassTree, strNavigation, strOption, strForm

    Select Case TopMenuType
    Case 0, 1
        strTopMenu = GetRootClass_Menu()
    Case 2
        strTopMenu = "var h,w,l,t;" & vbCrLf
        strTopMenu = strTopMenu & "var topMar = 1;" & vbCrLf
        strTopMenu = strTopMenu & "var leftMar = -2;" & vbCrLf
        strTopMenu = strTopMenu & "var space = 1;" & vbCrLf
        strTopMenu = strTopMenu & "var isvisible;" & vbCrLf
        strTopMenu = strTopMenu & "var MENU_SHADOW_COLOR='#999999';" & vbCrLf
        strTopMenu = strTopMenu & "var global = window.document" & vbCrLf
        strTopMenu = strTopMenu & "global.fo_currentMenu = null" & vbCrLf
        strTopMenu = strTopMenu & "global.fo_shadows = new Array" & vbCrLf
 
        strTopMenu = strTopMenu & GetJS_ClassMenu() & vbCrLf
        strTopMenu = strTopMenu & "document.write(" & Chr(34) & GetRootClass(1) & Chr(34) & ");"
    Case 3
        strTopMenu = "document.write(" & Chr(34) & GetRootClass(2) & Chr(34) & ");"
    End Select
    If Not fso.FolderExists(Server.MapPath(InstallDir & ChannelDir & "/js")) Then
        fso.CreateFolder Server.MapPath(InstallDir & ChannelDir & "/js")
    End If
    Call WriteToFile(InstallDir & ChannelDir & "/js/ShowClass_Menu.js", strTopMenu)
    
    strClassTree = GetClass_Tree()
    Call WriteToFile(InstallDir & ChannelDir & "/js/ShowClass_Tree.js", "document.write(""" & strClassTree & """);")


    Select Case ClassGuideType
    Case 1
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 2) & """);"
    Case 2
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 3) & """);"
    Case 3
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 4) & """);"
    Case 4
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 5) & """);"
    Case 5
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 6) & """);"
    Case 6
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 7) & """);"
    Case 7
        strNavigation = "document.write(""" & GetClass_Navigation(1, 0, 8) & """);"
    Case 8
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 2) & """);"
    Case 9
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 3) & """);"
    Case 10
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 4) & """);"
    Case 11
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 5) & """);"
    Case 12
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 6) & """);"
    Case 13
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 7) & """);"
    Case 14
        strNavigation = "document.write(""" & GetClass_Navigation(2, 1, 8) & """);"
    Case 15
        strNavigation = "document.write(""" & GetClass_Navigation(2, 2, 2) & """);"
    Case 16
        strNavigation = "document.write(""" & GetClass_Navigation(2, 2, 3) & """);"
    Case 17
        strNavigation = "document.write(""" & GetClass_Navigation(2, 2, 4) & """);"
    Case 18
        strNavigation = "document.write(""" & GetClass_Navigation(2, 2, 5) & """);"
    Case 19
        strNavigation = "document.write(""" & GetClass_Navigation(2, 2, 6) & """);"
    End Select
    Call WriteToFile(InstallDir & ChannelDir & "/js/ShowClass_Navigation.js", strNavigation)

    strOption = GetClass_Option(ChannelID, 0)
    Call WriteToFile(InstallDir & ChannelDir & "/js/ShowClass_Option.js", "document.write(""" & strOption & """);")
    
    strForm = ShowSearchForm(2, 0)
    Call WriteToFile(InstallDir & ChannelDir & "/js/ShowSearchForm.js", "document.write(""" & strForm & """);")
End Sub

Function GetClass_Option(iChannelID, CurrentID)
    Dim rsClass, sqlClass, strTemp, tmpDepth, i
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    sqlClass = "Select ClassID,ClassName,ClassType,Depth,NextID from PE_Class where ChannelID=" & iChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strTemp = "<option value=''>请先添加栏目</option>"
    Else
        strTemp = ""
        Do While Not rsClass.EOF
            tmpDepth = rsClass(3)
            If rsClass(4) > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            strTemp = strTemp & "<option value='" & rsClass(0) & "'"
            If CurrentID > 0 And rsClass(0) = CurrentID Then
                 strTemp = strTemp & " selected"
            End If
            strTemp = strTemp & ">"
            
            If tmpDepth > 0 Then
                For i = 1 To tmpDepth
                    strTemp = strTemp & "&nbsp;&nbsp;"
                    If i = tmpDepth Then
                        If rsClass(4) > 0 Then
                            strTemp = strTemp & "├&nbsp;"
                        Else
                            strTemp = strTemp & "└&nbsp;"
                        End If
                    Else
                        If arrShowLine(i) = True Then
                            strTemp = strTemp & "│"
                        Else
                            strTemp = strTemp & "&nbsp;"
                        End If
                    End If
                Next
            End If
            strTemp = strTemp & rsClass(1)
            If rsClass(2) = 2 Then
                strTemp = strTemp & "(外)"
            End If
            strTemp = strTemp & "</option>"
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing

    GetClass_Option = strTemp
End Function



Function GetOrderType_Option(OrderType)
    Dim strOrderType
    strOrderType = strOrderType & "<option value='1'"
    If OrderType = 1 Then strOrderType = strOrderType & " selected"
    strOrderType = strOrderType & ">" & ChannelShortName & "ID（降序）</option>"
    strOrderType = strOrderType & "<option value='2'"
    If OrderType = 2 Then strOrderType = strOrderType & " selected"
    strOrderType = strOrderType & ">" & ChannelShortName & "ID（升序）</option>"
    strOrderType = strOrderType & "<option value='3'"
    If OrderType = 3 Then strOrderType = strOrderType & " selected"
    strOrderType = strOrderType & ">更新时间（降序）</option>"
    strOrderType = strOrderType & "<option value='4'"
    If OrderType = 4 Then strOrderType = strOrderType & " selected"
    strOrderType = strOrderType & ">更新时间（升序）</option>"
    strOrderType = strOrderType & "<option value='5'"
    If OrderType = 5 Then strOrderType = strOrderType & " selected"
    strOrderType = strOrderType & ">点击次数（降序）</option>"
    strOrderType = strOrderType & "<option value='6'"
    If OrderType = 6 Then strOrderType = strOrderType & " selected"
    strOrderType = strOrderType & ">点击次数（升序）</option>"
    GetOrderType_Option = strOrderType
End Function

Function GetOpenType_Option(OpenType)
    Dim strOpenType
    strOpenType = "<option value='0'"
    If OpenType = 0 Then
        strOpenType = strOpenType & " selected"
    End If
    strOpenType = strOpenType & ">" & "在原窗口打开</option><option value='1'"
    If OpenType = 1 Then
        strOpenType = strOpenType & " selected"
    End If
    strOpenType = strOpenType & ">" & "在新窗口打开</option>"
    GetOpenType_Option = strOpenType
End Function

Function GetPath(ParentID, ParentPath)
    Dim strPath, i
    If ParentID <= 0 Then
        GetPath = "无（作为一级栏目）"
        Exit Function
    End If
    Dim rsParent, sqlParent
    sqlParent = "Select * from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
    Set rsParent = Conn.Execute(sqlParent)
    Do While Not rsParent.EOF
        For i = 1 To rsParent("Depth")
            strPath = strPath & "&nbsp;&nbsp;&nbsp;"
        Next
        If rsParent("Depth") > 0 Then
            strPath = strPath & "└&nbsp;"
        End If
        strPath = strPath & rsParent("ClassName") & "<br>"
        rsParent.MoveNext
    Loop
    rsParent.Close
    Set rsParent = Nothing
    GetPath = strPath
End Function

'=================================================
'函数名：GetRootClass_Menu
'作  用：得到栏目无级下拉菜单效果的HTML代码
'参  数：无
'返回值：栏目无级下拉菜单效果的HTML代码
'=================================================
Function GetRootClass_Menu()
    Dim Class_MenuTitle, strJS, strClassUrl
    ClassLink = XmlText("BaseText", "ClassLink", "|")
    pNum = 1
    pNum2 = 0
    strJS = "stm_bm(['uueoehr',400,'','" & strInstallDir & "images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);" & vbCrLf
    strJS = strJS & "stm_bp('p0',[0,4,0,0,2,2,0,0,100,'',-2,'',-2,90,0,0,'#000000','transparent','',3,0,0,'#000000']);" & vbCrLf
    strJS = strJS & "stm_ai('p0i0',[0,'" & ClassLink & "','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbCrLf
    If UseCreateHTML > 0 Then
        strClassUrl = ChannelUrl & "/Index" & FileExt_Index
    Else
        strClassUrl = ChannelUrl & "/Index.asp"
    End If
    strJS = strJS & "stm_aix('p0i1','p0i0',[0,'" & ChannelName & "首页','','',-1,-1,0,'" & strClassUrl & "','_self','" & strClassUrl & "','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体','9pt 宋体']);" & vbCrLf
    strJS = strJS & "stm_aix('p0i2','p0i0',[0,'" & ClassLink & "','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbCrLf

    Dim sqlRoot, rsRoot, j
    sqlRoot = "select * from PE_Class where ChannelID=" & ChannelID & " and Depth=0 and ShowOnTop=" & PE_True & " order by RootID"
    Set rsRoot = Conn.Execute(sqlRoot)
    If Not (rsRoot.BOF And rsRoot.EOF) Then
        j = 3
        Do While Not rsRoot.EOF
            If rsRoot("OpenType") = 0 Then
                OpenType_Class = "_self"
            Else
                OpenType_Class = "_blank"
            End If
            If Trim(rsRoot("Tips")) <> "" Then
                Class_MenuTitle = Replace(Replace(Replace(Replace(rsRoot("Tips"), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
            Else
                Class_MenuTitle = ""
            End If
            If rsRoot("ClassType") = 1 Then
                strClassUrl = GetClassUrl(rsRoot("ParentDir"), rsRoot("ClassDir"), rsRoot("ClassID"), rsRoot("ClassPurview"))
                strJS = strJS & "stm_aix('p0i" & j & "','p0i0',[0,'" & rsRoot("ClassName") & "','','',-1,-1,0,'" & strClassUrl & "','" & OpenType_Class & "','" & strClassUrl & "','" & Class_MenuTitle & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体','9pt 宋体']);" & vbCrLf
                If rsRoot("Child") > 0 Then
                    strJS = strJS & GetClassMenu(rsRoot("ClassID"), 0)
                End If
            Else
                strJS = strJS & "stm_aix('p0i" & j & "','p0i0',[0,'" & rsRoot("ClassName") & "','','',-1,-1,0,'" & rsRoot("LinkUrl") & "','" & OpenType_Class & "','" & rsRoot("LinkUrl") & "','" & Class_MenuTitle & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体','9pt 宋体']);" & vbCrLf
            End If
            strJS = strJS & "stm_aix('p0i2','p0i0',[0,'" & ClassLink & "','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbCrLf
            j = j + 1
            rsRoot.MoveNext
            If MaxPerLine > 0 Then
                If (j - 2) Mod MaxPerLine = 0 And Not rsRoot.EOF Then
                    strJS = strJS & "stm_em();" & vbCrLf
                    strJS = strJS & "stm_bm(['uueoehr',400,'','" & strInstallDir & "images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);" & vbCrLf
                    strJS = strJS & "stm_bp('p0',[0,4,0,0,2,2,0,0,100,'',-2,'',-2,90,0,0,'#000000','transparent','',3,0,0,'#000000']);" & vbCrLf
                    strJS = strJS & "stm_ai('p0i0',[0,'" & ClassLink & "','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbCrLf
                End If
            End If
        Loop
    End If
    rsRoot.Close
    Set rsRoot = Nothing
    strJS = strJS & "stm_em();" & vbCrLf

    GetRootClass_Menu = strJS
End Function

Function GetClassMenu(ID, ShowType)
    Dim sqlClass, rsClass, Sub_MenuTitle, k, strJS, strClassUrl
    strJS = ""
    If pNum = 1 Then
        strJS = strJS & "stm_bp('p" & pNum & "',[1,4,0,0,2,3,6,7,100,'progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)',-2,'',-2,67,2,3,'#999999','#ffffff','',3,1,1,'#aca899']);" & vbCrLf
    Else
        If ShowType = 0 Then
            strJS = strJS & "stm_bpx('p" & pNum & "','p" & pNum2 & "',[1,4,0,0,2,3,6]);" & vbCrLf
        Else
            strJS = strJS & "stm_bpx('p" & pNum & "','p" & pNum2 & "',[1,2,-2,-3,2,3,0]);" & vbCrLf
        End If
    End If
    
    k = 0
    sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " and ParentID=" & ID & " order by OrderID asc"
    Set rsClass = Conn.Execute(sqlClass)
    'set rsClass=conn.execute("GetChildClass_Article_Menu " & ID)
    Do While Not rsClass.EOF
        If rsClass("OpenType") = 0 Then
            OpenType_Class = "_self"
        Else
            OpenType_Class = "_blank"
        End If
        If Trim(rsClass("Tips")) <> "" Then
            Sub_MenuTitle = Replace(Replace(Replace(Replace(rsClass("Tips"), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
        Else
            Sub_MenuTitle = ""
        End If
        If rsClass("ClassType") = 1 Then
            strClassUrl = GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), rsClass("ClassPurview"))
            If rsClass("Child") > 0 Then
                strJS = strJS & "stm_aix('p" & pNum & "i" & k & "','p" & pNum2 & "i0',[0,'" & rsClass("ClassName") & "','','',-1,-1,0,'" & strClassUrl & "','" & OpenType_Class & "','" & strClassUrl & "','" & Sub_MenuTitle & "','','',6,0,0,'" & strInstallDir & "images/arrow_r.gif','" & strInstallDir & "images/arrow_w.gif',7,7,0,0,1,'#ffffff',0,'#cccccc',0,'','',3,3,0,0,'#fffff7','#000000','#000000','#ffffff','9pt 宋体']);" & vbCrLf
                pNum = pNum + 1
                pNum2 = pNum2 + 1
                strJS = strJS & GetClassMenu(rsClass("ClassID"), 1)
            Else
                strJS = strJS & "stm_aix('p" & pNum & "i" & k & "','p" & pNum2 & "i0',[0,'" & rsClass("ClassName") & "','','',-1,-1,0,'" & strClassUrl & "','" & OpenType_Class & "','" & strClassUrl & "','" & Sub_MenuTitle & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',0,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体']);" & vbCrLf
            End If
        Else
            strJS = strJS & "stm_aix('p" & pNum & "i" & k & "','p" & pNum2 & "i0',[0,'" & rsClass("ClassName") & "','','',-1,-1,0,'" & rsClass("LinkUrl") & "','" & OpenType_Class & "','" & rsClass("LinkUrl") & "','" & Sub_MenuTitle & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',0,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体']);" & vbCrLf
        End If
        k = k + 1
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    strJS = strJS & "stm_ep();" & vbCrLf

    GetClassMenu = strJS
End Function

Function GetJS_ClassMenu()
    Dim sqlMenu, rsMenu, strMenu, PrevRootID, tmpDepth, i, strClassUrl
    sqlMenu = "select * from PE_Class where ChannelID=" & ChannelID & " and Depth=1 order by RootID,OrderID"
    Set rsMenu = Conn.Execute(sqlMenu)
    If rsMenu.BOF And rsMenu.EOF Then
        strMenu = "var menu0='没有任何子栏目';"
    Else
        strMenu = "var menu" & rsMenu("RootID") & "=" & Chr(34)
        If rsMenu("ClassType") = 2 Then
            strClassUrl = rsMenu("LinkUrl")
        Else
            strClassUrl = GetClassUrl(rsMenu("ParentDir"), rsMenu("ClassDir"), rsMenu("ClassID"), rsMenu("ClassPurview"))
        End If
        strMenu = strMenu & "&nbsp;<a style=font-size:9pt;line-height:14pt; href='" & strClassUrl & "'>" & rsMenu("ClassName") & "</a><br>"
        PrevRootID = rsMenu("RootID")
        rsMenu.MoveNext
        Do While Not rsMenu.EOF
            If rsMenu("RootID") <> PrevRootID Then
                strMenu = strMenu & Chr(34) & ";" & vbCrLf & "var menu" & rsMenu("RootID") & "=" & Chr(34)
            End If
            If rsMenu("ClassType") = 2 Then
                strClassUrl = rsMenu("LinkUrl")
            Else
                strClassUrl = GetClassUrl(rsMenu("ParentDir"), rsMenu("ClassDir"), rsMenu("ClassID"), rsMenu("ClassPurview"))
            End If
            strMenu = strMenu & "&nbsp;<a style=font-size:9pt;line-height:14pt; href='" & strClassUrl & "'>" & rsMenu("ClassName") & "</a><br>"
            
            PrevRootID = rsMenu("RootID")
            rsMenu.MoveNext
        Loop
        strMenu = strMenu & Chr(34) & ";" & vbCrLf
    End If
    rsMenu.Close
    Set rsMenu = Nothing
    GetJS_ClassMenu = strMenu
End Function

'=================================================
'函数名：GetRootClass
'作  用：显示一级栏目（无特殊效果）
'参  数：ShowType   ----显示方式，1为普通下拉菜单式，2为纯文字式，无菜单效果
'=================================================
Function GetRootClass(ShowType)
    Dim sqlRoot, rsRoot, strRoot, strClassUrl, iCount
    ClassLink = XmlText("BaseText", "ClassLink", "|")
    sqlRoot = "select * from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and ShowOnTop=" & PE_True & " order by RootID"
    Set rsRoot = Conn.Execute(sqlRoot)
    If rsRoot.BOF And rsRoot.EOF Then
        strRoot = "还没有任何栏目，请首先添加栏目。"
    Else
        If UseCreateHTML > 0 Then
            strRoot = strRoot & "" & ClassLink & "<a href='" & ChannelUrl & "/Index" & FileExt_Index & "'>&nbsp;" & ChannelName & "首页&nbsp;</a>" & ClassLink & ""
        Else
            strRoot = strRoot & "" & ClassLink & "<a href='" & ChannelUrl & "/Index.asp'>&nbsp;" & ChannelName & "首页&nbsp;</a>" & ClassLink & ""
        End If
        Do While Not rsRoot.EOF
            If rsRoot("ClassType") = 2 Then
                strRoot = strRoot & "<a href='" & rsRoot("LinkUrl") & "' target='_blank'>&nbsp;" & rsRoot("ClassName") & "&nbsp;</a>" & ClassLink & ""
            Else
                strClassUrl = GetClassUrl(rsRoot("ParentDir"), rsRoot("ClassDir"), rsRoot("ClassID"), rsRoot("ClassPurview"))
                strRoot = strRoot & "<a href='" & strClassUrl & "'"
                If rsRoot("Child") > 0 And ShowType = 1 Then
                    strRoot = strRoot & " onMouseOver='ShowMenu(menu" & rsRoot("RootID") & ",100)'"
                End If
                strRoot = strRoot & ">&nbsp;" & rsRoot("ClassName") & "&nbsp;</a>" & ClassLink & ""
            End If
            rsRoot.MoveNext
            iCount = iCount + 1
            If iCount Mod MaxPerLine = 0 And Not rsRoot.EOF Then
                strRoot = strRoot & "<br>" & ClassLink & ""
            End If
        Loop
    End If
    rsRoot.Close
    Set rsRoot = Nothing
    GetRootClass = strRoot
End Function


'=================================================
'函数名：GetClass_Tree
'作  用：得到所有栏目的树形目录效果的HTML代码
'参  数：无
'返回值：栏目的树形目录效果的HTML代码
'=================================================
Function GetClass_Tree()
    Dim arrShowLine(20), Class_MenuTitle, i, strClassUrl
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    Dim rsClass, sqlClass, tmpDepth, strClassTree
    sqlClass = "select ClassID,ClassName,Depth,ParentID,NextID,LinkUrl,Child,Readme,ClassType,ParentDir,ClassDir,OpenType,ClassPurview from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strClassTree = "没有任何栏目"
    Else
        strClassTree = ""
        Do While Not rsClass.EOF
            tmpDepth = rsClass(2)
            If rsClass(4) > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            If Trim(rsClass(7)) <> "" Then
                Class_MenuTitle = Replace(Replace(Replace(Replace(rsClass(7), "'", ""), """", ""), Chr(10), ""), Chr(13), "")
            Else
                Class_MenuTitle = ""
            End If
            If tmpDepth > 0 Then
                For i = 1 To tmpDepth
                    If i = tmpDepth Then
                        If rsClass(4) > 0 Then
                            strClassTree = strClassTree & "<img src='"& strInstallDir &"images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
                        Else
                            strClassTree = strClassTree & "<img src='"& strInstallDir &"images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
                        End If
                    Else
                        If arrShowLine(i) = True Then
                            strClassTree = strClassTree & "<img src='"& strInstallDir &"images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
                        Else
                            strClassTree = strClassTree & "<img src='"& strInstallDir &"images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
                        End If
                    End If
                Next
            End If
            If rsClass(6) > 0 Then
                strClassTree = strClassTree & "<img src='"& strInstallDir &"images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
            Else
                strClassTree = strClassTree & "<img src='"& strInstallDir &"images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
            End If
            
            If rsClass("ClassType") = 2 Then
                strClassUrl = rsClass("LinkUrl")
            Else
                strClassUrl = GetClassUrl(rsClass("ParentDir"), rsClass("ClassDir"), rsClass("ClassID"), rsClass("ClassPurview"))
            End If
            strClassTree = strClassTree & "<a href='" & strClassUrl & "' title='" & Class_MenuTitle & "'"
            If rsClass(11) = 0 Then
                strClassTree = strClassTree & " target='_top'"
            Else
                strClassTree = strClassTree & " target='_blank'"
            End If
            If rsClass(2) = 0 Then
                strClassTree = strClassTree & "><b>" & rsClass(1) & "</b>"
            Else
                strClassTree = strClassTree & ">" & rsClass(1)
            End If
            If rsClass(8) = 2 Then
                strClassTree = strClassTree & "(外)"
            End If
            strClassTree = strClassTree & "</a>"
            If rsClass(6) > 0 Then
                strClassTree = strClassTree & "（" & rsClass(6) & "）"
            End If
            strClassTree = strClassTree & "<br>"
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    GetClass_Tree = strClassTree
End Function

'==================================================
'函数名：ShowSearchForm
'作  用：显示搜索表单
'参  数：ShowType ----显示方式。1为简洁模式，2为标准模式，3为高级模式
'        CurrentID ----当前栏目ID
'返回值：搜索表单的HTML代码
'==================================================
Function ShowSearchForm(ShowType, CurrentID)
    Dim strForm
    If ShowType <> 1 And ShowType <> 2 And ShowType <> 3 Then
        ShowType = 1
    End If
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='" & ChannelUrl & "/Search.asp'>"
    strForm = strForm & "<tr><td height='28' align='center'>"
    If ShowType = 1 Then
        Select Case ModuleType
        Case 1
            strForm = strForm & "<input type='hidden' name='field' value='Title'>"
        Case 2
            strForm = strForm & "<input type='hidden' name='field' value='SoftName'>"
        Case 3
            strForm = strForm & "<input type='hidden' name='field' value='PhotoName'>"
        Case 5
            strForm = strForm & "<input type='hidden' name='field' value='ProductName'>"
        End Select
        strForm = strForm & "<input type='text' name='keyword'  size='15' value='关键字' maxlength='50' onFocus='this.select();'>&nbsp;"
        strForm = strForm & "<input type='submit' name='Submit'  value='搜索'>"
    ElseIf ShowType = 2 Then
        strForm = strForm & "<select name='Field' size='1'>"
        Select Case ModuleType
        Case 1
            strForm = strForm & "<option value='Title' selected>" & ChannelShortName & "标题</option>"
        Case 2
            strForm = strForm & "<option value='SoftName' selected>" & ChannelShortName & "名称</option>"
        Case 3
            strForm = strForm & "<option value='PhotoName' selected>" & ChannelShortName & "名称</option>"
        Case 5
            strForm = strForm & "<option value='ProductName' selected>" & ChannelShortName & "名称</option>"
        End Select
        If SearchContent = True Then
            Select Case ModuleType
            Case 1
                strForm = strForm & "<option value='Content'>" & ChannelShortName & "内容</option>"
            Case 2
                strForm = strForm & "<option value='SoftIntro'>" & ChannelShortName & "简介</option>"
            Case 3
                strForm = strForm & "<option value='PhotoIntro'>" & ChannelShortName & "简介</option>"
            Case 5
                strForm = strForm & "<option value='ProductIntro'>" & ChannelShortName & "简介</option>"
            End Select
        End If
        If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Then
            strForm = strForm & "<option value='Author'>" & ChannelShortName & "作者</option>"
            strForm = strForm & "<option value='Inputer'>录 入 者</option>"
        ElseIf ModuleType = 5 Then
            strForm = strForm & "<option value='ProducerName'>厂商</option>"
            strForm = strForm & "<option value='TrademarkName'>品牌/商标</option>"
            If IsExists("MY_Namepy","PE_Product") = True Then strForm = strForm & "<option value='MY_Namepy'>品名简拼</option>"
            strForm = strForm & "<option value='ProductNum'>商品编号</option>"
        End If
        strForm = strForm & "<option value='Keywords'>关键字</option>"
        If SearchContent = True Then
            Select Case ModuleType
            Case 1
                strForm = strForm & "<option value='ArticleID'>" & ChannelShortName & "ID</option>"
            Case 2
                strForm = strForm & "<option value='SoftID'>" & ChannelShortName & "ID</option>"
            Case 3
                strForm = strForm & "<option value='PhotoID'>" & ChannelShortName & "ID</option>"
            Case 5
                strForm = strForm & "<option value='ProductID'>" & ChannelShortName & "ID</option>"
            End Select
        End If
        strForm = strForm & "</select>&nbsp;"
        strForm = strForm & "<select name='ClassID'><option value=''>所有栏目</option>" & GetClass_Option(ChannelID, 0) & "</select>&nbsp;"
        strForm = strForm & "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>&nbsp;"
        strForm = strForm & "<input type='submit' name='Submit'  value=' 搜索 '>"
    ElseIf ShowType = 3 Then
    
    End If
    strForm = strForm & "</td></tr></form></table>"
    ShowSearchForm = strForm
End Function


Sub DelInfo(arrClassID)
    'On Error Resume Next
    Dim sqlDel, rsDel
    Dim InfoPath, FileExt

    If IsValidID(arrClassID) = False Then Exit Sub
    Select Case ModuleType
    Case 1
        sqlDel = "select ArticleID as InfoID,UpdateTime,Inputer,Deleted,PaginationType from PE_Article"
    Case 2
        sqlDel = "select SoftID as InfoID,UpdateTime,Inputer,Deleted from PE_Soft"
    Case 3
        sqlDel = "select PhotoID as InfoID,UpdateTime,Inputer,Deleted from PE_Photo"
    Case 5
        sqlDel = "select ProductID as InfoID,UpdateTime,Inputer,Deleted from PE_Product"
    End Select
    If InStr(arrClassID, ",") > 0 Then
        sqlDel = sqlDel & " where ClassID in (" & arrClassID & ")"
    Else
        sqlDel = sqlDel & " where ClassID=" & arrClassID & ""
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        InfoPath = HtmlDir & GetItemPath(StructureType, "", "", rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("InfoID"))
        If fso.FileExists(Server.MapPath(InfoPath & FileExt_Item)) Then
            fso.DeleteFile Server.MapPath(InfoPath & FileExt_Item)
        End If
        If ModuleType = 1 Then
            If rsDel("PaginationType") > 0 Then
                DelSerialFiles (Server.MapPath(InfoPath) & "_*" & FileExt_Item)
            End If
        End If

        rsDel("Deleted") = True
        rsDel.Update
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
End Sub

Sub AddGroupPurview(PurviewName, arrGroupID, iClassID)
    If arrGroupID = "" Then Exit Sub
    Dim sqlGroup, rsGroup
    sqlGroup = "select GroupID,GroupName,arrClass_" & PurviewName & " from PE_UserGroup where GroupID in (" & arrGroupID & ")"
    Set rsGroup = Server.CreateObject("ADODB.Recordset")
    rsGroup.Open sqlGroup, Conn, 1, 3
    Do While Not rsGroup.EOF
        rsGroup(2) = rsGroup(2) & "," & iClassID
        rsGroup.Update
        rsGroup.MoveNext
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
End Sub

Sub ModifyGroupPurview(PurviewName, arrGroupID, iClassID)
    Dim sqlGroup, rsGroup
    sqlGroup = "select GroupID,GroupName,arrClass_" & PurviewName & " from PE_UserGroup"
    Set rsGroup = Server.CreateObject("ADODB.Recordset")
    rsGroup.Open sqlGroup, Conn, 1, 3
    Do While Not rsGroup.EOF
        If FoundInArr(arrGroupID, rsGroup(0), ",") = True Then
            If FoundInArr(rsGroup(2), iClassID, ",") = False Then
                rsGroup(2) = rsGroup(2) & "," & iClassID
            End If
        Else
            rsGroup(2) = RemoveClassID(rsGroup(2), iClassID)
        End If
        rsGroup.Update
        rsGroup.MoveNext
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
End Sub

Function GetClassUrl(sParentDir, sClassDir, iClassID, iClassPurview)
    Dim strClassUrl
    If (UseCreateHTML = 1 Or UseCreateHTML = 3) And iClassPurview < 2 Then
        strClassUrl = ChannelUrl & GetListPath(StructureType, ListFileType, sParentDir, sClassDir) & GetListFileName(ListFileType, iClassID, 1, 1) & FileExt_List
    Else
        strClassUrl = ChannelUrl & "/ShowClass.asp?ClassID=" & iClassID
    End If
    GetClassUrl = strClassUrl
End Function

Function UpdateClassPurview(arrClassID)
    Dim rsClass, sqlClass, rsPurview, iClassPurview
    sqlClass = "select ClassPurview,ParentID,ParentPath,Child,arrChildID from PE_Class where ClassID in (" & arrClassID & ")"
    Set rsClass = Server.CreateObject("Adodb.recordset")
    rsClass.Open sqlClass, Conn, 1, 3
    Do While Not rsClass.EOF
        iClassPurview = rsClass("ClassPurview")
        If iClassPurview < 2 And rsClass("ParentID") > 0 Then
            Set rsPurview = Conn.Execute("select max(ClassPurview) from PE_Class where ClassID in (" & rsClass("ParentPath") & ")")
            If rsPurview(0) > iClassPurview Then iClassPurview = rsPurview(0)
            rsPurview.Close
            Set rsPurview = Nothing
            If iClassPurview > rsClass("ClassPurview") Then
                rsClass("ClassPurview") = iClassPurview
                rsClass.Update
            End If
        End If
        If iClassPurview > 0 And rsClass("Child") > 0 Then
            Conn.Execute ("update PE_Class set ClassPurview=" & iClassPurview & " where ClassID in (" & rsClass("arrChildID") & ") and ClassPurview<" & iClassPurview & "")
        End If
        rsClass.MoveNext
    Loop
    rsClass.Close
    Set rsClass = Nothing
End Function

Function GetChannel_Option(iModuleType, iChannelID)
    Dim rsGetAdmin, rsChannel
    Dim strChannel
    Set rsGetAdmin = Conn.Execute("select * from PE_Admin where AdminName='" & AdminName & "'")
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName,ChannelDir from PE_Channel  where ModuleType=" & iModuleType & " and Disabled=" & PE_False & " and ChannelType<=1 order by OrderID")
    Do While Not rsChannel.EOF
        If AdminPurview = 1 Or rsGetAdmin("AdminPurview_" & rsChannel("ChannelDir")) = 1 Then
            If rsChannel(0) = iChannelID Then
                strChannel = strChannel & "<option value='" & rsChannel(0) & "' selected>" & rsChannel(1) & "</option>"
            Else
                strChannel = strChannel & "<option value='" & rsChannel(0) & "'>" & rsChannel(1) & "</option>"
            End If
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
    GetChannel_Option = strChannel
End Function

Function GetUserGroup(PurviewName, iClassID, Cols)
    Dim rsGroup, strGroup, i
    strGroup = "<table width='100%'><tr>"
    Set rsGroup = Conn.Execute("select GroupID,GroupName,arrClass_" & PurviewName & " from PE_UserGroup order by GroupType asc,GroupID asc")
    Do While Not rsGroup.EOF
        strGroup = strGroup & "<td><input type='checkbox' name='arrGroupID_" & PurviewName & "' value='" & rsGroup(0) & "'"
        If iClassID > 0 Then
            If FoundInArr(rsGroup(2), iClassID, ",") = True Then
                strGroup = strGroup & " checked"
            End If
        End If
        strGroup = strGroup & ">" & rsGroup(1) & "</td>"
        i = i + 1
        rsGroup.MoveNext
        If i Mod Cols = 0 And Not rsGroup.EOF Then
            strGroup = strGroup & "</tr><tr>"
        End If
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
    strGroup = strGroup & "</table>"
    GetUserGroup = strGroup
End Function
%>

