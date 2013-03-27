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
Const PurviewLevel_Channel = 1   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim HtmlDir
Dim ManageType, InfoShortName

FileExt_SiteSpecial = arrFileExt(FileExt_SiteSpecial)

HtmlDir = InstallDir & ChannelDir
ManageType = Trim(Request("ManageType"))

Response.Write "<html><head><title>" & ChannelShortName & "专题管理</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
If ChannelID > 0 Then
     Call ShowPageTitle(ChannelName & "管理----专题管理", 10004)
Else
     Call ShowPageTitle("全站专题管理", 10004)
End If
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>"
Response.Write "    <td>"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "专题管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Add'>添加" & ChannelShortName & "专题</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Order'>" & ChannelShortName & "专题排序</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Unite'>合并" & ChannelShortName & "专题</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Batch'>批量设置</a>"
If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&ManageType=HTML'><b>生成HTML管理</b></a>"
End If
Response.Write "</td></tr></table>"

Action = Trim(Request("Action"))
Select Case Action
Case "Add"
    Call AddSpecial
Case "SaveAdd"
    Call SaveAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Del"
    Call DelSpecial
Case "Clear"
    Call ClearSpecial
Case "UpOrder"
    Call UpOrder
Case "DownOrder"
    Call DownOrder
Case "Unite"
    Call ShowUniteForm
Case "UniteSpecial"
    Call UniteSpecial
Case "Batch"
    Call ShowBatch
Case "DoBatch"
    Call DoBatch
Case "Order"
    Call ShowOrder
Case "CreateSpecialDir"
    Call CreateSpecialDir1
Case "CreateAllSpecialDir"
    Call CreateAllSpecialDir
Case "DelSpecialDir"
    Call DelSpecialDir1
Case "DelAllSpecialDir"
    Call DelAllSpecialDir
Case "CreateJS"
    Call CreateJS_Special
    Call WriteSuccessMsg("已经成功生成专题JS文件。", ComeUrl)
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 20
    strFileName = "Admin_Special.asp?ChannelID=" & ChannelID
    If Request("page") <> "" Then
        CurrentPage = PE_CLng(Request("page"))
    Else
        CurrentPage = 1
    End If

    Dim arrOpenType
    arrOpenType = Array("原窗口打开", "新窗口打开")

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>专题名称</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>专题目录</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>打开方式</strong></td>"
    Response.Write "    <td width='80' align='center'><strong>推荐专题</strong></td>"
    Response.Write "    <td width='200' align='center'><strong>专题提示</strong></td>"
    Response.Write "    <td width='100' height='22' align='center'><strong>常规操作</strong></td>"
    Response.Write "  </tr>"
    Dim rsSpecial, sql
    sql = "select * from PE_Special where ChannelID=" & ChannelID & " order by OrderID"
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 1
    If rsSpecial.BOF And rsSpecial.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>没有任何专题</td></tr>"
        totalPut = 0
    Else
        totalPut = rsSpecial.RecordCount
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
                rsSpecial.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim i
        i = 0
        Do While Not rsSpecial.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If ChannelID > 0 Then
                Response.Write "    <td align='center'><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ManageType=Special&SpecialID=" & rsSpecial("SpecialID") & "' title='点击进入管理此专题的" & InfoShortName & "'>" & rsSpecial("SpecialName") & "</a></td>"
            Else
                Response.Write "    <td align='center'>" & rsSpecial("SpecialName") & "</td>"
            End If
            Response.Write "    <td width='80' align='center'>" & rsSpecial("SpecialDir") & "</td>"
            Response.Write "    <td width='80' align='center'>" & arrOpenType(rsSpecial("OpenType")) & "</td>"
            Response.Write "    <td width='80' align='center'>"
            If rsSpecial("IsElite") = True Then
                Response.Write "<font color=green>是</font>"
            Else
                Response.Write "否"
            End If
            Response.Write "</td>"
            Response.Write "    <td width='200'>" & PE_HTMLEncode(rsSpecial("Tips")) & "</td>"
            If ManageType = "HTML" Then
                Response.Write "    <td width='240' align='center'>"
                Response.Write "<a href='Admin_Create" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=CreateSpecial&SpecialID=" & rsSpecial("SpecialID") & "' title='生成本专题的" & InfoShortName & "列表HTML页面'>生成列表页</a>&nbsp;|&nbsp;"
                Response.Write "<a href='" & HtmlDir & "/Special/" & rsSpecial("SpecialDir") & "/Index.html' title='查看本专题的" & InfoShortName & "列表HTML页面' target='_blank'>查看列表页</a>"
                If Not fso.FolderExists(Server.MapPath(HtmlDir & "/Special/" & rsSpecial("SpecialDir"))) Then
                    Response.Write "&nbsp;|&nbsp;<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=CreateSpecialDir&SpecialID=" & rsSpecial("SpecialID") & "' title='生成本专题的目录'>生成专题目录</a>"
                Else
                    Response.Write "&nbsp;|&nbsp;<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=DelSpecialDir&SpecialID=" & rsSpecial("SpecialID") & "' title='此操作将删除本专题的目录' onclick=""return confirm('此操作将删除本专题的目录，你可以重新生成目录。');"">删除专题目录</a>"
                End If
            Else
                Response.Write "    <td width='100' align='center'>"
                Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&action=Modify&SpecialID=" & rsSpecial("SpecialID") & "'>修改</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Del&SpecialID=" & rsSpecial("SpecialID") & "' onClick=""return confirm('确定要删除此专题吗？删除此专题后原属于此专题的" & InfoShortName & "将不属于任何专题。');"">删除</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=Clear&SpecialID=" & rsSpecial("SpecialID") & "' onClick=""return confirm('确定要清空此专题中的" & InfoShortName & "吗？本操作将原属于此专题的" & InfoShortName & "改为不属于任何专题。');"">清空</a>"
            End If
            Response.Write "</td></tr>"
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsSpecial.MoveNext
        Loop
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个专题", True)

    If ManageType = "HTML" Then
        Response.Write "<br><table align='center'><tr><form name='form1' action='Admin_Special.asp' method='post'><td>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='CreateAllSpecialDir'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='创建所有专题的目录' style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form><form name='form2' action='Admin_Create" & ModuleName & ".asp' method='post'><td><input name='CreateType' type='hidden' value='2'>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='CreateSpecial'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='生成所有专题的" & InfoShortName & "列表页' style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form><form name='form4' action='Admin_Special.asp' method='post'><td><input name='ManageType' type='hidden' value='HTML'>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='DelAllSpecialDir'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='删除所有专题的目录' onclick=""return confirm('此操作将删除所有专题的目录，你可以重新生成目录。如果你的系统中的专题列表文件发生混乱，可以使用此功能来删除所有目录，然后重新生成。');"" style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form></tr></table><br>"
        Response.Write "<b>注意：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;1、各项生成HTML操作之前，必须确保已经生成所有专题的目录。否则可能会导致生成出错。若专题目录为红色，表示此专题还没有创建相关的目录。请使用“生成专题目录”功能重新创建此专题的目录。"
        Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;2、因为生成操作会耗费大量的服务器资源，并可能需要相当长时间。<font color=red>在生成过程中千万不要刷新页面！！！</font>同时建议大家尽量在网站访问量比较小时进行。并尽量不要使用批量生成功能。"
    Else
        Response.Write "<table width='100%'><tr><form name='form1' action='Admin_Special.asp' method='post'><td align='center'>"
        Response.Write "<input name='Action' type='hidden' id='Action' value='CreateJS'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='submit' type='submit' value='刷新专题JS' style='cursor:hand;'"
        If ObjInstalled_FSO = False Then
            Response.Write " disabled"
        End If
        Response.Write "></td></form></tr></table>"
        If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
            Response.Write "<br><b>注意：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;若专题目录为红色，表示此专题还没有创建相关的目录。请到“生成HTML管理”页面使用“生成专题目录”功能重新创建此专题的目录。<br>"
        End If
    End If
End Sub

Sub ShowOrder()
    Dim iCount, i, j
    Dim rsSpecial, sql
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    sql = "select * from PE_Special where ChannelID=" & ChannelID & " Order by OrderID"
    rsSpecial.Open sql, Conn, 1, 1
    iCount = rsSpecial.RecordCount
    j = 1
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4' align='center'><strong>" & ChannelShortName & "专题排序</strong></td>"
    Response.Write "  </tr>"
    Do While Not rsSpecial.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""> "
        Response.Write "    <td align='center'>" & rsSpecial("SpecialName") & "</td>"
        Response.Write "    <form action='Admin_Special.asp?Action=UpOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If j > 1 Then
            Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>"
            For i = 1 To j - 1
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=SpecialID value=" & rsSpecial("SpecialID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsSpecial("OrderID") & ">&nbsp;<input type=submit name=Submit value=修改>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td></form>"
        Response.Write "    <form action='Admin_Special.asp?Action=DownOrder' method='post'>"
        Response.Write "      <td width='120' align='center'>"
        If iCount > j Then
            Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>"
            For i = 1 To iCount - j
                Response.Write "<option value=" & i & ">" & i & "</option>"
            Next
            Response.Write "</select><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
            Response.Write "<input type=hidden name=SpecialID value=" & rsSpecial("SpecialID") & ">"
            Response.Write "<input type=hidden name=cOrderID value=" & rsSpecial("OrderID") & ">&nbsp;<input type=submit name=Submit value=修改>"
        Else
            Response.Write "&nbsp;"
        End If
        Response.Write "</td>"
        Response.Write "      <td width='200' align='center'>&nbsp;</td>"
        Response.Write "    </form>"
        Response.Write "  </tr>"
        j = j + 1
        rsSpecial.MoveNext
    Loop
    Response.Write "</table> "
    rsSpecial.Close
    Set rsSpecial = Nothing
End Sub

Sub AddSpecial()

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
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Special.asp?ChannelID=" & ChannelID & "'>专题管理</a>&nbsp;&gt;&gt;&nbsp;添加专题</td></tr></table>"
    Response.Write "<form method='post' action='Admin_Special.asp' name='form1'>"

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本设置</td>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>自设内容</td>" & vbCrLf
    End If
    Response.Write "   <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
    Response.Write "   <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>专题名称：</strong></td>"
    Response.Write "      <td class='tdbg'><input name='SpecialName' type='text' id='SpecialName' size='49' maxlength='30'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>专题目录：</strong><br>只能是英文，不能带空格或“\”、“/”等符号。<br>本功能需要服务器支持FSO。但即使你的服务器不支持FSO，也请认真录入，因为可以在换了空间再批量生成。</td>"
    Response.Write "      <td class='tdbg'><input name='SpecialDir' type='text' id='SpecialDir' size='49' maxlength='30'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>专题图片：</strong></td>"
    Response.Write "      <td class='tdbg'><input name='SpecialPicUrl' type='text' id='SpecialPicUrl' size='49' maxlength='200'>&nbsp;</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>打开方式：</strong></td>"
    Response.Write "      <td><input name='OpenType' type='radio' value='0' checked>在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1'>在新窗口打开</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>是否为推荐专题：</strong></td>"
    Response.Write "      <td><input name='IsElite' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>专题提示：</strong><br>鼠标移至专题名称上时将显示设定的提示文字（不支持HTML）</td>"
    Response.Write "      <td class='tdbg'><textarea name='Tips' cols='60' rows='3' id='Tips'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>专题说明：</strong><br>用于专题页对专题进行说明（支持HTML）</td>"
    Response.Write "      <td class='tdbg'><textarea name='Readme' cols='60' rows='3' id='Readme'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>每页显示的" & InfoShortName & "数：</strong></td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, 20) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='350' class='tdbg5'><strong>默认配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td class='tdbg'><select name='SkinID'>" & GetSkin_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='350' class='tdbg5'><strong>版面设计模板：</strong><br>相关模板中包含了版面设计的版式等信息，如果是自行添加的设计模板，可能会导致“专题配色风格”失效。</td>"
    Response.Write "      <td class='tdbg'><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 4, 0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write " </tbody>" & vbCrLf
    If IsCustom_Content = True And ModuleType <> 6 Then
        Call EditCustom_Content("Add", "", "Special")
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input  type='submit' name='Submit' value=' 添 加 '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Special.asp'"" style='cursor:hand;'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim SpecialID, rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的专题ID！</li>"
        Exit Sub
    End If
    sql = "Select * from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的专题，可能已经被删除！</li>"
    Else
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
        Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Special.asp?ChannelID=" & ChannelID & "'>专题管理</a>&nbsp;&gt;&gt;&nbsp;修改专题设置：<font color='red'>" & rsSpecial("SpecialName") & "</td></tr></table>"
        Response.Write "<form method='post' action='Admin_Special.asp' name='form1'>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
        Response.Write "  <tr align='center' height='24'>" & vbCrLf
        Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本设置</td>" & vbCrLf
        If IsCustom_Content = True And ModuleType <> 6 Then
            Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>自设内容</td>" & vbCrLf
        End If
        Response.Write "   <td>&nbsp;</td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
        Response.Write "   <tbody id='Tabs' style='display:'>" & vbCrLf
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>专题名称：</strong></td>"
        Response.Write "      <td class='tdbg'><input name='SpecialName' type='text' id='SpecialName' value='" & rsSpecial("SpecialName") & "' size='49' maxlength='30'><input name='SpecialID' type='hidden' id='SpecialID' value='" & rsSpecial("SpecialID") & "'></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>专题目录：</strong></td>"
        Response.Write "      <td class='tdbg'><input name='SpecialDir' type='text' id='SpecialDir' value='" & rsSpecial("SpecialDir") & "' size='49' maxlength='30' disabled></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>专题图片：</strong></td>"
        Response.Write "      <td class='tdbg'><input name='SpecialPicUrl' type='text' id='SpecialPicUrl' value='" & rsSpecial("SpecialPicUrl") & "' size='49' maxlength='200'>&nbsp;</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>打开方式：</strong></td>"
        Response.Write "      <td><input name='OpenType' type='radio' value='0'  " & RadioValue(rsSpecial("OpenType"), 0) & ">在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1' " & RadioValue(rsSpecial("OpenType"), 1) & ">在新窗口打开</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>是否为推荐专题：</strong></td>"
        Response.Write "      <td><input name='IsElite' type='radio' value='True' " & RadioValue(rsSpecial("IsElite"), True) & ">是&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'" & RadioValue(rsSpecial("IsElite"), False) & ">否</td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>专题提示：</strong><br>鼠标移至专题名称上时将显示设定的提示文字（不支持HTML）</td>"
        Response.Write "      <td class='tdbg'><textarea name='Tips' cols='60' rows='3' id='Tips'>" & rsSpecial("Tips") & "</textarea></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>专题说明：</strong><br>用于专题页对专题进行说明（支持HTML）</td>"
        Response.Write "      <td class='tdbg'><textarea name='Readme' cols='60' rows='3' id='Readme'>" & rsSpecial("Readme") & "</textarea></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>每页显示的" & InfoShortName & "数：</strong></td>"
        Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, rsSpecial("MaxPerPage")) & "</select></td>"
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td width='350' class='tdbg5'><strong>默认配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
        Response.Write "      <td class='tdbg'><select name='SkinID'>" & GetSkin_Option(rsSpecial("SkinID")) & "</select></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='350' class='tdbg5'><strong>版面设计模板：</strong><br>相关模板中包含了版面设计的版式等信息，如果是自行添加的设计模板，可能会导致“专题配色风格”失效。</td>"
        Response.Write "      <td class='tdbg'><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 4, rsSpecial("TemplateID")) & "</select></td>"
        Response.Write "    </tr>"
        Response.Write " </tbody>" & vbCrLf
        If IsCustom_Content = True And ModuleType <> 6 Then
            Call EditCustom_Content("Modify", rsSpecial("Custom_Content"), "Special")
        End If
        Response.Write "    <tr class='tdbg'> "
        Response.Write "      <td colspan='2' align='center' class='tdbg'><input name='Action' type='hidden' id='Action' value='SaveModify'>"
        Response.Write "        <input  type='submit' name='Submit' value='保存修改结果'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>&nbsp; "
        Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Special.asp'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
End Sub

Sub ShowUniteForm()
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'> "
    Response.Write "    <td height='22' colspan='3' align='center'><strong>合并" & ChannelShortName & "专题</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='100'><form name='myform' method='post' action='Admin_Special.asp' onSubmit='return ConfirmUnite();'>"
    Response.Write "        &nbsp;&nbsp;将专题 <select name='SpecialID' id='SpecialID'>" & GetSpecial_Option(0) & "</select> 合并到 <select name='TargetSpecialID' id='TargetSpecialID'>" & GetSpecial_Option(0) & "</select>"
    Response.Write "        <br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='UniteSpecial'>"
    Response.Write "        <input type='submit' name='Submit' value=' 合并专题 ' style='cursor:hand;'>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Special.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'>"
    Response.Write "      </form></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='60'><strong>注意事项：</strong><br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;所有操作不可逆，请慎重操作！！！<br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;不能在同一个专题内进行操作。<br>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;合并后您所指定的专题将被删除，所有" & InfoShortName & "将转移到目标专题中。</td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function ConfirmUnite(){" & vbCrLf
    Response.Write "  if (document.myform.SpecialID.value==document.myform.TargetSpecialID.value){" & vbCrLf
    Response.Write "    alert('请不要在相同专题内进行操作！');" & vbCrLf
    Response.Write " document.myform.TargetSpecialID.focus();" & vbCrLf
    Response.Write " return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub ShowBatch()
    Response.Write "<form name='form1' method='post' action='Admin_Special.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='3' align='center'><strong>批量设置" & ChannelShortName & "专题属性</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' valign='top'><font color='red'>提示：</font>可以按住“Shift”<br>或“Ctrl”键进行多个专题的选择<br>"
    Response.Write "      <select name='SpecialID' size='2' multiple style='height:200px;width:200px;'>" & GetSpecial_Option(0) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  选定所有专题  ' onclick='SelectAll()'><br>"
    Response.Write "      <input type='button' name='Submit' value='取消选定所有专题' onclick='UnSelectAll()'></div></td>"
    Response.Write "      <td>"
    Response.Write "     <table id='SpecialSettings' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF' style='display:'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyOpenType' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>打开方式：</strong></td>"
    Response.Write "      <td><input name='OpenType' type='radio' value='0' checked>在原窗口打开&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name='OpenType' type='radio' value='1'>在新窗口打开</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyIsElite' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>是否为推荐专题：</strong></td>"
    Response.Write "      <td><input name='IsElite' type='radio' value='True' checked>是&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='radio' value='False'>否</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyMaxPerPage' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>每页显示的" & InfoShortName & "数：</strong></td>"
    Response.Write "      <td><select name='MaxPerPage'>" & GetNumber_Option(5, 100, 20) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifySkinID' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>专题配色风格：</strong><br>相关模板中包含CSS、颜色、图片等信息</td>"
    Response.Write "      <td><select name='SkinID' id='SkinID'>" & GetSkin_Option(0) & "</td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td width='30' align='center'><input type='checkbox' name='ModifyTemplateID' value='Yes'></td>"
    Response.Write "      <td width='300' class='tdbg5'><strong>版面设计模板：</strong><br>相关模板中包含了专题设计的版式等信息，如果是自行添加的设计模板，可能会导致“专题配色风格”失效。 </td>"
    Response.Write "      <td><select name='TemplateID' id='TemplateID'>" & GetTemplate_Option(ChannelID, 4, 0) & "</select></td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='3' align='center'><input name='Action' type='hidden' id='Action' value='DoBatch'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Submit' type='submit' value=' 执行批处理 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Special.asp?ChannelID=" & ChannelID & "'"" style='cursor:hand;'></td></tr>"
    Response.Write "  </table>"
    Response.Write "</td></tr></table>"
    Response.Write "</form>" & vbCrLf
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.form1.SpecialID.length;i++){" & vbCrLf
    Response.Write "    document.form1.SpecialID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.form1.SpecialID.length;i++){" & vbCrLf
    Response.Write "    document.form1.SpecialID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub SaveAdd()
    Dim SpecialName, SpecialDir, SpecialID, OrderID
    Dim rsSpecial, sql
    SpecialName = ReplaceBadChar(Trim(Request("SpecialName")))
    SpecialDir = Trim(Request("SpecialDir"))
    If SpecialName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>专题名称不能为空！</li>"
    End If
    If SpecialDir = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>专题目录不能为空！</li>"
    Else
        If IsValidStr(SpecialDir) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>专题目录名只能是英文！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    SpecialID = GetNewID("PE_Special", "SpecialID")
    OrderID = GetMinID("PE_Special", "OrderID")
    Conn.Execute ("update PE_Special set OrderID=OrderID+1 where ChannelID=" & ChannelID & "")
    
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open "Select * from PE_Special Where ChannelID=" & ChannelID & " and (SpecialName='" & SpecialName & "' or SpecialDir='" & SpecialDir & "')", Conn, 1, 3
    If Not (rsSpecial.BOF And rsSpecial.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>专题名称或专题目录已经存在！</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
        Exit Sub
    End If
    
    rsSpecial.addnew
    rsSpecial("SpecialID") = SpecialID
    rsSpecial("ChannelID") = ChannelID
    rsSpecial("OrderID") = OrderID
    rsSpecial("SpecialName") = SpecialName
    rsSpecial("SpecialDir") = SpecialDir
    rsSpecial("SpecialPicUrl") = Trim(Request("SpecialPicUrl"))
    rsSpecial("IsElite") = PE_CBool(Trim(Request("IsElite")))
    rsSpecial("OpenType") = PE_CLng(Trim(Request("OpenType")))
    rsSpecial("Tips") = Trim(Request("Tips"))
    rsSpecial("ReadMe") = Trim(Request("ReadMe"))
    rsSpecial("MaxPerPage") = PE_CLng(Trim(Request("MaxPerPage")))
    rsSpecial("SkinID") = PE_CLng(Trim(Request("SkinID")))
    rsSpecial("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
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
    rsSpecial("Custom_Content") = Custom_Content
    rsSpecial.Update
    rsSpecial.Close
    Set rsSpecial = Nothing
    Conn.Execute ("update PE_Channel set SpecialCount=SpecialCount+1 where ChannelID=" & ChannelID & "")
    Call CreateJS_Special
    If UseCreateHTML > 0 Then
        Call CreateSpecialDir(SpecialDir)
    End If
    Call ClearSiteCache(ChannelID)
    Call CloseConn
    Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
End Sub

Sub CreateSpecialDir(DirName)
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim tmpDir
    tmpDir = InstallDir & ChannelDir & "/Special"
    If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.CreateFolder Server.MapPath(tmpDir)
    End If
    tmpDir = tmpDir & "/" & DirName
    If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.CreateFolder Server.MapPath(tmpDir)
    End If
End Sub

Sub CreateSpecialDir1()
    Dim SpecialID, rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的专题ID！</li>"
        Exit Sub
    End If
    sql = "Select SpecialDir from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的专题，可能已经被删除！</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
    Else
        Call CreateSpecialDir(rsSpecial(0))
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    If FoundErr = False Then Call WriteSuccessMsg("创建专题目录成功！", ComeUrl)
End Sub

Sub SaveModify()
    Dim SpecialID, SpecialName
    Dim rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的专题ID！</li>"
        Exit Sub
    End If
    SpecialName = ReplaceBadChar(Trim(Request.Form("SpecialName")))
    If SpecialName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>专题名称不能为空！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    sql = "Select * from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的专题，可能已经被删除！</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
    Else
        rsSpecial("SpecialName") = SpecialName
        rsSpecial("SpecialPicUrl") = Trim(Request("SpecialPicUrl"))
        rsSpecial("IsElite") = PE_CBool(Trim(Request("IsElite")))
        rsSpecial("OpenType") = PE_CLng(Trim(Request("OpenType")))
        rsSpecial("Tips") = Trim(Request("Tips"))
        rsSpecial("ReadMe") = Trim(Request("ReadMe"))
        rsSpecial("MaxPerPage") = PE_CLng(Trim(Request("MaxPerPage")))
        rsSpecial("SkinID") = PE_CLng(Trim(Request("SkinID")))
        rsSpecial("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
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
        rsSpecial("Custom_Content") = Custom_Content
        rsSpecial.Update
        rsSpecial.Close
        Set rsSpecial = Nothing
        Call CreateJS_Special
        Call ClearSiteCache(ChannelID)
        Call CloseConn
        Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
    End If
End Sub

Sub DelSpecial()
    Dim SpecialID
    SpecialID = Trim(Request("SpecialID"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的专题ID！</li>"
        Exit Sub
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If UseCreateHTML > 0 Then
        Dim trs, SpecialDir
        Set trs = Conn.Execute("select SpecialDir from PE_Special where SpecialID=" & SpecialID)
        SpecialDir = trs(0)
        Set trs = Nothing
        Call DelSpecialDir(SpecialDir)
    End If
    If FoundErr = True Then Exit Sub

    Dim rsInfo
    Set rsInfo = Conn.Execute("select * from PE_InfoS where SpecialID=" & SpecialID & " order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & rsInfo("ModuleType") & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 1 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=0 where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing

    Conn.Execute ("delete from PE_Special where SpecialID=" & SpecialID)
    Conn.Execute ("update PE_Channel set SpecialCount=SpecialCount-1 where ChannelID=" & ChannelID & "")
    Call CreateJS_Special
    Call CloseConn
    Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
End Sub

Sub DelSpecialDir(DirName)
    On Error Resume Next
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim tmpDir
    tmpDir = InstallDir & ChannelDir & "/Special/" & DirName
    If fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.DeleteFolder Server.MapPath(tmpDir)
    End If
    If Err Then
        Error.Clear
        FoundErr = True
        ErrMsg = ErrMsg & "<li>专题目录无法删除！可能有文件正在使用中。请稍后再试！</li>"
    End If
End Sub

Sub DelSpecialDir1()
    Dim SpecialID, rsSpecial, sql
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    If SpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的专题ID！</li>"
        Exit Sub
    End If
    sql = "Select SpecialDir from PE_Special Where SpecialID=" & SpecialID
    Set rsSpecial = Server.CreateObject("Adodb.RecordSet")
    rsSpecial.Open sql, Conn, 1, 3
    If rsSpecial.BOF And rsSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的专题，可能已经被删除！</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
    Else
        Call DelSpecialDir(rsSpecial(0))
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing

    If FoundErr = False Then Call WriteSuccessMsg("删除专题目录成功！", ComeUrl)
End Sub

Sub ClearSpecial()
    Dim SpecialID
    SpecialID = Trim(Request("SpecialID"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的专题ID！</li>"
        Exit Sub
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If UseCreateHTML > 0 Then
        Dim trs, SpecialDir
        Set trs = Conn.Execute("select SpecialDir from PE_Special where SpecialID=" & SpecialID)
        SpecialDir = trs(0)
        Set trs = Nothing
        Call ClearSpecialDir(SpecialDir)
    End If
    If FoundErr = True Then Exit Sub

    Dim rsInfo
    Set rsInfo = Conn.Execute("select * from PE_InfoS where SpecialID=" & SpecialID & " order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & rsInfo("ModuleType") & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 1 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=0 where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing
    
    Call CloseConn
    Response.Redirect "Admin_Special.asp?ChannelID=" & ChannelID
End Sub

Sub ClearSpecialDir(DirName)
    On Error Resume Next
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim tmpDir
    tmpDir = InstallDir & ChannelDir & "/Special/" & DirName
    If fso.FolderExists(Server.MapPath(tmpDir)) Then
        fso.DeleteFile Server.MapPath(tmpDir) & "\*.*"
    End If
    If Err Then
        Error.Clear
        FoundErr = True
        ErrMsg = ErrMsg & "<li>无法完全清除此专题目录下的文件！可能有文件正在使用中。请稍后再试！</li>"
    End If
End Sub

Sub UpOrder()
    Dim SpecialID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsSpecial
    SpecialID = Trim(Request("SpecialID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        SpecialID = PE_CLng(SpecialID)
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
    Set mrs = Conn.Execute("select max(OrderID) from PE_Special")
    MaxOrderID = mrs(0) + 1
    '先将当前专题移至最后，包括子专题
    Conn.Execute ("update PE_Special set OrderID=" & MaxOrderID & " where SpecialID=" & SpecialID)
    
    '然后将位于当前专题以上的专题的OrderID依次加一，范围为要提升的数字
    sqlOrder = "select * from PE_Special where OrderID<" & cOrderID & " order by OrderID desc"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前专题已经在最上面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID，包括子专题
        Conn.Execute ("update PE_Special set OrderID=OrderID+1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前专题从最后移到相应位置，包括子专题
    Conn.Execute ("update PE_Special set OrderID=" & tOrderID & " where SpecialID=" & SpecialID)
    Call CreateJS_Special
    Call CloseConn
    Response.Redirect "Admin_Special.asp?Action=Order&ChannelID=" & ChannelID
End Sub

Sub DownOrder()
    Dim SpecialID, sqlOrder, rsOrder, MoveNum, cOrderID, tOrderID, i, rsSpecial, PrevID, NextID
    SpecialID = Trim(Request("SpecialID"))
    cOrderID = Trim(Request("cOrderID"))
    MoveNum = Trim(Request("MoveNum"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    Else
        SpecialID = PE_CLng(SpecialID)
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
    Set mrs = Conn.Execute("select max(OrderID) from PE_Special")
    MaxOrderID = mrs(0) + 1
    '先将当前专题移至最后，包括子专题
    Conn.Execute ("update PE_Special set OrderID=" & MaxOrderID & " where SpecialID=" & SpecialID)
    
    '然后将位于当前专题以下的专题的OrderID依次减一，范围为要下降的数字
    sqlOrder = "select * from PE_Special where OrderID>" & cOrderID & " order by OrderID"
    Set rsOrder = Server.CreateObject("adodb.recordset")
    rsOrder.Open sqlOrder, Conn, 1, 3
    If rsOrder.BOF And rsOrder.EOF Then
        Exit Sub        '如果当前专题已经在最下面，则无需移动
    End If
    i = 1
    Do While Not rsOrder.EOF
        tOrderID = rsOrder("OrderID")     '得到要提升位置的OrderID，包括子专题
        Conn.Execute ("update PE_Special set OrderID=OrderID-1 where OrderID=" & tOrderID)
        i = i + 1
        If i > MoveNum Then
            Exit Do
        End If
        rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    
    '然后再将当前专题从最后移到相应位置，包括子专题
    Conn.Execute ("update PE_Special set OrderID=" & tOrderID & " where SpecialID=" & SpecialID)
    Call CreateJS_Special
    Call CloseConn
    Response.Redirect "Admin_Special.asp?Action=Order&ChannelID=" & ChannelID
End Sub

Sub UniteSpecial()
    Dim SpecialID, TargetSpecialID, SuccessMsg
    SpecialID = Trim(Request("SpecialID"))
    TargetSpecialID = Trim(Request("TargetSpecialID"))
    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要合并的专题！</li>"
    Else
        SpecialID = PE_CLng(SpecialID)
    End If
    If TargetSpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定目标专题！</li>"
    Else
        TargetSpecialID = PE_CLng(TargetSpecialID)
    End If
    If SpecialID = TargetSpecialID Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请不要在相同专题内进行操作</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Dim rsInfo
    Set rsInfo = Conn.Execute("select * from PE_InfoS where SpecialID=" & SpecialID & " order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & rsInfo("ModuleType") & " and SpecialID=" & TargetSpecialID & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 0 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=" & TargetSpecialID & " where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing
    

    '删除被合并专题
    Conn.Execute ("delete from PE_Special where SpecialID=" & SpecialID)
    Conn.Execute ("update PE_Channel set SpecialCount=SpecialCount-1 where ChannelID=" & ChannelID & "")
    SuccessMsg = "专题合并成功！已经将被合并专题的所有数据转入目标专题中。"
    Call CreateJS_Special
    Call WriteSuccessMsg(SuccessMsg, ComeUrl)
End Sub

Sub CreateAllSpecialDir()
    'On Error Resume Next
    If Not fso.FolderExists(Server.MapPath(HtmlDir & "/Special")) Then
        fso.CreateFolder Server.MapPath(HtmlDir & "/Special")
    End If

    Dim sqlSpecial, rsSpecial, i, iDepth
    sqlSpecial = "select * from PE_Special where ChannelID=" & ChannelID & " order by OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    Do While Not rsSpecial.EOF
        If Not fso.FolderExists(Server.MapPath(HtmlDir & "/Special/" & rsSpecial("SpecialDir"))) Then
            fso.CreateFolder Server.MapPath(HtmlDir & "/Special/" & rsSpecial("SpecialDir"))
        End If
        rsSpecial.MoveNext
    Loop
    rsSpecial.Close
    Set rsSpecial = Nothing
    Call WriteSuccessMsg("生成所有专题的目录成功！", ComeUrl)
End Sub

Sub DelAllSpecialDir()
    On Error Resume Next
    Dim theFolder, theSubFolder, strFolderName
    Set theFolder = fso.GetFolder(Server.MapPath(HtmlDir & "/Special"))
    For Each theSubFolder In theFolder.SubFolders
        strFolderName = theSubFolder.name
        theSubFolder.Delete
        If Err Then
            Err.Clear
            FoundErr = True
            ErrMsg = ErrMsg & "<li>删除目录：" & strFolderName & "失败！可能当前目录正在使用中。请稍后再试！</li>"
        End If
    Next
    If FoundErr <> True Then
        Call WriteSuccessMsg("删除所有栏目的目录成功！", ComeUrl)
    End If
End Sub

Sub CreateJS_Special()

    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    Dim hf, strSpecial, SpecialPath
    '全站专题
    If ChannelID = 0 Then
        SpecialPath = InstallDir & "js"
        JS_SpecialNum = 10
    Else
        SpecialPath = InstallDir & ChannelDir & "/js"
    End If

    If Not fso.FolderExists(Server.MapPath(SpecialPath)) Then
        fso.CreateFolder (Server.MapPath(SpecialPath))
    End If
    
    strSpecial = GetSpecialList(JS_SpecialNum)
    Call WriteToFile(SpecialPath & "/ShowSpecialList.js", "document.write(""" & strSpecial & """);")
End Sub

Sub DoBatch()
    Dim SpecialID
    Dim sql, rsSpecial, i, trs
    SpecialID = Trim(Request("SpecialID"))
	If IsValidID(SpecialID) = False Then
		SpecialID = ""
	End If

    If SpecialID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定要批量修改设置的专题！</li>"
    End If


    If FoundErr = True Then
        Exit Sub
    End If
    
    If InStr(SpecialID, ",") > 0 Then
        SpecialID = ReBuild(SpecialID)
        sql = "select * from PE_Special where SpecialID in (" & SpecialID & ")"
    Else
        sql = "select * from PE_Special where SpecialID=" & SpecialID
    End If
    Set rsSpecial = Server.CreateObject("Adodb.recordset")
    rsSpecial.Open sql, Conn, 1, 3
    Do While Not rsSpecial.EOF
        If Trim(Request("ModifyIsElite")) = "Yes" Then rsSpecial("IsElite") = PE_CBool(Trim(Request("IsElite")))
        If Trim(Request("ModifyOpenType")) = "Yes" Then rsSpecial("OpenType") = PE_CLng(Trim(Request("OpenType")))
        If Trim(Request("ModifyMaxPerPage")) = "Yes" Then rsSpecial("MaxPerPage") = PE_CLng(Trim(Request("MaxPerPage")))
        If Trim(Request("ModifySkinID")) = "Yes" Then rsSpecial("SkinID") = PE_CLng(Trim(Request("SkinID")))
        If Trim(Request("ModifyTemplateID")) = "Yes" Then rsSpecial("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
        rsSpecial.Update
        rsSpecial.MoveNext
        Set trs = Nothing
    Loop
    rsSpecial.Close
    Set rsSpecial = Nothing
    Call ClearSiteCache(ChannelID)
    Call CreateJS_Special
    Call WriteSuccessMsg("批量设置专题属性成功！", ComeUrl)
End Sub

Function ReBuild(ByVal iSpecialID)
    Dim arrSpecialID, SpecialArr, i
    arrSpecialID = Split(iSpecialID, ",")
    SpecialArr = ""
    For i = 0 To UBound(arrSpecialID)
        If Trim(arrSpecialID(i)) <> "" And Trim(arrSpecialID(i)) <> "0" Then
            If SpecialArr = "" Then
                SpecialArr = arrSpecialID(i)
            Else
                SpecialArr = SpecialArr & "," & arrSpecialID(i)
            End If
        End If
    Next
    ReBuild = SpecialArr
End Function



'=================================================
'函数名：GetSpecialList
'作  用：以竖向列表方式显示专题名称
'参  数：SpecialNum  ------最多显示多少个专题名称
'=================================================
Function GetSpecialList(SpecialNum)
    Dim sqlSpecial, rsSpecial, strSpecial, i
    If SpecialNum <= 0 Or SpecialNum > 100 Then
        SpecialNum = 10
    End If
    sqlSpecial = "select SpecialID,SpecialName,SpecialDir,Tips from PE_Special where ChannelID=" & ChannelID & " and IsElite=" & PE_True & " order by OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    If rsSpecial.BOF And rsSpecial.EOF Then
        strSpecial = "&nbsp;没有任何专题栏目"
    Else

        i = 0
        Do While Not rsSpecial.EOF
            If ChannelID = 0 Then
                If FileExt_SiteSpecial <> ".asp" Then
                    strSpecial = strSpecial & "<li><a href='" & InstallDir & "Special/" & rsSpecial(2) & "/Index" & FileExt_SiteSpecial & "' title='" & rsSpecial(3) & "'>" & rsSpecial(1) & "</a></li>"
                Else
                    strSpecial = strSpecial & "<li><a href='" & InstallDir & "ShowSpecial.asp?SpecialID=" & rsSpecial(0) & "' title='" & Trim(nohtml(rsSpecial(3))) & "'>" & rsSpecial(1) & "</a></li>"
                End If
            Else
                If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
                    strSpecial = strSpecial & "<li><a href='" & ChannelUrl & "/Special/" & rsSpecial(2) & "/Index" & FileExt_List & "' title='" & rsSpecial(3) & "'>" & rsSpecial(1) & "</a></li>"
                Else
                    strSpecial = strSpecial & "<li><a href='" & ChannelUrl & "/ShowSpecial.asp?SpecialID=" & rsSpecial(0) & "' title='" & Trim(nohtml(rsSpecial(3))) & "'>" & rsSpecial(1) & "</a></li>"
                End If
            End If

            rsSpecial.MoveNext
            i = i + 1
            If i >= SpecialNum Then Exit Do
        Loop
    End If
    If Not rsSpecial.EOF Then
        strSpecial = strSpecial & "<p align='right'><a href='" & ChannelUrl & "/SpecialList.asp'>更多专题</a></p>"
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    GetSpecialList = strSpecial
End Function
%>
