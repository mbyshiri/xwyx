<!--#include file="Admin_Common.asp"-->
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
Const PurviewLevel_Others = "Template"   '其他权限

strFileName = "Admin_TemplateProject.asp?Action=" & Action

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>网站模板方案管理</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

If Action <> "TemplateProject" Then
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
    Call ShowPageTitle("网站模板方案管理", 10005)
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""70"" height=""30""><strong>管理导航：</strong></td>" & vbCrLf
    Response.Write "    <td height=""30""><a href=Admin_TemplateProject.asp?Action=Main>管理首页</a> | <a href=""Admin_TemplateProject.asp?Action=AddProject"">添加新模板方案项目</a> | <a href=""Admin_TemplateProject.asp?Action=Import"">导入模板方案</a> | <a href=""Admin_TemplateProject.asp?Action=Export"">导出模板方案</a> | <a href=""Admin_TemplateProject.asp?Action=TemplateBatchMove"">方案间模板迁移 </a> | <a href=""Admin_TemplateProject.asp?Action=SkinBatchMove"">方案间风格迁移</a> | </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
End If

Select Case Action
Case "AddProject", "ModifyProject"
    Call AddProject
Case "SaveAdd", "SaveModify"
    Call SaveProject
Case "Del"
    Call DelTemplateProject
Case "Del2"
    Call DelTemplateProject2
Case "Set"
    Call SetDefault
Case "Import"                   '项目导入第一步
    Call Import
Case "Import2"                  '项目导入第二步
    Call Import2
Case "DoImport"                 '导入项目处理
    Call DoImport
Case "Export"                   '导出方案
    Call Export
Case "DoExport"                 '导出方案处理
    Call DoExport
Case "TemplateBatchMove"                '模板批量迁移
    Call TemplateBatchMove
Case "DoTemplateBatchMove"              '模板批量迁移处理
    Call DoTemplateBatchMove
Case "SkinBatchMove"                    '风格批量迁移
    Call SkinBatchMove
Case "DoSkinBatchMove"                  '风格批量迁移处理
    Call DoSkinBatchMove
Case "TemplateProject"
    Call TemplateProject
Case Else
        Call main
End Select
Response.Write "</body></html>"
Call CloseConn


'=================================================
'过程名：main
'作  用：管理项目
'=================================================
Sub main()
    Dim rs, sql, sysIsDefault

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_TemplateProject.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "  <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "      <td width='50' align='center'><strong>选中</strong></td>"
    Response.Write "      <td align='center' width='80'><strong>方案名称</strong></td>"
    Response.Write "      <td align='center' width='200'><strong>方案简介</strong></td>"
    Response.Write "      <td width='60' align='center'><strong>是否默认</strong></td>"
    Response.Write "      <td width='240' height='22' align='center'><strong> 方案管理</strong></td>"
    Response.Write "      <td width='200' height='22' align='center'><strong> 方案操作</strong></td>"
    Response.Write "  </tr>"

    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' align='center' height='50'><br>还没有模板方案！<br><br></td></tr>"
    Else

        Do While Not rs.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "      <td width='50' align='center' height=""30"">" & rs("TemplateProjectID") & "</td>"
            Response.Write "      <td align='center' width='80'>" & rs("TemplateProjectName") & "</td>"
            Response.Write "      <td align='center' width='200'>" & rs("Intro") & "</td>"
            Response.Write "      <td width='60' align='center'>"

            If rs("IsDefault") = True Then
                Response.Write "<b>√</b>"
            End If

            Response.Write "</td>"
            Response.Write " <td align='center' width='240'>"
            Response.Write " <a href='Admin_Template.asp?Action=Main&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "' >管理该方案下的模板</a>" & vbCrLf
            Response.Write " <a href='Admin_Skin.asp?Action=main&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "' >管理该方案下的风格</a>" & vbCrLf
            Response.Write "</td>"

            Response.Write "      <td width='200' align='center'><a href='Admin_TemplateProject.asp?Action=ModifyProject&TemplateProjectID=" & rs("TemplateProjectID") & "'>修改方案</a>&nbsp;&nbsp;"

            If rs("IsDefault") = False Then
                Response.Write "<a href='Admin_TemplateProject.asp?Action=Del&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "' onClick=""return confirm('确定要删除此方案吗？删除此方案后方案隶属的模板,风格 都将会被删除,请严格注意!');"">删除方案</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_TemplateProject.asp?Action=Set&TemplateProjectID=" & rs("TemplateProjectID") & "&ProjectName=" & Server.UrlEncode(rs("TemplateProjectName")) & "'  onClick=""return confirm('您确定该方案的模板和风格都有默认数据了么,如果没有请先添加或方案迁移!');"">设为默认</a>"
            Else
                Response.Write "<font color='gray'>删除方案&nbsp;&nbsp;设为默认</font>"
            End If

            Response.Write "      </td>"
            Response.Write "    </tr>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing

    Response.Write "</table>"
    Response.Write "</form></tr></table>"
End Sub

'=================================================
'过程名：AddProject
'作  用：添加项目
'=================================================
Sub AddProject()
        
    '变量声明区域 请填写
    Dim rsItem, sql, TemplateProjectID
    Dim SaveType, SaveName
    Dim TemplateProjectName, Intro, IsDefault
    Dim iTemplateType, i, Num
    Dim SkinID

    '变量获取区 请填写
    TemplateProjectID = PE_CLng(Request("TemplateProjectID"))
    FoundErr = False
    SaveType = "SaveAdd"
    SaveName = " 添 加 "

    '是否是修改
    If TemplateProjectID > 0 Then
        SaveType = "SaveModify"
        SaveName = " 修 改 "
        '取出数据
        sql = "select TemplateProjectID,TemplateProjectName,Intro,IsDefault from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
        Set rsItem = Server.CreateObject("adodb.recordset")
        rsItem.Open sql, Conn, 1, 1

        If rsItem.EOF Then   '没有找到该项目
            FoundErr = True
            ErrMsg = ErrMsg & "<li>错误参数！没有找到该方案！</li>"
        Else
            TemplateProjectID = rsItem("TemplateProjectID")
            TemplateProjectName = rsItem("TemplateProjectName")
            Intro = rsItem("Intro")
            IsDefault = rsItem("IsDefault")
        End If

        rsItem.Close
        Set rsItem = Nothing
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<script language = ""JavaScript"">" & vbCrLf
    Response.Write "    function CheckForm(){" & vbCrLf
    Response.Write "        if (document.myform.TemplateProjectName.value==""""){" & vbCrLf
    Response.Write "            alert(""方案名称不能为空！"");" & vbCrLf
    Response.Write "            document.myform.TemplateProjectName.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (document.myform.Intro.value==""""){" & vbCrLf
    Response.Write "            alert(""方案简介不能为空！"");" & vbCrLf
    Response.Write "            document.myform.Intro.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<FORM name=myform action='Admin_TemplateProject.asp' method=post>" & vbCrLf
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    Response.Write "      <td height='22' colspan='2'><strong> " & SaveName & " 方 案</strong></td>"
    Response.Write "    </tr>"
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "     <td class='tdbg'  valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr class='tdbg'> " & vbCrLf '文本
    Response.Write "          <td width='150' class='tdbg5' align='right' ><strong> 方案名称：&nbsp;</strong></td>" & vbCrLf
    Response.Write "          <td class='tdbg'>" & vbCrLf
    Response.Write "            <input name='TemplateProjectName' type='text' id='TemplateProjectName' size='30' maxlength='30' value='" & TemplateProjectName & "'>" & vbCrLf
    Response.Write "            <font color=red> * </font>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'> " & vbCrLf '文本框
    Response.Write "          <td width='150' class='tdbg5' align='right'><strong> 方案简介：&nbsp;</strong></td>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <textarea name='Intro' style='width:450px;height:100px' id='Intro'>" & PE_ConvertBR(Intro) & "</textarea>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "     </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<center>" & vbCrLf
    Response.Write "  <Input id='TemplateProjectID' type='hidden' value=" & TemplateProjectID & " name='TemplateProjectID'>" & vbCrLf
    Response.Write "  <Input id='Action' type='hidden' value='" & SaveType & "' name='Action'>" & vbCrLf
    Response.Write "  <Input type='submit' value=' 确 定 ' name='Submit' onClick=""return CheckForm();"">&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "  <Input type='Reset' name='Reset' value=' 清 除 '>" & vbCrLf
    Response.Write "</center>" & vbCrLf
    Response.Write "</FORM>" & vbCrLf

End Sub

'=================================================
'过程名：Save
'作  用：保存项目
'=================================================
Sub SaveProject()
    '变量声明区
    Dim TemplateProjectName, Intro, SaveName, TemplateProjectID
    Dim rsItem, rsModify, mrs, sql

    '变量获取区
    TemplateProjectID = PE_CLng(Request("TemplateProjectID"))
    TemplateProjectName = Replace(ReplaceBadChar(ReplaceText(Trim(Request("TemplateProjectName")), 2)), "nbsp", "")
    Intro = Trim(Request("Intro"))

    '变量检测区
    If TemplateProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>方案标题不能为空！</li>"
    End If

    If Len(TemplateProjectName) > 250 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>方案标题过长（应小于250）！</li>"
    End If

    If Intro = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>方案简介不能为空！</li>"
    End If
        
    sql = "Select TemplateProjectName From PE_TemplateProject Where TemplateProjectName='" & TemplateProjectName & "'"
    Set rsItem = Server.CreateObject("Adodb.Recordset")
    rsItem.Open sql, Conn, 1, 1

    If rsItem.EOF And rsItem.BOF Then
    Else

        If Action = "SaveModify" Then
            sql = "select * from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
            Set rsModify = Server.CreateObject("Adodb.Recordset")
            rsModify.Open sql, Conn, 1, 3

            If rsModify.BOF And rsModify.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到指定的方案！</li>"
            Else

                If TemplateProjectName <> rsModify("TemplateProjectName") Then
                    FoundErr = True
                End If
            End If

            rsModify.Close
            Set rsModify = Nothing
        Else
            FoundErr = True
        End If

        ErrMsg = ErrMsg & "<li>方案管理中已经有相应的方案名称,请返回重新输入名称！</li>"
    End If

    rsItem.Close
    Set rsItem = Nothing

    '这里根据需要填写逻辑处理
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    TemplateProjectName = PE_HTMLEncode(TemplateProjectName)
    Intro = PE_HTMLEncode(Intro)
        
    If FoundErr <> True Then
        '数据存储区
        Set rsItem = Server.CreateObject("adodb.recordset")

        If Action = "SaveAdd" Then
            SaveName = "添加"
            Set mrs = Conn.Execute("select max(TemplateProjectID) from PE_TemplateProject")

            If IsNull(mrs(0)) Then
                TemplateProjectID = 1
            Else
                TemplateProjectID = mrs(0) + 1
            End If

            Set mrs = Nothing
            sql = "select top 1 * from PE_TemplateProject"
            rsItem.Open sql, Conn, 1, 3
            rsItem.addnew
            rsItem("TemplateProjectID") = TemplateProjectID
        ElseIf Action = "SaveModify" Then
            SaveName = "修改"

            If TemplateProjectID = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>不能确定方案的ID!</li>"
                Exit Sub
            Else
                sql = "select * from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
                rsItem.Open sql, Conn, 1, 3

                If rsItem.BOF And rsItem.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>找不到指定的方案！</li>"
                    rsItem.Close
                    Set rsItem = Nothing
                    Exit Sub
                End If
            End If
        End If

        '更改模板,风格
        Conn.Execute ("update PE_Skin set ProjectName='" & TemplateProjectName & "' where ProjectName='" & rsItem("TemplateProjectName") & "'")
        Conn.Execute ("update PE_Template set ProjectName='" & TemplateProjectName & "' where ProjectName='" & rsItem("TemplateProjectName") & "'")

        rsItem("TemplateProjectName") = TemplateProjectName
        rsItem("Intro") = Intro
        rsItem.Update
        rsItem.Close
        Set rsItem = Nothing
    Else
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Call WriteSuccessMsg("<Li>" & SaveName & "方案成功！", "Admin_TemplateProject.asp?Action=Main")
    Call CloseConn

End Sub

'=================================================
'过程名：Import
'作  用：导入项目第一步
'=================================================
Sub Import()

    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' action='Admin_TemplateProject.asp' method='post' >"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>网站方案导入（第一步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的方案数据库的文件名："
    Response.Write "        <input name='ItemMdb' type='text' id='ItemMdb' value='../temp/PE_TemplateProject.mdb' size='50' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'> </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：Import2
'作  用：导入模板方案第二步
'=================================================
Sub Import2()
    On Error Resume Next
    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    mdbname = Replace(Trim(Request.Form("ItemMdb")), "'", "")

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入数据库名"
    End If

    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='post' action='Admin_TemplateProject.asp?action=DoImport'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>网站方案导入（第二步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>将被导入的方案项目</strong><br>"
    Response.Write "<select name='TemplateProjectID' size='2' multiple style='height:300px;width:250px;'>"
    sql = "select * from PE_TemplateProject"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何方案项目</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectID") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</select></td>"
    Response.Write "            <td width='80'><input type='submit' name='Submit' value='导入&gt;&gt;' "

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "></td>"
    Response.Write "            <td><strong>系统中已经存在的方案项目</strong><br>"
    Response.Write "             <select name='tItemID' size='2' multiple style='height:300px;width:250px;' disabled>"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何方案项目</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectID") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "              </select></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "            <br><b>提示：按住“Ctrl”或“Shift”键可以多选</b><br>"
    Response.Write "        <input name='mdbname' type='hidden' id='mdbname' value='" & mdbname & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "        <br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：DoImport
'作  用：导入模板方案项目处理
'=================================================
Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, rs, trs, mrs
    Dim rsTemplate, trsTemplate, rsSkin, trsSkin, rsLabel, trsLabel
    Dim TemplateProjectID
    
    TemplateProjectID = Trim(Request("TemplateProjectID"))
    If IsValidID(TemplateProjectID) = False Then
        TemplateProjectID = ""
    End If

    '获得导入模版数据库路径
    mdbname = Replace(Trim(Request.Form("mdbname")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("mdbname")), "'", "")
    End If

    mdbname = Replace(mdbname, "＄", "/") '防止外部链接安全问题

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入模版数据库名"
        Exit Sub
    End If

    If TemplateProjectID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导出的网站方案ID!</li>"
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

    '方案导入
    Set rs = tconn.Execute("select * from PE_TemplateProject where TemplateProjectID in (" & TemplateProjectID & ")  order by TemplateProjectID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_TemplateProject", Conn, 1, 3

    Do While Not rs.EOF

        If PE_CLng(Conn.Execute("select count(*) from PE_TemplateProject where TemplateProjectName='" & rs("TemplateProjectName") & "'")(0)) > 0 Then
            ErrMsg = ErrMsg & "<li><font color=red >" & rs("TemplateProjectName") & "</font>系统中已经有相同的方案没有导入!</li>"
        Else
            Set mrs = Conn.Execute("select max(TemplateProjectID) from PE_TemplateProject")

            If IsNull(mrs(0)) Then
                TemplateProjectID = 1
            Else
                TemplateProjectID = mrs(0) + 1
            End If

            Set mrs = Nothing

            trs.addnew
            trs("TemplateProjectID") = TemplateProjectID
            trs("TemplateProjectName") = rs("TemplateProjectName")
            trs("Intro") = rs("Intro")
            trs("IsDefault") = False
            '模板隶属方案导入
            Set rsTemplate = tconn.Execute("select * from PE_Template where ProjectName='" & rs("TemplateProjectName") & "' order by TemplateID")
            Set trsTemplate = Server.CreateObject("adodb.recordset")
            trsTemplate.Open "select * from PE_Template", Conn, 1, 3

            If rsTemplate.BOF Or rsTemplate.EOF Then
            Else

                Do While Not rsTemplate.EOF
                    trsTemplate.addnew
                    trsTemplate("ChannelID") = rsTemplate("ChannelID")
                    trsTemplate("TemplateName") = rsTemplate("TemplateName")
                    trsTemplate("TemplateType") = rsTemplate("TemplateType")
                    trsTemplate("TemplateContent") = rsTemplate("TemplateContent")
                    trsTemplate("IsDefault") = False
                    trsTemplate("ProjectName") = rsTemplate("ProjectName")
                    trsTemplate("IsDefaultInProject") = rsTemplate("IsDefaultInProject")
                    trsTemplate("Deleted") = rsTemplate("Deleted")
                    trsTemplate.Update
                    rsTemplate.MoveNext
                Loop

            End If

            trsTemplate.Close
            Set trsTemplate = Nothing
            rsTemplate.Close
            Set rsTemplate = Nothing
            '风格隶属方案导入
            Set rsSkin = tconn.Execute("select * from PE_Skin where ProjectName='" & rs("TemplateProjectName") & "' order by SkinID")
            Set trsSkin = Server.CreateObject("adodb.recordset")
            trsSkin.Open "select * from PE_Skin", Conn, 1, 3

            If rsSkin.BOF Or rsSkin.EOF Then
            Else

                Do While Not rsSkin.EOF
                    trsSkin.addnew
                    trsSkin("SkinName") = rsSkin("SkinName")
                    trsSkin("IsDefault") = False
                    trsSkin("Skin_CSS") = rsSkin("Skin_CSS")
                    trsSkin("IsDefaultInProject") = rsSkin("IsDefaultInProject")
                    trsSkin("ProjectName") = rsSkin("ProjectName")
                    trsSkin.Update
                    rsSkin.MoveNext
                Loop

            End If

            trsSkin.Close
            Set trsSkin = Nothing
            rsSkin.Close
            Set rsSkin = Nothing
            ErrMsg = ErrMsg & "<li><font color=blue >" & rs("TemplateProjectName") & "</font>方案导入成功!</li>"
            trs.Update
        End If

        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    '自定义标签导入
    Set trsLabel = tconn.Execute("select * from PE_Label")
    Set rsLabel = Server.CreateObject("adodb.recordset")
    rsLabel.Open "select * from PE_Label", Conn, 1, 3

    If Not trsLabel.EOF Then

        Do While Not trsLabel.EOF

            If PE_CLng(Conn.Execute("select count(*) from PE_Label where LabelName='" & trsLabel("LabelName") & "'")(0)) > 0 Then
            Else
                rsLabel.addnew
                rsLabel("LabelName") = trsLabel("LabelName")
                rsLabel("LabelClass") = trsLabel("LabelClass")
                rsLabel("LabelType") = trsLabel("LabelType")
                rsLabel("PageNum") = trsLabel("PageNum")
                rsLabel("reFlashTime") = trsLabel("reFlashTime")
                rsLabel("fieldlist") = trsLabel("fieldlist")
                rsLabel("LabelIntro") = trsLabel("LabelIntro")
                rsLabel("Priority") = trsLabel("Priority")
                rsLabel("LabelContent") = trsLabel("LabelContent")
                rsLabel("AreaCollectionID") = trsLabel("AreaCollectionID")
                rsLabel.Update
            End If

            trsLabel.MoveNext
        Loop

    End If

    Set trsLabel = Nothing
    rsLabel.Close
    Set rsLabel = Nothing
   
    tconn.Close
    Set tconn = Nothing
    Response.Write "<br>"
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>方案导入提示信息</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>" & ErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td><a href='" & ComeUrl & "'>&lt;&lt; 返回上一页</a></td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Call CreatSkinFile
End Sub

'=================================================
'过程名：Export
'作  用：导出模板方案项目
'=================================================
Sub Export()
    Dim rs, sql, iCount
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)
    Response.Write "<br>" & vbCrLf
    Response.Write "<FORM name=myform action='Admin_TemplateProject.asp' method=post>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>网站方案导出</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='TemplateProjectID' size='2' multiple style='height:300px;width:450px;'>"

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>还没有方案项目！</option>"
        '关闭提交按钮
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("TemplateProjectID") & "'>" & rs("TemplateProjectName") & "</option>"
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
    Response.Write "        <td colspan='2'>目标数据库：<input name='Itemmdb' type='text' id='ItemMdb' value='../Temp/PE_TemplateProject.mdb' size='30' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> 先清空目标数据库</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='执行导出操作'>"
    Response.Write "          <input name='Action' type='hidden' id='Action' value='DoExport'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateProjectID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateProjectID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateProjectID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateProjectID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'过程名：DoExport
'作  用：导出模板方案项目
'=================================================
Sub DoExport()
    On Error Resume Next
    
    Dim rs, trs, sql, rsLabel, trsLabel, rsTemplate, trsTemplate, rsSkin, trsSkin
    Dim mdbname, tconn
    Dim TemplateProjectID, TemplateProjectName, FormatConn

    FormatConn = Request.Form("FormatConn")
    TemplateProjectID = Trim(Request("TemplateProjectID"))
    mdbname = Replace(Trim(Request.Form("Itemmdb")), "'", "")
    If IsValidID(TemplateProjectID) = False Then
        TemplateProjectID = ""
    End If
    
    If TemplateProjectID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导出的网站方案ID!</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导出数据库名"
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
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    tconn.Execute ("select TemplateProjectID from PE_TemplateProject")

    If Err Then
        Set trs = Nothing
        ErrMsg = ErrMsg & "<li>您要导出的数据库,不是系统方案数据库,请使用系统方案数据库。"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If FormatConn <> "" Then '要删除的数据
        tconn.Execute ("delete from PE_Label")
        tconn.Execute ("delete from PE_Skin")
        tconn.Execute ("delete from PE_Template")
        tconn.Execute ("delete from PE_TemplateProject")
    End If

    '方案导出
    Set rs = Conn.Execute("select * from PE_TemplateProject where TemplateProjectID in (" & TemplateProjectID & ")  order by TemplateProjectID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_TemplateProject", tconn, 1, 3

    Do While Not rs.EOF
        trs.addnew
        trs("TemplateProjectID") = rs("TemplateProjectID")
        trs("TemplateProjectName") = rs("TemplateProjectName")
        trs("Intro") = rs("Intro")
        trs("IsDefault") = rs("IsDefault")
        '模板隶属方案导出
        Set rsTemplate = Conn.Execute("select * from PE_Template where ProjectName='" & rs("TemplateProjectName") & "' order by TemplateID")
        Set trsTemplate = Server.CreateObject("adodb.recordset")
        trsTemplate.Open "select * from PE_Template", tconn, 1, 3

        If rsTemplate.BOF Or rsTemplate.EOF Then
        Else

            Do While Not rsTemplate.EOF
                trsTemplate.addnew
                trsTemplate("TemplateID") = rsTemplate("TemplateID")
                trsTemplate("ChannelID") = rsTemplate("ChannelID")
                trsTemplate("TemplateName") = rsTemplate("TemplateName")
                trsTemplate("TemplateType") = rsTemplate("TemplateType")
                trsTemplate("TemplateContent") = rsTemplate("TemplateContent")
                trsTemplate("IsDefault") = rsTemplate("IsDefault")
                trsTemplate("ProjectName") = rsTemplate("ProjectName")
                trsTemplate("IsDefaultInProject") = rsTemplate("IsDefaultInProject")
                trsTemplate("Deleted") = rsTemplate("Deleted")
                trsTemplate.Update
                rsTemplate.MoveNext
            Loop

        End If

        trsTemplate.Close
        Set trsTemplate = Nothing
        rsTemplate.Close
        Set rsTemplate = Nothing
        '风格隶属方案导出
        Set rsSkin = Conn.Execute("select * from PE_Skin where ProjectName='" & rs("TemplateProjectName") & "' order by SkinID")
        Set trsSkin = Server.CreateObject("adodb.recordset")
        trsSkin.Open "select * from PE_Skin", tconn, 1, 3

        If rsSkin.BOF Or rsSkin.EOF Then
        Else

            Do While Not rsSkin.EOF
                trsSkin.addnew
                trsSkin("SkinID") = rsSkin("SkinID")
                trsSkin("SkinName") = rsSkin("SkinName")
                trsSkin("IsDefault") = rsSkin("IsDefault")
                trsSkin("Skin_CSS") = rsSkin("Skin_CSS")
                trsSkin("IsDefaultInProject") = rsSkin("IsDefaultInProject")
                trsSkin("ProjectName") = rsSkin("ProjectName")
                trsSkin.Update
                rsSkin.MoveNext
            Loop

        End If

        trsSkin.Close
        Set trsSkin = Nothing
        rsSkin.Close
        Set rsSkin = Nothing

        trs.Update
        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing

    '自定义标签导出
    Set trsLabel = Conn.Execute("select * from PE_Label")
    Set rsLabel = Server.CreateObject("adodb.recordset")
    rsLabel.Open "select * from PE_Label", tconn, 1, 3

    If Not trsLabel.EOF Then

        Do While Not trsLabel.EOF
            rsLabel.addnew
            rsLabel("LabelName") = trsLabel("LabelName")
            rsLabel("LabelClass") = trsLabel("LabelClass")
            rsLabel("LabelType") = trsLabel("LabelType")
            rsLabel("PageNum") = trsLabel("PageNum")
            rsLabel("reFlashTime") = trsLabel("reFlashTime")
            rsLabel("fieldlist") = trsLabel("fieldlist")
            rsLabel("LabelIntro") = trsLabel("LabelIntro")
            rsLabel("Priority") = trsLabel("Priority")
            rsLabel("LabelContent") = trsLabel("LabelContent")
            rsLabel("AreaCollectionID") = trsLabel("AreaCollectionID")
            rsLabel.Update
            trsLabel.MoveNext
        Loop

    End If

    Set trsLabel = Nothing
    rsLabel.Close
    Set rsLabel = Nothing
   
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("已经成功将所选中的方案导出到指定的数据库中！", ComeUrl)
End Sub

'*************************  类模块主区域结束  *******************************
'*************************  类模块扩展域开始  *******************************
'=================================================
'过程名：SetDefault
'作  用：设置方案默认
'=================================================
Sub SetDefault()
    Dim TemplateProjectID, ProjectName
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If TemplateProjectID = 0 Then
        FoundErr = True
        ErrMsg = "<li>方案ID 不能为空!</li>"
        Exit Sub
    End If

    '定义风格系统默认
    Conn.Execute ("update PE_Skin set IsDefault=" & PE_False & " where IsDefault=" & PE_True & "")
    Conn.Execute ("update PE_Skin set IsDefault=" & PE_True & " where IsDefaultInProject=" & PE_True & " and ProjectName='" & ProjectName & "'")
    '定义模板系统默认
    Conn.Execute ("update PE_Template set IsDefault=" & PE_False & " where IsDefault=" & PE_True & "")
    Conn.Execute ("update PE_Template set IsDefault=" & PE_True & " where IsDefaultInProject=" & PE_True & " and ProjectName='" & ProjectName & "'")
    '定义方案系统默认
    Conn.Execute ("update PE_TemplateProject set IsDefault=" & PE_False & " where IsDefault=" & PE_True & "")
    Conn.Execute ("update PE_TemplateProject set IsDefault=" & PE_True & " where TemplateProjectName='" & ProjectName & "'")

    Call WriteSuccessMsg("<li>成功将选定的方案设置为方案默认方案</li><li>成功将选定的风格设置为方案默认风格</li><li>成功将选定的模板设置为方案默认模板</li>", ComeUrl)
    Call CreatSkinFile
    Call ClearSiteCache(0)
End Sub

'=================================================
'过程名：DelTemplateProject
'作  用：确认删除方案
'=================================================
Sub DelTemplateProject()
    Dim TemplateProjectID, ProjectName, strTemp
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
    Response.Write "        <br>" & vbCrLf
    Response.Write "        <table border='0' align='center' cellpadding='0' cellspacing='1' width='350' height='150' class='border'>" & vbCrLf
    Response.Write "          <tr class='title' height='22'>" & vbCrLf
    Response.Write "           <td align='center' ><strong>您确认删除方案么</strong></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td  align='center' class='tdbg' valign='top'>"
    Response.Write "           <br><br>&nbsp;&nbsp;确定要<FONT color='red'>删除此方案吗？</font>删除此方案后方案隶属的<FONT color='blue'>模板,风格</font> 都将会被删除,请绝对注意!<br><br><br>"
    Response.Write "                <FONT color='red'> <a href='Admin_TemplateProject.asp?action=Del2&TemplateProjectID=" & TemplateProjectID & "&ProjectName=" & ProjectName & "'>确认删除</a></FONT>&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "                <FONT color='blue'> <a href='Admin_TemplateProject.asp?action=main'> 返 回 </a></FONT> " & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
        
End Sub

'=================================================
'过程名：DelTemplateProject2
'作  用：澈底删除方案
'=================================================
Sub DelTemplateProject2()
    Dim rs, sql
    Dim TemplateProjectID, ProjectName, strTemp
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If TemplateProjectID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定TemplateProjectID</li>"
        Exit Sub
    End If

    sql = "select * from PE_TemplateProject where TemplateProjectID=" & TemplateProjectID
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的方案！</li>"
    Else

        If rs("IsDefault") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>当前方案为默认方案，不能删除。请先将默认改为其他方案后再来删除此方案。</li>"
        End If
    End If

    If FoundErr = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.Delete
    rs.Update
    rs.Close
    Set rs = Nothing

    Conn.Execute ("delete from PE_Skin where ProjectName='" & ProjectName & "'")
    Conn.Execute ("delete from PE_Template where ProjectName='" & ProjectName & "'")

    strTemp = strTemp & "<li>成功删除选定的方案。</li>"
    strTemp = strTemp & "<li>成功删除选定的方案中的所有模板。</li>"
    strTemp = strTemp & "<li>成功删除选定的方案中的所有风格。</li>"

    Call WriteSuccessMsg(strTemp, "Admin_TemplateProject.asp?Action=main")
End Sub

'=================================================
'过程名：TemplateProject
'作  用：模板方案频道选项
'=================================================
Sub TemplateProject()

    Dim sql, rs
    Dim iTemplateType, iChannelID, i, Num
    iChannelID = 0
    iTemplateType = 0
    i = 0
    Num = 1
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
        
    sql = "select * from PE_Template where Deleted=" & PE_False & " and ChannelID=" & ChannelID & " order by TemplateType,ChannelID"
        
    Set rs = Conn.Execute(sql)
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "     <tr class='title' height='22'>"
    Response.Write "      <td width='30' align='center'><strong>选择</strong></td>"
    Response.Write "      <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "      <td width='150' align='center'><b>模板类型</b></td>"
    Response.Write "      <td height='22' align='center'><strong>模板名称</strong></td>"
    Response.Write "      <td width='80' align='center'><strong>是否默认</strong></td>"
    Response.Write "     </tr>"
    i = 0

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td width='100%' colspan='6' align='center'> 没 有 任 何 模 板</td></tr>"
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

            If rs("IsDefault") = True Then Response.Write "checked"
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "      <td width='30' align='center'>" & rs("TemplateID") & "</td>"
            Response.Write "      <td width='150' align='center'>" & GetTemplateTypeName(rs("TemplateType"), rs("ChannelID")) & "</td>"
            Response.Write "      <td align='center'><a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&TemplateID=" & rs("TemplateID") & "'>" & rs("TemplateName") & "</a></td>"
            Response.Write "      <td width='80' align='center'><b>"

            If rs("IsDefault") = True Then
                Response.Write "√"
            Else
                Response.Write "×"
            End If

            Response.Write "</td>"
            Response.Write "</tr>"

            rs.MoveNext
        Loop

        Response.Write "<Input TYPE='hidden' Name='Num' value='" & Num & "'>"

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</table>  "
    Response.Write "</form>"
End Sub

'=================================================
'过程名：CreatSkinFile
'作  用：显示处理结果生成css文件
'=================================================
Sub CreatSkinFile()

    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    If Not fso.FolderExists(Server.MapPath(InstallDir)) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先进行网站配置后再进行此项操作。</li>"
        Exit Sub
    End If

    If Not fso.FolderExists(Server.MapPath(InstallDir & "Skin")) Then
        fso.CreateFolder (Server.MapPath(InstallDir & "Skin"))
    End If

    Dim rsSkin, sqlSkin, hf, strSkin
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Conn.Execute(sqlSkin)

    Do While Not rsSkin.EOF
        strSkin = Replace_CaseInsensitive(rsSkin("Skin_CSS"), "Skin/", InstallDir & "Skin/")
        Call WriteToFile(InstallDir & "Skin/Skin" & rsSkin("SkinID") & ".css", strSkin)
        rsSkin.MoveNext
    Loop

    rsSkin.Close
    sqlSkin = "select * from PE_Skin where IsDefault=" & PE_True & ""
    Set rsSkin = Conn.Execute(sqlSkin)

    If rsSkin.BOF And rsSkin.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你还没有将其中一个风格设为默认风格哦。请记得一定要做这一步呀。</li>"
    Else
        strSkin = Replace_CaseInsensitive(rsSkin("Skin_CSS"), "Skin/", InstallDir & "Skin/")
        Call WriteToFile(InstallDir & "Skin/DefaultSkin.css", strSkin)
    End If

    rsSkin.Close
    Set rsSkin = Nothing
End Sub

'=================================================
'过程名：TemplateBatchMove
'作  用：批量迁移模板
'=================================================
Sub TemplateBatchMove()
    Dim rs, sql
    Dim TemplateID, TemplateProjectID, ProjectName, TemplateChannelID

    TemplateID = ReplaceBadChar(Trim(Request("TemplateID")))
    TemplateChannelID = PE_CLng(Trim(Request("TemplateChannelID")))
    TemplateProjectID = ReplaceBadChar(Trim(Request("TemplateProjectID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If ProjectName = "" Then
        Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")

        If rs.BOF And rs.EOF Then
            Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
            Exit Sub
        Else
            ProjectName = rs("TemplateProjectName")
        End If

        Set rs = Nothing
    End If
    
    Response.Write "<form method=""post"" action=""Admin_TemplateProject.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td  align='center'><b>方案间模板迁移 </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "          <td class='tdbg5' valign='top' width='50%'>" & vbCrLf
    Response.Write "            <table width='100%' border='0' cellpadding='2' cellspacing='1'>" & vbCrLf
    Response.Write "              <tr>" & vbCrLf
    Response.Write "                <td width='80'></td>" & vbCrLf
    Response.Write "                <td>" & vbCrLf
    Response.Write "                                &nbsp;&nbsp;&nbsp;&nbsp;<b>选择方案中要迁移的的模板</b><br>" & vbCrLf
    Response.Write "            <select name='ProjectName' style='width:150px;'  onChange='document.form1.submit();'>"
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
    sql = "SELECT DISTINCT t.ChannelID, c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID"
    Set rs = Conn.Execute(sql)
    Response.Write "<select name='TemplateChannelID' id='TemplateChannelID' onChange='document.form1.submit();'>"

    If rs.BOF And rs.EOF Then
        Response.Write "<option value="" selected>还没有添加频道！</option> "
    Else

        Do While Not rs.EOF
            Response.Write "<option value=" & rs("ChannelID") & " " & OptionValue(rs("ChannelID"), TemplateChannelID) & ">" & rs("ChannelName") & "</option>"
            rs.MoveNext
        Loop

        Response.Write "<option value='0' " & OptionValue(0, TemplateChannelID) & ">系统通用模板</option> "
        Response.Write "<option value='999999' " & OptionValue(999999, TemplateChannelID) & ">方案所有模板</option> "
    End If

    Response.Write "</select>"
    rs.Close
    Set rs = Nothing
    Response.Write "              <br>"
    sql = "select ChannelID,TemplateID,TemplateName from PE_Template where "

    If TemplateChannelID <> 999999 Then
        If TemplateChannelID > 0 Then
            sql = sql & " ChannelID=" & TemplateChannelID & " and "
        ElseIf TemplateChannelID = 0 Then
            sql = sql & " ChannelID=0 and "
        End If
    End If

    sql = sql & " ProjectName='" & ProjectName & "' and Deleted=" & PE_False
    '显示模版
    Response.Write "              <select name='BatchTemplateID' id='BatchTemplateID' size='2' multiple style='height:250px;width:250px;' >"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        '没有模版时指定关闭提交按钮
        Response.Write "                <option value='0'>该方案还有没有任何模板</option>"
    Else

        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"

    Response.Write "  <br>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' 选定所有 ' onclick='SelectAll()'>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'><br>" & vbCrLf
    Response.Write "  <FONT style='font-size:12px' color=''><b>按住“Ctrl”或“Shift”键可以多选</b></FONT>" & vbCrLf
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "    function SelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.BatchTemplateID.length;i++){" & vbCrLf
    Response.Write "        document.form1.BatchTemplateID.options[i].selected=true;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function UnSelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.BatchTemplateID.length;i++){" & vbCrLf
    Response.Write "        document.form1.BatchTemplateID.options[i].selected=false;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function CheckForm(){" & vbCrLf
    Response.Write "        if (document.form1.BatchTemplateID.value==""""){" & vbCrLf
    Response.Write "            alert(""迁移模板不能为空！"");" & vbCrLf
    Response.Write "            document.form1.BatchTemplateID.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (document.form1.MoveTemplateProjectName.value==""""){" & vbCrLf
    Response.Write "            alert(""迁移的方案不能为空！"");" & vbCrLf
    Response.Write "            document.form1.MoveTemplateProjectName.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (document.form1.ProjectName.value==document.form1.MoveTemplateProjectName.value){" & vbCrLf
    Response.Write "            alert(""方案迁移不能自己给自己移动复制！"");" & vbCrLf
    Response.Write "            document.form1.ProjectName.focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        document.form1.Action.value='DoTemplateBatchMove';" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "                </td>" & vbCrLf
    Response.Write "              </tr>" & vbCrLf
    Response.Write "            </table>" & vbCrLf
    Response.Write "          </td>" & vbCrLf

    Response.Write "          <td width='80' class='tdbg' align='center'>" & vbCrLf
        
    Response.Write "<Input TYPE='radio' Name='BatchTypeName' value='移动' > 移动到 &gt;&gt;" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<Input TYPE='radio' Name='BatchTypeName' value='复制' > 复制到 &gt;&gt;" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "           <input type='submit' name='Submit' value=' 确 定 ' onClick=""javascript:return CheckForm()"" >" & vbCrLf
    Response.Write "          </td>"
    Response.Write "          <td class='tdbg' align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>系统中已经存在的方案项目</strong><br>"
    Response.Write "             &nbsp;&nbsp;<select name='MoveTemplateProjectName' size='2'  style='height:300px;width:200px;' >"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何方案项目</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "           </select></td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "     </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<center><FONT color='red'> 注:</FONT>移动的时候,<FONT color='#3366FF'>系统默认,方案默认</FONT>是不会移动的。</center> " & vbCrLf
    Response.Write "<input name=""Action"" type=""hidden"" id=""Action"" value=""TemplateBatchMove"">" & vbCrLf
    Response.Write "</form>" & vbCrLf

End Sub

'=================================================
'过程名：DoTemplateBatchMove
'作  用：批量迁移模板处理
'=================================================
Sub DoTemplateBatchMove()

    Dim rs, trs, jrs, sql
    Dim TemplateType, TemplateID, TemplateProjectName, TemplateChannelID, BatchTemplateID
    Dim ProjectName, MoveTemplateProjectName, BatchTypeName
    Dim tempIsDefault, tempIsDefaultInProject, SysDefault '临时数据
        
    FoundErr = False
    tempIsDefault = False
    tempIsDefaultInProject = False

    BatchTypeName = Trim(Request.Form("BatchTypeName"))
    TemplateProjectName = ReplaceBadChar(Trim(Request.Form("TemplateProjectName")))
    TemplateChannelID = PE_CLng(Trim(Request.Form("TemplateChannelID")))
    BatchTemplateID = Trim(Request.Form("BatchTemplateID"))
    ProjectName = Trim(Request.Form("ProjectName"))
    MoveTemplateProjectName = ReplaceBadChar(Trim(Request.Form("MoveTemplateProjectName")))
    If IsValidID(BatchTemplateID) = False Then
        BatchTemplateID = ""
    End If
    
    If BatchTypeName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请选择要迁移的类型,是移动还是复制。</li>"
    End If

    If BatchTemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有模板ID号,请返回输入要" & BatchTypeName & "的模板ID</li>"
    End If

    If FoundInArr(MoveTemplateProjectName, ProjectName, ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>相同的方案不能" & BatchTypeName & ",请返回输入" & BatchTypeName & "不同的方案</li>"
    End If

    TemplateID = BatchTemplateID

    If MoveTemplateProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择要" & BatchTypeName & "的方案</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    '得到系统方案默认名称
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>选择的模板类型不对</li>"
    Else
        SysDefault = rs("TemplateProjectName")
    End If

    Set rs = Nothing

    If FoundErr = True Then
        Exit Sub
    End If

    sql = "select * from PE_Template where "

    If InStr(TemplateID, ",") > 0 Then
        sql = sql & " TemplateID in (" & TemplateID & ")"
    Else
        sql = sql & " TemplateID=" & TemplateID
    End If

    If BatchTypeName = "移动" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, Conn, 1, 3

        '这要加计算
        Do While Not rs.EOF
            If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
                ErrMsg = ErrMsg & "<li>是" & rs("ProjectName") & "方案的默认模板不能<FONT color='red'>移动</Font>!"
            Else
                rs("IsDefault") = False
                rs("IsDefaultInProject") = False
                rs("ProjectName") = MoveTemplateProjectName
                ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("TemplateName") & "</FONT>模板成功" & BatchTypeName & "到" & MoveTemplateProjectName & "方案!"
                rs.Update
            End If
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
    Else
        '这要加计算
        Set rs = Conn.Execute(sql)
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Template", Conn, 1, 3

        Do While Not rs.EOF
            trs.addnew
            trs("ChannelID") = rs("ChannelID")
            trs("TemplateName") = rs("TemplateName")
            trs("TemplateType") = rs("TemplateType")
            trs("TemplateContent") = rs("TemplateContent")
            
            '检测有无重复
            Set jrs = Conn.Execute("select * from PE_Template where ChannelID=" & trs("ChannelID") & " and ProjectName='" & MoveTemplateProjectName & "' and TemplateType=" & trs("TemplateType"))

            If jrs.BOF And jrs.EOF Then
                tempIsDefault = True
            Else

                Do While Not jrs.EOF

                    If tempIsDefault = False Then
                        If jrs("IsDefault") = True Or trs("ProjectName") = MoveTemplateProjectName Or jrs("ProjectName") <> SysDefault Then
                            tempIsDefault = True
                        End If
                    End If

                    If tempIsDefaultInProject = False Then
                        If jrs("IsDefaultInProject") = True Or trs("ProjectName") = MoveTemplateProjectName Then
                            tempIsDefaultInProject = True
                        End If
                    End If

                    If tempIsDefault = True And tempIsDefaultInProject = True Then
                        Exit Do
                    End If

                    jrs.MoveNext
                Loop

            End If

            Set jrs = Nothing

            If tempIsDefault = True Then
                trs("IsDefault") = False
            Else
                trs("IsDefault") = rs("IsDefault")
            End If

            If tempIsDefaultInProject = True Then
                trs("IsDefaultInProject") = False
            Else
                trs("IsDefaultInProject") = rs("IsDefaultInProject")
            End If

            trs("ProjectName") = MoveTemplateProjectName
            trs("Deleted") = rs("Deleted")
            ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("TemplateName") & "</FONT>模板成功" & BatchTypeName & "到" & MoveTemplateProjectName & "方案!"
            tempIsDefaultInProject = False
            trs.Update
            rs.MoveNext
        Loop

        trs.Close
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If

    Call WriteSuccessMsg(ErrMsg, "Admin_TemplateProject.asp?action=Main&ProjectName=" & SysDefault)
End Sub

'=================================================
'过程名：SkinBatchMove
'作  用：批量迁移风格
'=================================================
Sub SkinBatchMove()

    Dim rs, sql
    Dim SkinID, ProjectName, BatchTypeName

    SkinID = ReplaceBadChar(Trim(Request("SkinID")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
    BatchTypeName = Trim(Request("BatchTypeName"))

    If ProjectName = "" Then
        ProjectName = "所有方案"
    End If

    Response.Write "<form method=""post"" action=""Admin_TemplateProject.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td  align='center'><b>方案间风格迁移 </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "          <td class='tdbg5' valign='top' width='50%' >" & vbCrLf
    Response.Write "                <table border='0' cellpadding='0' cellspacing='1' width='100%' height='100%'>" & vbCrLf
    Response.Write "                  <tr class='tdbg'>" & vbCrLf
    Response.Write "                    <td width='100' class='tdbg5'></td>" & vbCrLf
    Response.Write "                    <td align='left' class='tdbg5'>" & vbCrLf
    Response.Write "                     <select name='ProjectName' style='width:150px;' onChange='document.form1.submit();'>"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "                <option value='0'>没有任何方案项目</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

        Response.Write "<option value='所有方案' " & OptionValue("所有方案", ProjectName) & ">所有方案</option>"
    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                       </select>"
    Response.Write "                </td>" & vbCrLf
    Response.Write "                   </tr>" & vbCrLf
    Response.Write "                   <tr class='tdbg'>" & vbCrLf
    Response.Write "                     <td width='100' class='tdbg5'></td>" & vbCrLf
    Response.Write "                     <td  align='left' class='tdbg5'>"
    Response.Write "                    <select name='SkinID'  size='2' multiple style='height:250px;width:250px;'>"
    sql = "select * from PE_Skin"

    If ProjectName <> "所有方案" And ProjectName <> "" Then
        sql = sql & " where ProjectName='" & ProjectName & "'"
    End If

    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何风格</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("SkinID") & "' "

            If SkinID <> "" Then
                If InStr(SkinID, ",") > 0 Then
                    If FoundInArr(SkinID, rs("SkinID"), ",") = True Then Response.Write "selected"
                Else
                    SkinID = PE_CLng(Trim(SkinID))

                    If rs("SkinID") = SkinID Then Response.Write "selected"
                End If
            End If

            Response.Write ">" & rs("SkinName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"
    Response.Write "  <br>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' 选定所有 ' onclick='SelectAll()'>" & vbCrLf
    Response.Write "  <Input type='button' name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'><br>" & vbCrLf
    Response.Write "  <FONT style='font-size:12px' color=''><b>按住“Ctrl”或“Shift”键可以多选</b></FONT>" & vbCrLf
    Response.Write "  <script language='javascript'>" & vbCrLf
    Response.Write "    function SelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.SkinID.length;i++){" & vbCrLf
    Response.Write "        document.form1.SkinID.options[i].selected=true;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function UnSelectAll(){" & vbCrLf
    Response.Write "        for(var i=0;i<document.form1.SkinID.length;i++){" & vbCrLf
    Response.Write "        document.form1.SkinID.options[i].selected=false;}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  </script>" & vbCrLf
    Response.Write "                  </td>" & vbCrLf
    Response.Write "                     </tr>" & vbCrLf
    Response.Write "                    </table>" & vbCrLf
    Response.Write "               </td>" & vbCrLf
    Response.Write "               <td width='80' class='tdbg' align='center'>" & vbCrLf
    Response.Write "                 <Input TYPE='radio' Name='BatchTypeName' value='移动' " & IsRadioChecked(BatchTypeName, "移动") & "   > 移动到 &gt;&gt;<br>" & vbCrLf
    Response.Write "                 <Input TYPE='radio' Name='BatchTypeName' value='复制' " & IsRadioChecked(BatchTypeName, "复制") & " > 复制到 &gt;&gt;<br>" & vbCrLf
    Response.Write "                 <Input type='submit' name='Submit' value=' 确 定 ' onClick=""document.form1.Action.value='DoSkinBatchMove';"" >" & vbCrLf
    Response.Write "               </td>"
    Response.Write "               <td class='tdbg' align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>系统中已经存在的方案项目</strong><br>"
    Response.Write "             &nbsp;&nbsp;<select name='MoveTemplateProjectName' size='2'  style='height:300px;width:200px;' >"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何方案项目</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "'>" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "           </select></td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "     </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<center><FONT color='red'> 注:</FONT>移动的时候,<FONT color='#3366FF'>系统默认,方案默认</FONT>是不会移动的。</center> " & vbCrLf
    Response.Write "<input name=""Action"" type=""hidden"" id=""Action"" value=""SkinBatchMove"">" & vbCrLf
    Response.Write "</form>" & vbCrLf

End Sub

'=================================================
'过程名：DoSkinBatchMove
'作  用：处理批量迁移风格
'=================================================
Sub DoSkinBatchMove()

    Dim rs, trs, jrs, sql
    Dim SkinID
    Dim MoveTemplateProjectName, BatchTypeName, SysDefault

    Dim tempIsDefault, tempIsDefaultInProject '临时数据
        
    FoundErr = False
    tempIsDefault = False
    tempIsDefaultInProject = False

    BatchTypeName = Trim(Request.Form("BatchTypeName"))
    SkinID = Trim(Request.Form("SkinID"))
    MoveTemplateProjectName = ReplaceBadChar(Trim(Request.Form("MoveTemplateProjectName")))
    If IsValidID(SkinID) = False Then
        SkinID = ""
    End If

    If BatchTypeName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择移动或复制类型</li>"
    End If

    If SkinID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择要" & BatchTypeName & "的风格</li>"
    End If

    If MoveTemplateProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择要" & BatchTypeName & "的方案</li>"
    End If

    '得到系统方案默认名称
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>选择的模板类型不对</li>"
    Else
        SysDefault = rs("TemplateProjectName")
    End If

    Set rs = Nothing

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    sql = "select * from PE_Skin where "

    If InStr(SkinID, ",") > 0 Then
        sql = sql & " SkinID in (" & SkinID & ")"
    Else
        sql = sql & " SkinID=" & SkinID
    End If

    If BatchTypeName = "移动" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, Conn, 1, 3

        Do While Not rs.EOF
            If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
                ErrMsg = ErrMsg & "<li><FONT color='red'> " & rs("SkinName") & "</FONT>是" & rs("ProjectName") & "方案的默认风格不能移动!"
            Else
                rs("IsDefault") = False
                rs("IsDefaultInProject") = False
                rs("ProjectName") = MoveTemplateProjectName
                ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("SkinName") & "</FONT>风格成功" & BatchTypeName & "到" & MoveTemplateProjectName & "方案!"
                rs.Update
            End If
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
    Else
        Set rs = Conn.Execute(sql)
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Skin", Conn, 1, 3

        Do While Not rs.EOF
            trs.addnew
            trs("SkinName") = rs("SkinName")
            trs("Skin_CSS") = rs("Skin_CSS")
            '检测有无重复
            Set jrs = Conn.Execute("select * from PE_Skin where ProjectName='" & MoveTemplateProjectName & "'")

            If jrs.BOF And jrs.EOF Then
                tempIsDefault = True
            Else

                Do While Not rs.EOF

                    If tempIsDefault = False Then
                        If jrs("IsDefault") = True Or trs("ProjectName") = MoveTemplateProjectName Or jrs("ProjectName") <> SysDefault Then
                            tempIsDefault = True
                        End If
                    End If

                    If tempIsDefaultInProject = False Then
                        If jrs("IsDefaultInProject") = True Or trs("ProjectName") = MoveTemplateProjectName Then
                            tempIsDefaultInProject = True
                        End If
                    End If

                    If tempIsDefault = True And tempIsDefaultInProject = True Then
                        Exit Do
                    End If

                    jrs.MoveNext
                Loop

            End If

            Set jrs = Nothing

            If tempIsDefault = True Then
                trs("IsDefault") = False
            Else
                trs("IsDefault") = rs("IsDefault")
            End If

            If tempIsDefaultInProject = True Then
                trs("IsDefaultInProject") = False
            Else
                trs("IsDefaultInProject") = rs("IsDefaultInProject")
            End If

            trs("ProjectName") = MoveTemplateProjectName
            ErrMsg = ErrMsg & "<li><FONT color='blue'> " & rs("SkinName") & "</FONT>风格成功" & BatchTypeName & "到" & MoveTemplateProjectName & "方案!"
            tempIsDefaultInProject = False
            trs.Update
            rs.MoveNext
        Loop

        trs.Close
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If

    Call WriteSuccessMsg(ErrMsg, "Admin_TemplateProject.asp?action=main")
End Sub

'*************************  类模块扩展域结束  *******************************
'*************************  类模块函数通用开始  *****************************
'=================================================
'函数名：GetTemplateTypeName
'作  用：显示当前频道的模版类型
'参  数：iTemplateType --- 代入的模版值
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

                Case 13
                    GetTemplateTypeName = "在线支付第一步模板"

                Case 14
                    GetTemplateTypeName = "在线支付第二步模板"

                Case 15
                    GetTemplateTypeName = "在线支付第三步模板"

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
                GetTemplateTypeName = "来源列表页模板"
				
            Case 103
                GetTemplateTypeName = "匿名投稿模板"				

            Case 14
                GetTemplateTypeName = "厂商显示页模板"

            Case 15
                GetTemplateTypeName = "厂商列表页模板"

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
                'Case 22
                '    GetTemplateTypeName = "更改密码页模板 (后台)"
                'Case 23
                '    GetTemplateTypeName = "更改资料页模板 (后台)"
                'Case 24
                '    GetTemplateTypeName = "查看资料页模板 (后台)"
                'Case 999
                '    GetTemplateTypeName = "通用显示页模板"
        End Select

    End If

    If iTemplateType = 0 Then
        GetTemplateTypeName = "当前类型所有模板"
    End If

End Function

'**************************************************
'函数名：ReplaceText
'作  用：过滤非法字符串
'参  数：iText-----输入字符串
'返回值：替换后字符串
'**************************************************
Function ReplaceText(iText, _
                             iType)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol

    If PE_Cache.GetValue("Site_ReplaceText") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText from PE_KeyLink where isUse=1 and LinkType=" & iType
        rsKey.Open sqlKey, Conn, 1, 1

        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_ReplaceText_" & iType, rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceText = iText
            Exit Function
        End If
    End If

    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_ReplaceText_" & iType), "@@@")

    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        rText = Replace(rText, Keycol(0), Keycol(1))
    Next

    ReplaceText = rText
End Function

'**************************************************
'函数名：IsOptionSelected
'作  用：下拉菜单默认比较
'参  数：Compare1-----比较值1
'参  数：Compare2-----比较值2
'返回值：替换后字符串
'**************************************************
Function IsOptionSelected(ByVal Compare1, _
                                  ByVal Compare2)

    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If

End Function

'**************************************************
'函数名：IsFontChecked
'作  用：单选,多选默认
'参  数：Compare1-----比较值1
'参  数：Compare2-----比较值2
'返回值：替换后字符串
'**************************************************
Function IsFontChecked(ByVal Compare1, _
                               ByVal Compare2)

    If Compare1 = Compare2 Then
        IsFontChecked = " color='red'"
    Else
        IsFontChecked = ""
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
%>
