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
Const PurviewLevel_Others = "Skin"   '其他权限

Dim ProjectName, rs, sql
ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
If ProjectName = "" Then
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rs.BOF And rs.EOF Then
        Call WriteErrMsg("<li>系统中还没有默认方案,请到方案管理指定默认方案！</li>", ComeUrl)
        Response.End
    Else
        ProjectName = rs("TemplateProjectName")
    End If
    Set rs = Nothing
End If

Response.Write "<html><head><title>" & ProjectName & "方案 ---- 风格管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle(ProjectName & "方案 ---- 风 格 管 理", 10007)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Skin.asp'>风格管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Skin.asp?Action=Add'>添加风格</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Skin.asp?Action=Export'>风格导出</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Skin.asp?Action=Import'>风格导入</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Action = Trim(Request("Action"))

Select Case Action

    Case "Add"
        Call Add

    Case "Modify"
        Call Modify

    Case "SaveAdd"
        Call SaveAdd

    Case "SaveModify"
        Call SaveModify

    Case "Set"
        Call SetDefault

    Case "Del"
        Call DelSkin

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

    Case "Refresh"
        Call CreatSkinFile

        If FoundErr = False Then
            Call WriteSuccessMsg("刷新CSS风格文件成功！", ComeUrl)
        End If

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
'作  用：调用风格首页
'=================================================
Sub main()

    Dim rs, sql
    Dim rsTemplateProject, rsProjectName, sqlTemplateProject, i, SysDefault

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
    Response.Write "</script>" & vbCrLf
        
    Response.Write "<form name='myform' method='post' action='Admin_Skin.asp'>"
    Response.Write "<IMG SRC='images/img_u.gif' height='12'>您现在的位置：网站风格管理&nbsp;&gt;&gt;&nbsp;" & ProjectName

    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22' align='center'>"
    Response.Write "      <td width='30' align='center'><strong>选择</strong></td>"
    Response.Write "      <td width='50'><strong>ID</strong></td>"
    Response.Write "      <td width='100'><strong>方案名称</strong></td>"
    Response.Write "      <td ><strong>风格名称</strong></td>"

    If SysDefault = ProjectName Then
        Response.Write "      <td width='60'><strong>系统默认</strong></td>"
    Else
        Response.Write "      <td width='60'><strong>方案默认</strong></td>"
    End If

    Response.Write "      <td width='300' height='22' align='center'><strong> 操作</strong></td>"
    Response.Write "    </tr>"

    If ProjectName = "" Then
        sql = "select * from PE_Skin where ProjectName='' or ProjectName is null"
    ElseIf ProjectName = "所有方案" Then
        sql = "select * from PE_Skin"
    Else
        sql = "select * from PE_Skin where ProjectName='" & ProjectName & "'"
    End If
    
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td width='100%' colspan='8' align='center'>"

        If ProjectName = "" Then
            Response.Write "没 有 任 何 风 格"
        Else
            Response.Write "在 该 方 案 还 没 有 定 义 风 格"
        End If

        Response.Write "</td></tr>"
    Else

        Do While Not rs.EOF
            Response.Write "<tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "  <td width=""30"" align=""center"" height=""30"">" & vbCrLf
            Response.Write "    <input type=""checkbox"" value=" & rs("SkinID") & " name=""SkinID"""

            If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then Response.Write "disabled"
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
                        
            Response.Write "      <td width='50' align='center'>" & rs("SkinID") & "</td>"
            Response.Write "      <td align='center' width='100'>" & rs("ProjectName") & "</td>"
            Response.Write "      <td align='center'>" & rs("SkinName") & "</td>"
           
            If SysDefault = ProjectName Then
                Response.Write "      <td width='60' align='center'>"

                If rs("IsDefault") = True Then
                    Response.Write "<FONT style='font-size:12px' color='#008000'><b>√</b></FONT>"
                End If

                Response.Write "</td>"
            Else
                Response.Write "      <td width='60' align='center'>"

                If rs("IsDefaultInProject") = True Then
                    Response.Write "<b>√</b>"
                Else
                End If

                Response.Write "</td>"
            End If

            Response.Write "      <td width='300' align='center'>"

            If SysDefault = ProjectName Then
                If rs("IsDefault") = False And ProjectName = SysDefault Then
                    Response.Write "&nbsp;<a href='Admin_Skin.asp?Action=Set&DefaultType=1&SkinID=" & rs("SkinID") & "&ProjectName=" & ProjectName & "'>设为系统默认</a>"
                Else
                    Response.Write "<font color='gray'>&nbsp;设为系统默认</font>"
                End If

            Else
                        
                If rs("IsDefaultInProject") = False Then
                    Response.Write "&nbsp;&nbsp;<a href='Admin_Skin.asp?Action=Set&DefaultType=2&SkinID=" & rs("SkinID") & "&ProjectName=" & ProjectName & "'>设为方案默认</a>"
                Else
                    Response.Write "<font color='gray'>&nbsp;&nbsp;设为方案默认</font>"
                End If
            End If

            Response.Write "&nbsp;&nbsp;<a href='Admin_Skin.asp?Action=Modify&ProjectName=" & ProjectName & "&SkinID=" & rs("SkinID") & "'>修改风格</a>&nbsp;&nbsp;"

            If rs("IsDefaultInProject") = False And rs("IsDefault") = False Then
                Response.Write "<a href='Admin_Skin.asp?Action=Del&SkinID=" & rs("SkinID") & "&ProjectName=" & ProjectName & "' onClick=""return confirm('确定要删除此风格吗？删除此风格后原使用此风格的文章将改为使用系统默认风格。');"">删除风格</a>"
            Else
                Response.Write "<font color='gray'>删除风格</font>"
            End If

            Response.Write "      </td>"
            Response.Write "    </tr>"
            rs.MoveNext
        Loop

        Response.Write "    <tr class=""tdbg""> " & vbCrLf
        Response.Write "      <td colspan=8 height=""30"">" & vbCrLf
        Response.Write "        <input name=""Action"" type=""hidden""  value=""Del"">   " & vbCrLf
        Response.Write "        <input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" >选中所有项目" & vbCrLf
        Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;将选定的项目： " & vbCrLf
        Response.Write "        <input type=""submit"" value=""批&nbsp;量&nbsp;删&nbsp;除 "" name=""Del"" onclick='return confirm(""确定要删除此风格吗？删除此风格后原使用此风格的文章将改为使用系统默认风格。"");' >&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <Input TYPE='hidden' Name='BatchTypeName' value='移动'>" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr> " & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='40' colspan='7' align='center'><input type='submit' name='Submit' value='刷新风格CSS文件' onclick=""document.myform.Action.value='Refresh'""></td>"
        Response.Write "    </tr>"
    End If

    Response.Write "  </table>"
    Response.Write "</form>"
    rs.Close
    Set rs = Nothing
End Sub

'=================================================
'过程名：Export
'作  用：导出风格
'=================================================
Sub Export()

    Dim rs, sql, iCount

    sql = "select * from PE_Skin"
    Set rs = Conn.Execute(sql)
 
    Response.Write "<form name='myform' method='post' action='Admin_Skin.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>风格导出</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='SkinID' size='2' multiple style='height:300px;width:450px;'>"
    
    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>还没有风格！</option>"
        '关闭提交按钮
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("SkinID") & "'>" & rs("SkinName") & "</option>"
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
    Response.Write "        <td colspan='2'>目标数据库：<input name='SkinMdb' type='text' id='SkinMdb' value='../Skin/Skin.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> 先清空目标数据库</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='执行导出操作' onClick=""document.myform.Action.value='DoExport';"">"
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
    Response.Write "  for(var i=0;i<document.myform.SkinID.length;i++){" & vbCrLf
    Response.Write "    document.myform.SkinID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.SkinID.length;i++){" & vbCrLf
    Response.Write "    document.myform.SkinID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'过程名：Import
'作  用：导入风格第一步
'=================================================
Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Skin.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>风格导入（第一步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的风格数据库的文件名："
    Response.Write "        <input name='SkinMdb' type='text' id='SkinMdb' value='../Skin/Skin.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'> </td>"
    Response.Write " </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：Import2
'作  用：导入风格第二步
'=================================================
Sub Import2()
    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    mdbname = Replace(Trim(Request.Form("skinmdb")), "'", "")

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入模板数据库名"
        Exit Sub
    End If
    
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If

    Response.Write "<form name='myform' method='post' action='Admin_Skin.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>风格导入（第二步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>将被导入的风格</strong><br>"
    Response.Write "<select name='SkinID' size='2' multiple style='height:300px;width:250px;'>"

    sql = "select * from PE_Skin"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何模板</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("SkinID") & "'>" & rs("SkinName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</select></td>"
    Response.Write "            <td width='80'><input type='submit' name='Submit' value='导入&gt;&gt;' "

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "></td>"
    Response.Write "            <td><strong>系统中已经存在的风格</strong><br>"
    Response.Write "             <select name='tSkinID' size='2' multiple style='height:300px;width:250px;' disabled>"

    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>没有任何模板</option>"
    Else

        Do While Not rs.EOF
            Response.Write "<option value='" & rs("SkinID") & "'>" & rs("SkinName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing

    Response.Write "              </select></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "     <br><b>提示：按住“Ctrl”或“Shift”键可以多选</b><br>"
    Response.Write "        <input name='SkinMdb' type='hidden' id='SkinMdb' value='" & mdbname & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "        <br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：ADD
'作  用：添加风格
'=================================================
Sub Add()
    Dim rs, sql, CssContent
    sql = "select * from PE_Skin where IsDefault=" & PE_True & ""
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
    Else
        CssContent = rs("Skin_CSS")
    End If

    rs.Close
    Set rs = Nothing

    Response.Write "<form name='myform' method='post' action='Admin_Skin.asp'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    Response.Write "      <td height='22' colspan='2'><strong>添加新风格</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='100'><strong>选择方案：</strong></td>"
    Response.Write "      <td> <select name='ProjectName' id='ProjectName'>" & GetProject_Option(ProjectName) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='100'><strong>风格名称：</strong></td>"
    Response.Write "      <td> <input name='SkinName' type='text' id='SkinName' value='' size='50' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='100'><strong>风格配色设置</strong><br>"
    Response.Write "        <br>"
    Response.Write "      修改风格设置必须具备一定网页设计知识<br><br>"
    Response.Write "      不能使用单引号或双引号，否则会容易造成程序错误</td>"
    Response.Write "      <td><textarea name='Skin_CSS' cols='80' rows='20' id='Skin_CSS'>" & CssContent & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td height='50' colspan='2'><input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input type='submit' name='Submit' value=' 添 加 '></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：Modify
'作  用：修改风格
'=================================================
Sub Modify()
    Dim SkinID, IsDefault
    Dim rs, sql
    SkinID = PE_CLng(Trim(Request.QueryString("SkinID")))
    IsDefault = Trim(Request.QueryString("IsDefault"))

    If SkinID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定SkinID</li>"
        Exit Sub
    End If
    
    If IsDefault = "" Then
        sql = "select * from PE_Skin where SkinID=" & SkinID
    Else
        sql = "select * from PE_Skin where IsDefault=" & PE_True
    End If

    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的风格！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<form name='myform' method='post' action='Admin_Skin.asp'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr align='center' class='title'>"
    Response.Write "      <td height='22' colspan='2'><strong>修改风格设置</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='100'><strong> 选择方案：</strong></td>"
    Response.Write "      <td><select name='ProjectName' id='ProjectName'"

    If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
        Response.Write " disabled"
    End If

    Response.Write ">" & GetProject_Option(ProjectName) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='100'><strong>风格名称：</strong></td>"
    Response.Write "      <td> <input name='SkinName' type='text' id='SkinName' value='" & rs("SkinName") & "' size='50' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='100'><strong>风格配色设置</strong><br>"
    Response.Write "        <br>"
    Response.Write "      修改风格设置必须具备一定网页设计知识<br><br>"
    Response.Write "      不能使用单引号或双引号，否则会容易造成程序错误</td>"
    Response.Write "      <td><textarea name='Skin_CSS' cols='80' rows='20' id='Skin_CSS'>" & rs("Skin_CSS") & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center' class='tdbg'>"
    Response.Write "      <td height='50' colspan='2'><input name='SkinID' type='hidden' id='SkinID' value='" & SkinID & "'><input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input type='submit' name='Submit' value=' 保存修改结果 '></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"

    rs.Close
    Set rs = Nothing
End Sub

'=================================================
'过程名：SaveAdd
'作  用：保存风格
'=================================================
Sub SaveAdd()
    Dim SkinName, Skin_CSS, ProjectName
    Dim rs, sql
    SkinName = Trim(Request("SkinName"))
    Skin_CSS = Trim(Request("Skin_CSS"))
    ProjectName = Trim(Request("ProjectName"))

    If ProjectName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>方案不能为空！</li>"
    End If

    If SkinName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>风格名称不能为空！</li>"
    End If

    If Skin_CSS = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>风格名称不能为空！</li>"
    End If

    If FoundErr = True Then Exit Sub
    
    sql = "select top 1 * from PE_Skin"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3
    rs.addnew
    rs("IsDefault") = False
    rs("SkinName") = SkinName
    rs("Skin_CSS") = Skin_CSS
    rs("ProjectName") = ProjectName
    rs.Update
    rs.Close
    Set rs = Nothing
    Call WriteSuccessMsg("成功添加新的风格：" & Trim(Request("SkinName")), ComeUrl)
    Call CreatSkinFile
End Sub

'=================================================
'过程名：SaveModify
'作  用：保存修改风格
'=================================================
Sub SaveModify()
    Dim rs, sql
    Dim SkinID, SkinName, Skin_CSS
    SkinID = PE_CLng(Trim(Request("SkinID")))
    SkinName = Trim(Request("SkinName"))
    Skin_CSS = Trim(Request("Skin_CSS"))
    
    If SkinID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定SkinID</li>"
    End If

    If SkinName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>风格名称不能为空！</li>"
    End If

    If Skin_CSS = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>风格名称不能为空！</li>"
    End If

    If FoundErr = True Then Exit Sub
    
    sql = "select * from PE_Skin where SkinID=" & SkinID
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的风格！</li>"
    Else
        rs("SkinName") = Trim(Request("SkinName"))
        rs("Skin_CSS") = Trim(Request("Skin_CSS"))
        rs.Update
        Call WriteSuccessMsg("保存风格设置成功！", ComeUrl)
    End If

    rs.Close
    Set rs = Nothing
    Call CreatSkinFile
    Call ClearSiteCache(0)
End Sub

'=================================================
'过程名：SetDefault
'作  用：保存指定默认风格
'=================================================
Sub SetDefault()
    Dim SkinID, DefaultType, setUpdateItem, setUpdateItem2, strTemp

    SkinID = PE_CLng(Trim(Request("SkinID")))
    DefaultType = PE_CLng(Trim(Request("DefaultType")))

    If SkinID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定SkinID!</li>"
        Exit Sub
    End If
        
    If DefaultType = 1 Then
        setUpdateItem = "IsDefault=" & PE_False & ",IsDefaultInProject=" & PE_False
        setUpdateItem2 = "IsDefault=" & PE_True & ",IsDefaultInProject=" & PE_True
        strTemp = "<li>成功将选定的风格,设置为<FONT style='font-size:12px' color='#008000'>系统默认</FONT>风格.</li>"
    ElseIf DefaultType = 2 Then
        setUpdateItem = "IsDefaultInProject=" & PE_False
        setUpdateItem2 = "IsDefaultInProject=" & PE_True
        strTemp = "<li>成功将选定的风格,设置为<FONT style='font-size:12px' color='#3366FF'>方案默认</FONT>风格.</li>"
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>设定的默认类型不对!</li>"
        Exit Sub
    End If

    Conn.Execute ("update PE_Skin set " & setUpdateItem & " where ProjectName='" & ProjectName & "'")
    Conn.Execute ("update PE_Skin set " & setUpdateItem2 & " where SkinID=" & SkinID & " and ProjectName='" & ProjectName & "'")
    Call WriteSuccessMsg(strTemp, ComeUrl)
    Call CreatSkinFile
    Call ClearSiteCache(0)
End Sub

'=================================================
'过程名：DelSkin
'作  用：删除指定风格
'=================================================
Sub DelSkin()
    Dim SkinID
    Dim rs, sql
    SkinID = Trim(Request("SkinID"))
	If IsValidID(SkinID) = False Then
		SkinID = ""
	End If

    If SkinID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定SkinID</li>"
        Exit Sub
    End If

    If InStr(SkinID, ",") > 0 Then
        sql = "select * from PE_Skin where SkinID In (" & SkinID & ")"
    Else
        sql = "select * from PE_Skin where SkinID=" & PE_CLng(SkinID)
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的风格！</li>"
    Else
        Do While Not rs.EOF
            If rs("IsDefaultInProject") = False And rs("IsDefault") = False Then
                Conn.Execute ("update PE_Channel set DefaultSkinID=0 where DefaultSkinID=" & rs("SkinID"))
                Conn.Execute ("update PE_Class set SkinID=0 where SkinID=" & rs("SkinID"))
                Conn.Execute ("update PE_Class set DefaultItemSkin=0 where DefaultItemSkin=" & rs("SkinID"))
                Conn.Execute ("update PE_Article set SkinID=0 where SkinID=" & rs("SkinID"))
                Conn.Execute ("update PE_Soft set SkinID=0 where SkinID=" & rs("SkinID"))
                Conn.Execute ("update PE_Photo set SkinID=0 where SkinID=" & rs("SkinID"))
                Conn.Execute ("update PE_Special set SkinID=0 where SkinID=" & rs("SkinID"))
                Call CreatSkinFile
                rs.Delete
                rs.Update
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Call WriteSuccessMsg("成功删除选定的风格。", ComeUrl)
End Sub

'=================================================
'过程名：DoExport
'作  用：导出风格处理
'=================================================
Sub DoExport()
    On Error Resume Next
    Dim rs
    Dim mdbname, tconn, trs
    Dim SkinID, FormatConn

    FormatConn = Request.Form("FormatConn")
    SkinID = Trim(Request("SkinID"))
    mdbname = Replace(Trim(Request.Form("skinmdb")), "'", "")
    If IsValidID(SkinID) = False Then
        SkinID = ""
    End If
    
    If SkinID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导出的模板</li>"
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

    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_Skin")
    End If

    Set rs = Conn.Execute("select * from PE_Skin where SkinID in (" & SkinID & ")  order by SkinID ")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Skin", tconn, 1, 3

    Do While Not rs.EOF
        trs.addnew
        trs("SkinName") = rs("SkinName")
        trs("Skin_CSS") = rs("Skin_CSS")
        trs("IsDefault") = False
        trs.Update
        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("已经成功将所选中的风格设置导出到指定的数据库中！<br><br>你还需要将Skin文件夹中图片文件一起打包。", ComeUrl)
End Sub

'=================================================
'过程名：DoImport
'作  用：导入风格处理
'=================================================
Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, trs
    Dim SkinID
    Dim rs
    SkinID = Trim(Request("SkinID"))
    mdbname = Replace(Trim(Request.Form("skinmdb")), "'", "")
    If IsValidID(SkinID) = False Then
        SkinID = ""
    End If

    If SkinID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导入的模板</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入模板数据库名"
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
    
    Set rs = tconn.Execute(" select * from PE_Skin where SkinID in (" & SkinID & ")  order by SkinID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Skin", Conn, 1, 3

    Do While Not rs.EOF
        trs.addnew
        trs("SkinName") = rs("SkinName")
        trs("Skin_CSS") = rs("Skin_CSS")
        trs("ProjectName") = ProjectName
        trs("IsDefault") = False
        trs.Update
        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("已经成功从指定的数据库中导入选中的风格！<br><br>你还需要将图片文件复制到Skin目录中的相应文件夹中才真正完成导入工作。", ComeUrl)
    Call CreatSkinFile
End Sub

'*************************  类模块主区域结束  *******************************
'*************************  类模块扩展域开始  *******************************
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

'*************************  类模块扩展域结束  *******************************
'*************************  类模块函数通用开始  *****************************
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
