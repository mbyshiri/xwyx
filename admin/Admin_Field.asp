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
Const PurviewLevel_Others = ""   '其他权限

Dim tempModuleType
    
If ChannelID = 0 Then
    FoundErr = True
    Response.Write "<li>频道参数丢失！</li>"
    Response.End
End If
tempModuleType = 0 - ModuleType

'检查管理员操作权限
If AdminPurview > 1 Then
    PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Field_" & ChannelDir)
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Response
    End If
End If
    
    
Response.Write "<html><head><title>字段管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle(ChannelName & "管理－－自定义字段管理", 10014)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td>"
Response.Write "<a href='Admin_Field.asp?ChannelID=" & ChannelID & "'>自定义字段管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Field.asp?ChannelID=" & ChannelID & "&Action=Add'>添加新字段</a>&nbsp;|&nbsp;"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Action = Trim(Request("Action"))
Select Case Action
Case "Add"
    Call Add
Case "SaveAdd"
    Call SaveAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Del"
    Call DelField
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Response.Write "<form name='myform' method='post' action=''>"
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='100' height='22'>字段名称</td>"
    Response.Write "    <td width='100'>字段标题</td>"
    Response.Write "    <td>附加提示</td>"
    Response.Write "    <td width='100'>调用标签</td>"
    Response.Write "    <td width='60'>字段类型</td>"
    Response.Write "    <td width='100'>默认值</td>"
    Response.Write "    <td width='50'>必填字段</td>"
    Response.Write "    <td width='80'>是否前台显示</td>"
    Response.Write "    <td width='70' align='center'>操作</td>"
    Response.Write "  </tr>"
    Dim sqlField, rsField
    sqlField = "select * from PE_Field where ChannelID=" & tempModuleType & " or ChannelID=" & ChannelID & " Order by FieldID"
    Set rsField = Conn.Execute(sqlField)
    Do While Not rsField.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td width='100' align='center'>" & rsField("FieldName") & "</td>"
        Response.Write "    <td width='100' align='center'>" & rsField("Title") & "</td>"
        Response.Write "    <td>" & PE_HTMLEncode(rsField("Tips")) & "</td>"
        Response.Write "    <td width='100' align='center'>" & rsField("LabelName") & "</td>"
        Response.Write "    <td width='60' align='center'>"
        Select Case rsField("FieldType")
        Case 1
            Response.Write "单行文本"
        Case 2
            Response.Write "多行文本"
        Case 3
            Response.Write "下拉列表"
        Case 4
            Response.Write "图片"
        Case 5
            Response.Write "文件"
        Case 6
            Response.Write "日期"
        Case 7
            Response.Write "数字"		
        Case 8
            Response.Write "单行文本(支持html)"	
        Case 9
            Response.Write "多行文本(支持html)"										
        End Select
        Response.Write "    </td>"
        Response.Write "    <td width='100' align='center'>" & rsField("DefaultValue") & "</td>"
        Response.Write "    <td width='50' align='center'>"
        If rsField("EnableNull") = True Then
            Response.Write "否"
        Else
            Response.Write "是"
        End If
        Response.Write "</td>"
        Response.Write "    <td width='80' align='center'>"
        If rsField("ShowOnForm") = True Then
            Response.Write "是"
        Else
            Response.Write "否"
        End If
        Response.Write "</td>"
        Response.Write "    <td width='70' align='center'>"
        Response.Write "<a href='Admin_Field.asp?ChannelID=" & ChannelID & "&Action=Modify&FieldID=" & rsField("FieldID") & "'>修改</a>&nbsp;&nbsp;"
        Response.Write "<a href='Admin_Field.asp?ChannelID=" & ChannelID & "&Action=Del&FieldID=" & rsField("FieldID") & "' onclick=""return confirm('真的要删除此字段吗？');"">删除</a>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        rsField.MoveNext
    Loop
    Response.Write "</table>"
    Response.Write "</form>"
    rsField.Close
    Set rsField = Nothing
End Sub

Sub Add()
    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write "  <!--" & vbCrLf
    Response.Write "  //检验文本框中内容是否超长" & vbCrLf
    Response.Write "    function CheckTextareaLength(val, max_length) {" & vbCrLf
    Response.Write "        var str_area=document.forms[0].elements[val].value;" & vbCrLf
    Response.Write "        if (str_area!=null&&str_area.length > max_length && document.myform.FieldType.value!=2 && document.myform.FieldType.value!=9){" & vbCrLf
    Response.Write "            alert(""文本文字超长，最多可输入"" + max_length +""个字符，请重新输入！"");" & vbCrLf
    Response.Write "            document.forms[0].elements[val].focus();" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function FieldCheckForm(FieldTypeValue){" & vbCrLf
    Response.Write "        if(FieldTypeValue=='3'){" & vbCrLf
    Response.Write "            trOptions.style.display='';" & vbCrLf
    Response.Write "            document.myform.DefaultValue.rows=1;" & vbCrLf
    Response.Write "        }else if(FieldTypeValue=='2'||FieldTypeValue=='9'){" & vbCrLf
    Response.Write "            trOptions.style.display='none';" & vbCrLf
    Response.Write "            document.myform.DefaultValue.rows=10;" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            trOptions.style.display='none';" & vbCrLf
    Response.Write "            document.myform.DefaultValue.rows=1;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    -->" & vbCrLf
    Response.Write "  </script>" & vbCrLf

    Response.Write "<form action='Admin_Field.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td colspan='2' align='center'><strong>添 加 新 字 段</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>使用范围：</strong></td>"
    Response.Write "      <td><input name='AreaType' type='radio' value='0'>所有同类频道&nbsp;&nbsp;&nbsp;&nbsp;<input name='AreaType' type='radio' value='1' checked>当前频道</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>字段名称：</strong><br>字段的英文名称，一般为英文。如更新时间字段的名称为“UpdateTime”<br><font color='red'>为了和系统字段区分，系统会自动在字段名前加上“MY_”</font></td>"
    Response.Write "      <td>MY_<input name='FieldName' type='text' id='FieldName' size='30' maxlength='20' value='' onchange=""document.myform.LabelName.value='{$MY_'+this.value+'}';""></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>字段标题：</strong><br>字段的中文标题，一般为中文。如“UpdateTime”字段的中文标题为“更新时间”</td>"
    Response.Write "      <td><input name='Title' type='text' id='Title' size='30' maxlength='30'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>附加提示：</strong><br>后台录入时，在标题旁的提示信息</td>"
    Response.Write "      <td><textarea name='Tips' cols='40' rows='3' id='Tips'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>调用标签：</strong><br>前台模板调用此字段内容的标签名。</td>"
    Response.Write "      <td><input name='LabelName' type='text' id='LabelName' size='30' maxlength='30' readonly></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>字段类型：</strong></td>"
    Response.Write "      <td><select name='FieldType' onchange=""javascript:FieldCheckForm(this.options[this.selectedIndex].value)"">" & GetFieldType(1) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>默认值：</strong></td>"
    Response.Write "      <td> <TEXTAREA Name='DefaultValue' ROWS='1' COLS='50' ONKEYPRESS=""javascript:CheckTextareaLength('DefaultValue',99);""></TEXTAREA></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' id='trOptions' style='display:none'>"
    Response.Write "      <td width='300'><strong>列表项目：</strong><br>每一行为一个列表项目</td>"
    Response.Write "      <td><textarea name='Options' cols='40' rows='3' id='Options'></textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>是否必填：</strong></td>"
    Response.Write "      <td><input name='EnableNull' type='radio' value='No'>是&nbsp;&nbsp;&nbsp;&nbsp;<input name='EnableNull' type='radio' value='Yes' checked>否</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>是否在前台显示：</strong></td>"
    Response.Write "      <td><input name='ShowOnForm' type='radio' value='No'>是&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' value='Yes' checked>否</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 添 加 ' onCLICK=""return CheckTextareaLength('DefaultValue',99);"">"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim FieldID, sqlField, rsField, JsConfig
    FieldID = PE_CLng(Trim(Request("FieldID")))
    If FieldID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    sqlField = "select * from PE_Field where FieldID=" & FieldID
    Set rsField = Conn.Execute(sqlField)
    If rsField.BOF And rsField.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的字段！</li>"
        rsField.Close
        Set rsField = Nothing
        Exit Sub
    End If
    
    Response.Write "<form action='Admin_Field.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td colspan='2' align='center'><strong>修 改 字 段 属 性</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>使用范围：</strong></td>"
    Response.Write "      <td><input name='AreaType' type='radio' value='0'"
    If rsField("ChannelID") = tempModuleType Then Response.Write " checked"
    Response.Write ">所有同类频道&nbsp;&nbsp;&nbsp;&nbsp;<input name='AreaType' type='radio' value='1'"
    If rsField("ChannelID") > 0 Then Response.Write " checked"
    Response.Write ">当前频道</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>字段名称：</strong><br>字段的英文名称，一般为英文。如更新时间字段的名称为“UpdateTime”</td>"
    Response.Write "      <td><input name='FieldName' type='text' id='FieldName' size='30' maxlength='20' value='" & rsField("FieldName") & "' disabled></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>字段标题：</strong><br>字段的中文标题，一般为中文。如“UpdateTime”字段的中文标题为“更新时间”</td>"
    Response.Write "      <td><input name='Title' type='text' id='Title' size='30' maxlength='30' value='" & rsField("Title") & "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>附加提示：</strong><br>后台录入时，在标题旁的提示信息</td>"
    Response.Write "      <td><textarea name='Tips' cols='40' rows='3' id='Tips'>" & rsField("Tips") & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>调用标签：</strong><br>前台模板调用此字段内容的标签名。</td>"
    Response.Write "      <td><input name='LabelName' type='text' id='LabelName' size='30' maxlength='30' value='" & rsField("LabelName") & "' readonly></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>字段类型：</strong></td>"
    Response.Write "      <td><select name='FieldType' disabled>" & GetFieldType(rsField("FieldType")) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>默认值：</strong></td><input name='FieldType' type='hidden' id='FieldType' value='" & rsField("FieldType") & "'>"
    Response.Write "      <td>"
    If rsField("FieldType") <> 2 Then
        Response.Write " <input name='DefaultValue' type='text' id='DefaultValue' size='30' maxlength='30' value='" & Server.HTMLEncode(rsField("DefaultValue")) & "'>"
    Else
        Response.Write " <TEXTAREA Name='DefaultValue' ROWS='10' COLS='50' >" & Server.HTMLEncode(rsField("DefaultValue")) & "</TEXTAREA>"
    End If
    Response.Write "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' id='trOptions'"
    If rsField("FieldType") <> 3 Then Response.Write " style='display:none'"
    Response.Write ">"
    Response.Write "      <td width='300'><strong>列表项目：</strong><br>每一行为一个列表项目</td>"
    Response.Write "      <td><textarea name='Options' cols='40' rows='3' id='Options'>" & rsField("Options") & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>是否必填：</strong></td>"
    Response.Write "      <td><input name='EnableNull' type='radio' value='No'"
    If rsField("EnableNull") = False Then Response.Write " checked"
    Response.Write ">是&nbsp;&nbsp;&nbsp;&nbsp;<input name='EnableNull' type='radio' value='Yes'"
    If rsField("EnableNull") = True Then Response.Write " checked"
    Response.Write ">否</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='300'><strong>是否在前台显示：</strong></td>"
    Response.Write "      <td><input name='ShowOnForm' type='radio' value='Yes'"
    If rsField("ShowOnForm") = True Then Response.Write " checked"
    Response.Write ">是&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' value='No'"
    If rsField("ShowOnForm") = False Then Response.Write " checked"
    Response.Write ">否</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'><input name='ChannelID' type='hidden' id='Action' value='" & ChannelID & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'><input name='FieldID' type='hidden' id='FieldID' value='" & FieldID & "'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 保存修改结果 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    
    rsField.Close
    Set rsField = Nothing
End Sub

Sub SaveAdd()
    Dim FieldName, Title, Tips, LabelName, FieldType, DefaultValue, Options, EnableNull,ShowOnForm
    Dim rsField, sqlField, trs, i
    FieldName = Replace(ReplaceBadChar(Trim(Request("FieldName"))), " ", "")
    Title = Trim(Request("Title"))
    Tips = Trim(Request("Tips"))
    FieldType = PE_CLng(Trim(Request("FieldType")))
    DefaultValue = Trim(Request("DefaultValue"))
    Options = Trim(Request("Options"))
    EnableNull = Trim(Request("EnableNull"))
    ShowOnForm = Trim(Request("ShowOnForm"))
    If FieldName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>字段名称不能为空！</li>"
    Else
		If IsValidStr(FieldName) = False Then
			FoundErr = True
			ErrMsg = ErrMsg & "<li>请输入有效的字段名称！</li>"
			Exit Sub
		End If
        FieldName = "MY_" & FieldName
        Set trs = Conn.Execute("select top 1 * from " & SheetName & "")
        For i = 0 To trs.Fields.Count - 1
            If trs.Fields(i).name = FieldName Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>指定的字段名称已经存在！</li>"
                Exit For
            End If
        Next
        Set trs = Nothing
    End If
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>字段标题不能为空！</li>"
    End If
    
    If (FieldType <> 2 Or FieldType <> 9 )  And Len(DefaultValue) > 99 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>默认值不能大于100个字符！</li>"
    End If
    LabelName = "{$" & FieldName & "}"
    
    If FieldType = 3 And Options = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入项目列表！</li>"
    End If
    If EnableNull = "Yes" Then
        EnableNull = True
    Else
        EnableNull = False
    End If
    If ShowOnForm = "Yes" Then
        ShowOnForm = True
    Else
        ShowOnForm = False
	End If
    If FoundErr = True Then Exit Sub
    
    If SystemDatabaseType = "SQL" Then
        If FieldType = 2 Then
            sqlField = "alter table " & SheetName & " add " & FieldName & " ntext null"
        Elseif FieldType = 7 Then
            sqlField = "alter table " & SheetName & " add " & FieldName & " integer null"		    
		Else
            sqlField = "alter table " & SheetName & " add " & FieldName & " nvarchar(255) null"
        End If
    Else
        If FieldType = 2 Then
            sqlField = "alter table " & SheetName & " add " & FieldName & " text null"
        Elseif FieldType = 7 Then
            sqlField = "alter table " & SheetName & " add " & FieldName & " integer null"				
		Else
            sqlField = "alter table " & SheetName & " add " & FieldName & " varchar(255) null"
        End If
    End If
    If Table_AddField(sqlField) = True Then
        sqlField = "select top 1 * from PE_Field"
        Set rsField = Server.CreateObject("ADODB.Recordset")
        rsField.Open sqlField, Conn, 1, 3
        rsField.addnew
        rsField("FieldName") = FieldName
        rsField("Title") = Title
        rsField("Tips") = Tips
        rsField("LabelName") = LabelName
        rsField("FieldType") = FieldType
        rsField("DefaultValue") = DefaultValue
        rsField("Options") = Options
        rsField("EnableNull") = EnableNull
        rsField("ShowOnForm") = ShowOnForm
        If PE_CLng(Trim(Request("AreaType"))) = 0 Then
            rsField("ChannelID") = tempModuleType
        Else
            rsField("ChannelID") = ChannelID
        End If
        rsField.Update
        rsField.Close
        Set rsField = Nothing
        Call CloseConn
        Response.Redirect "Admin_Field.asp?ChannelID=" & ChannelID
    End If
End Sub

Function Table_AddField(sqlField)
    On Error Resume Next
    Conn.Execute (sqlField)
    If Err Then
        Err.Clear
        FoundErr = True
        ErrMsg = ErrMsg & "<li>向" & SheetName & "表中添加字段失败！如果是SQL数据库，请检查数据库用户是否拥有Owner权限。</li>"
        Table_AddField = False
    Else
        Table_AddField = True
    End If
End Function

Sub SaveModify()
    Dim FieldID, Title, Tips, FieldType, DefaultValue, Options, EnableNull,ShowOnForm
    Dim rsField, sqlField, trs, i
    FieldID = PE_CLng(Trim(Request("FieldID")))
    Title = Trim(Request("Title"))
    Tips = Trim(Request("Tips"))
    FieldType = PE_CLng(Trim(Request("FieldType")))
    DefaultValue = Trim(Request("DefaultValue"))
    Options = Trim(Request("Options"))
    EnableNull = Trim(Request("EnableNull"))
    ShowOnForm = Trim(Request("ShowOnForm"))
    
    If FieldID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定字段ID！</li>"
    End If
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>字段标题不能为空！</li>"
    End If
    If FieldType = 3 And Options = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入项目列表！</li>"
    End If
    If EnableNull = "Yes" Then
        EnableNull = True
    Else
        EnableNull = False
    End If
    If ShowOnForm = "Yes" Then
        ShowOnForm = True
    Else
        ShowOnForm = False
    End If
    If FoundErr = True Then Exit Sub
    
    sqlField = "select top 1 * from PE_Field where FieldID=" & FieldID
    Set rsField = Server.CreateObject("ADODB.Recordset")
    rsField.Open sqlField, Conn, 1, 3
    If rsField.BOF And rsField.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的字段记录！</li>"
        rsField.Close
        Set rsField = Nothing
        Exit Sub
    End If
    rsField("Title") = Title
    rsField("Tips") = Tips
    rsField("DefaultValue") = DefaultValue
    rsField("Options") = Options
    rsField("EnableNull") = EnableNull
    rsField("ShowOnForm") = ShowOnForm
    If PE_CLng(Trim(Request("AreaType"))) = 0 Then
        rsField("ChannelID") = tempModuleType
    Else
        rsField("ChannelID") = ChannelID
    End If
    rsField.Update
    rsField.Close
    Set rsField = Nothing
    Call CloseConn
    Response.Redirect "Admin_Field.asp?ChannelID=" & ChannelID
End Sub

Sub DelField()
    Dim FieldID, sqlField, rsField
    FieldID = PE_CLng(Trim(Request("FieldID")))
    If FieldID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    sqlField = "select * from PE_Field where FieldID=" & FieldID
    Set rsField = Server.CreateObject("ADODB.Recordset")
    rsField.Open sqlField, Conn, 1, 3
    If rsField.BOF And rsField.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的字段！</li>"
        rsField.Close
        Set rsField = Nothing
        Exit Sub
    End If
    On Error Resume Next
    Conn.Execute ("alter table " & SheetName & " drop COLUMN " & rsField("FieldName") & "")
    If Err Then
        Err.Clear
        FoundErr = True
        ErrMsg = ErrMsg & "<li>无法从" & SheetName & "表中删除字段。如果是SQL数据库，请检查是否有足够权限。</li>"
    Else
        rsField.Delete
        rsField.Update
    End If
    rsField.Close
    Set rsField = Nothing
    Call CloseConn
    If FoundErr <> True Then
        Response.Redirect "Admin_Field.asp?ChannelID=" & ChannelID
    End If
End Sub

Function GetFieldType(FieldType)
    Dim strFieldType
    strFieldType = "<option value='1'"
    If FieldType = 1 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">单行文本</option>"	
    strFieldType = strFieldType & "<option value='8'"
    If FieldType = 8 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">单行文本(支持html)</option>"	
    strFieldType = strFieldType & "<option value='2'"
    If FieldType = 2 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">多行文本</option>"
    strFieldType = strFieldType & "<option value='9'"
    If FieldType = 9 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">多行文本(支持html)</option>"	
    strFieldType = strFieldType & "<option value='3'"
    If FieldType = 3 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">下拉列表</option>"
    strFieldType = strFieldType & "<option value='4'"
    If FieldType = 4 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">图片</option>"
    strFieldType = strFieldType & "<option value='5'"
    If FieldType = 5 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">文件</option>"
    strFieldType = strFieldType & "<option value='6'"
    If FieldType = 6 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">日期</option>"
    strFieldType = strFieldType & "<option value='7'"
    If FieldType = 7 Then strFieldType = strFieldType & " selected"
    strFieldType = strFieldType & ">数字</option>"	
    GetFieldType = strFieldType
End Function
%>
