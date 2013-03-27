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

Dim AddType


'检查管理员操作权限
If AdminPurview > 1 Then
    PurviewPassed = CheckPurview_Other(AdminPurview_Others, "JsFile_" & ChannelDir)
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Call WriteEntry(6, AdminName, "越权操作")
        Response.End
    End If
End If

AddType = PE_CLng(Trim(Request("AddType")))

Response.Write "<html><head><title>JS代码管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle(ChannelName & "管理----JS文件管理", 10112)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td>"
Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "'>JS文件管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=0'>添加新的JS文件（普通列表方式）</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=1'>添加新的JS文件（图片列表方式）</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Class.asp?ChannelID=" & ChannelID & "&Action=CreateJS'>刷新栏目JS文件</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Special.asp?ChannelID=" & ChannelID & "&Action=CreateJS'>刷新专题JS文件</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    If AddType = 0 Then
        Call Add
    Else
        Call AddPic
    End If
Case "Modify"
    Call Modify
Case "ModifyPic"
    Call ModifyPic
Case "SaveAdd", "SaveModify"
    Call SaveJS_List
Case "SaveAddPic", "SaveModifyPic"
    Call SaveJS_Pic
Case "Preview"
    Call PreviewJS
Case "Del"
    Call DelJS
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Response.Write "<form name='myform' method='post' action='Admin_CreateJS.asp'>"
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='100' height='22'>JS代码名称</td>"
    Response.Write "    <td>简介</td>"
    Response.Write "    <td width='60'>代码类型</td>"
    Response.Write "    <td>JS文件名</td>"
    Response.Write "    <td width='260'>JS调用代码</td>"
    Response.Write "    <td width='100' align='center'>操作</td>"
    Response.Write "  </tr>"
    Dim sqlJs, rsJs, JsExists
    sqlJs = "select * from PE_JsFile where ChannelID=" & ChannelID & ""
    Set rsJs = Conn.Execute(sqlJs)
    Do While Not rsJs.EOF
        JsExists = False
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td width='100' align='center'>" & rsJs("JsName") & "</td>"
        Response.Write "    <td>" & PE_HTMLEncode(rsJs("JsReadMe")) & "</td>"
        Response.Write "    <td width='60' align='center'>"
        Select Case rsJs("JsType")
        Case 0
            Response.Write "普通列表"
        Case 1
            Response.Write "图片列表"
        End Select
        Response.Write "    </td>"
        Response.Write "    <td>"
        If ObjInstalled_FSO = True Then
            If fso.FileExists(Server.MapPath(InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName"))) Then
                JsExists = True
            End If
            If JsExists = True Then
                Response.Write rsJs("JsFileName")
            Else
                Response.Write "<font color='red'>" & rsJs("JsFileName") & "</font>"
            End If
        Else
            Response.Write rsJs("JsFileName")
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='260'><textarea name='textarea' cols='36' rows='3'>"
        If rsJs("ContentType") = 1 Then
            Response.Write "<!--#include File=""" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & """-->"
        Else
            Response.Write "<script language='javascript' src='" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & "'></script>"
        End If
        Response.Write "</textarea></td>"
        Response.Write "    <td width='100' align='center'>"
        If rsJs("JsType") = 0 Then
            Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=Modify&ID=" & rsJs("ID") & "'>参数设置</a>&nbsp;&nbsp;"
        Else
            Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=ModifyPic&ID=" & rsJs("ID") & "'>参数设置</a>&nbsp;&nbsp;"
        End If
        Response.Write "<a href='Admin_CreateJS.asp?ChannelID=" & ChannelID & "&Action=CreateJs&ID=" & rsJs("ID") & "'>刷新</a><br>"
        If JsExists = True Then
            Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=Preview&ID=" & rsJs("ID") & "'>预览效果</a>&nbsp;&nbsp;"
        Else
            Response.Write "<font color='gray'>预览效果</font>&nbsp;&nbsp;"
        End If
        Response.Write "<a href='Admin_" & ModuleName & "Js.asp?ChannelID=" & ChannelID & "&Action=Del&ID=" & rsJs("ID") & "' onclick=""return confirm('真的要删除此JS文件吗？如果有文件或模板中使用此JS文件，请注意修改过来呀！');"">删除</a>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        rsJs.MoveNext
    Loop
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellspacing='5' cellpadding='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td align='center'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='CreateAllJs'><input name='ShowBack' type='hidden' id='ShowBack' value='Yes'>"
    Response.Write "    <input type='submit' name='Submit' value='刷新所有JS文件'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<b>说明：</b><br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;这些JS代码是为了加快访问速度特别生成的。在添加/修改/审核/删除" & ChannelShortName & "时，系统会自动刷新各JS文件。必要时，你也可以手动刷新。如添加了新的JS文件，但还没有添加" & ChannelShortName & "，此时就可以手动刷新有关JS文件。<br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>若文件名为红色，表示此JS文件还没有生成。</font><br>"
    Response.Write "<b>使用方法：</b><br>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;将相关JS调用代码复制到页面或模板中的相关位置即可。可参见系统提供的各页面及模板。"
    Response.Write "</form>"
    rsJs.Close
    Set rsJs = Nothing
End Sub

'******************************************
'过程名：JsBaseInif
'作  用：Js管理基本信息
'******************************************
Sub JsBaseInif(ByVal JsName, ByVal JsReadme, ByVal ContentType, ByVal JsFileName)
    ContentType = PE_CLng(ContentType)
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>代码名称：</strong></td>"
    Response.Write "      <td height='25'><input name='JsName' type='text' id='JsName' value='" & JsName & "' size='49' maxlength='50'> <font color='#FF0000'>*</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>简介：</strong></td>"
    Response.Write "      <td height='25'><textarea name='JsReadme' cols='40' rows='3' id='JsReadme'>" & JsReadme & "</textarea></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>内容代码格式：</strong></td>"
    Response.Write "      <td height='25'> <Input TYPE='radio' Name='ContentType' value='0' " & RadioValue(ContentType, 0) & " onClick=""htmltype.style.display='none';jstype.style.display=''""> JS <Input TYPE='radio' Name='ContentType' value='1' " & RadioValue(ContentType, 1) & "  onClick=""htmltype.style.display='';jstype.style.display='none'""> Html <FONT color='blue'>注意：频道选择生成Shtml方式时可选用此项，可以在扩展名为.shtml的文件中使用<br>&lt;!--#include file=""aaaa.html""--&gt;这样的指令包含其他文件，这样对搜索引擎比使用JS代码调用会更友好。 </font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right'><strong>文件名：</strong></td>"
    Response.Write "      <td height='25'><input name='JsFileName' type='text' id='JsFileName'  value='" & JsFileName & "' size='49' maxlength='50'> <font color='#FF0000'>*</font>"
    Response.Write "       <Span Id='jstype' style=""display:"
    If ContentType = 0 Then
        Response.Write "''"
    Else
        Response.Write "'none'"
    End If
    Response.Write """>"
    Response.Write "<font color='red'>以.js为扩展名</font></Span>"
    Response.Write "<Span Id='htmltype' style=""display:"
    If ContentType = 1 Then
        Response.Write "''"
    Else
        Response.Write "'none'"
    End If
    Response.Write """><font color='red'>以.html为扩展名</font></Span></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg2'>"
    Response.Write "      <td height='25' colspan='2' align='center' ><strong>参数设置</strong></td>"
    Response.Write "    </tr>"
End Sub

Sub PreviewJS()
    Dim ID, sqlJs, rsJs
    ID = Trim(Request("ID"))
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    Else
        ID = PE_CLng(ID)
    End If
    sqlJs = "select * from PE_JsFile where ID=" & ID
    Set rsJs = Conn.Execute(sqlJs)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>预览JS文件效果----" & rsJs("JsName") & "</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='center'>"

    If rsJs("ContentType") = 1 Then
        Response.Write "<iframe marginwidth=0 marginheight=0 frameborder=0 name='libin' width='700' height='400' src=" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & "></iframe>"
    Else
        Response.Write "<script language='javascript' src='" & InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName") & "'></script>"
    End If

    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='center'>"
    Response.Write "        <a href='javascript:this.location.reload();'>刷新本页</a>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <a href='Admin_ArticleJS.asp?ChannelID=" & ChannelID & "'>返回上页</a>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"

    rsJs.Close
    Set rsJs = Nothing
End Sub

Sub DelJS()
    Dim ID, sqlJs, rsJs, tJsFileName
    ID = Trim(Request("ID"))
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    Else
        ID = PE_CLng(ID)
    End If
    sqlJs = "select * from PE_JsFile where ID=" & ID
    Set rsJs = Server.CreateObject("ADODB.Recordset")
    rsJs.Open sqlJs, Conn, 1, 3
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的JS文件！</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If
    If ObjInstalled_FSO = True Then
        tJsFileName = Server.MapPath(InstallDir & ChannelDir & "/JS/" & rsJs("JsFileName"))
        If fso.FileExists(tJsFileName) Then
            fso.DeleteFile tJsFileName
        End If
    End If
    rsJs.Delete
    rsJs.Update
    rsJs.Close
    Set rsJs = Nothing
    Call CloseConn
    Response.Redirect "Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID
End Sub

Sub CreateJS(ID)
    Response.Write "<br><iframe id='CreateJS' width='100%' height='100' frameborder='0' src='Admin_CreateJS.asp?ChannelID=" & ChannelID & "&Action=CreateJs&ID=" & ID & "'></iframe>"
End Sub

Function GetClass_Option(CurrentID)
    Dim rsClass, sqlClass, strTemp, tmpDepth, i
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    sqlClass = "Select ClassID,ClassName,ClassType,Depth,NextID from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
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

%>
