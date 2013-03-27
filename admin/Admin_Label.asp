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
Const PurviewLevel_Others = "Label"   '其他权限

Response.Write "<html><head><title>标签管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Call ShowPageTitle("标 签 管 理", 10026)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td>"
Response.Write "<a href='Admin_Label.asp'>标签管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Label.asp?Action=AddStat'>添加静态标签</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Label.asp?Action=AddDyna'>添加动态标签</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Label.asp?Action=AddDyna&addtype=3'>添加函数标签</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Label.asp?Action=AddCai'>添加采集标签</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Label.asp?Action=import'>导入标签</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Label.asp?Action=export'>导出标签</a>&nbsp;|&nbsp;"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "AddStat"
    Call Add(0)
Case "AddCai"
    Call Add(2)
Case "AddDyna"
    Call AddDyna
Case "AddDyna2"
    Call AddDyna2
Case "SaveAdd"
    Call Save
Case "Modify"
    Call Modify
Case "SaveModify"
    Call Save
Case "Del"
    Call DelLabel
Case "import"
    Call Import
Case "import2"
    Call import2
Case "Doimport"
    Call DoImport
Case "export"
    Call Export
Case "Doexport"
    Call DoExport
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim sqlLabel, rsLabel, ListType, rsLabelClass, ClassType
    Dim iCount
    ListType = PE_CLng(Trim(Request("ListType")))
    strFileName = "Admin_Label.asp?ListType=" & ListType
    ClassType = ReplaceBadChar(Trim(Request("ClassType")))
    If ClassType <> "" Then
        strFileName = strFileName & "&ClassType=" & ClassType
    End If
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'><tr class='title'><td height='22'>|&nbsp;<a href='Admin_Label.asp'>"
    If ListType = 0 Then
        Response.Write "<font color=red>静态标签</font> "
    Else
        Response.Write "静态标签 "
    End If
    Response.Write "</a> | <a href='Admin_Label.asp?ListType=1'>"
    If ListType = 1 Then
        Response.Write "<font color=red>动态标签</font> "
    Else
        Response.Write "动态标签 "
    End If
    Response.Write "</a> | <a href='Admin_Label.asp?ListType=3'>"
    If ListType = 3 Then
        Response.Write "<font color=red>函数标签</font> "
    Else
        Response.Write "函数标签 "
    End If
    Response.Write "</a> | <a href='Admin_Label.asp?ListType=2'>"
    If ListType = 2 Then
        Response.Write "<font color=red>采集标签</font> "
    Else
        Response.Write "采集标签 "
    End If
    Response.Write "</a> | </td></tr></table><br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'><tr class='title'><td height='22'>"
    If ClassType = "" Then
        Response.Write "|&nbsp;<font color=red>全部分类</font>&nbsp;|"
    Else
        Response.Write "|&nbsp;<a href='Admin_Label.asp?ListType=" & ListType & "'>全部分类</a>&nbsp;|"
    End If
    Set rsLabelClass = Conn.Execute("select LabelClass from PE_Label Where LabelType=" & ListType & " GROUP BY LabelClass")
    Do While Not rsLabelClass.EOF
        If ClassType <> "" And ClassType = rsLabelClass(0) Then
            Response.Write "&nbsp;<font color=red>" & rsLabelClass(0) & "</font>&nbsp;|"
        Else
            If Trim(rsLabelClass(0) & "") <> "" Then
                Response.Write "&nbsp;<a href='Admin_Label.asp?ListType=" & ListType & "&ClassType=" & rsLabelClass(0) & "'>" & rsLabelClass(0) & "</a>&nbsp;|"
            End If
        End If
        rsLabelClass.MoveNext
    Loop
    Set rsLabelClass = Nothing
    Response.Write "</td></tr></table><br>"

    Response.Write "<form name='myform' method='post' action=''>"
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='150' height='22'>标签名称</td>"
    Response.Write "    <td width='40'>优先级</td>"
    Response.Write "    <td width='80'>标签分类</td>"
    Response.Write "    <td width='60'>标签类型</td>"
    If ListType = 0 Then
        Response.Write "    <td>标签简介</td>"
    ElseIf ListType = 1 Or ListType = 3 Then
        Response.Write "    <td>查询语句</td>"
    ElseIf ListType = 2 Then
        Response.Write "    <td>连接地址</td>"
    End If
    Response.Write "    <td width='70' align='center'>操作</td>"
    Response.Write "  </tr>"
    
   
    Set rsLabel = Server.CreateObject("Adodb.RecordSet")
    sqlLabel = "select * from PE_Label Where LabelType=" & ListType
    If ClassType <> "" Then sqlLabel = sqlLabel & " and LabelClass='" & ClassType & "'"
    sqlLabel = sqlLabel & " Order by Priority asc,LabelID asc"
    rsLabel.Open sqlLabel, Conn, 1, 1
    If rsLabel.BOF And rsLabel.EOF Then
        rsLabel.Close
        Set rsLabel = Nothing
        Select Case ListType
        Case 1
            Response.Write "<tr><td colspan='6' align='center'>尚未添加动态自定义标签！</td></tr></table></form>"
        Case 2
            Response.Write "<tr><td colspan='6' align='center'>尚未添加采集自定义标签！</td></tr></table></form>"
        Case 3
            Response.Write "<tr><td colspan='6' align='center'>尚未添加动态函数标签！</td></tr></table></form>"
        Case Else
            Response.Write "<tr><td colspan='6' align='center'>尚未添加静态自定义标签！</td></tr></table></form>"
        End Select
        Exit Sub
    Else
        totalPut = rsLabel.RecordCount
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
                rsLabel.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Do While Not rsLabel.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td align='left'>{$" & rsLabel("LabelName") & "}</td>"
            Response.Write "    <td align='center'>" & rsLabel("Priority") & "</td>"
            Response.Write "    <td align='center'>" & rsLabel("LabelClass") & "</td>"
            If rsLabel("LabelType") = 0 Then
                Response.Write "    <td align='center'>静态标签</td>"
            ElseIf rsLabel("LabelType") = 1 Then
                Response.Write "    <td align='center'>动态标签</td>"
            ElseIf rsLabel("LabelType") = 2 Then
                Response.Write "    <td align='center'>采集标签</td>"
            ElseIf rsLabel("LabelType") = 3 Then
                Response.Write "    <td align='center'>函数标签</td>"
            End If
            Response.Write "    <td style='word-break:break-all;Width:fixed'><a href='Admin_Label.asp?Action=Modify&LabelID=" & rsLabel("LabelID") & "'>" & PE_HTMLEncode(rsLabel("LabelIntro")) & "</a></td>"
            Response.Write "    <td align='center'>"
            Response.Write "<a href='Admin_Label.asp?Action=Modify&LabelID=" & rsLabel("LabelID") & "'>修改</a>&nbsp;&nbsp;"
            Response.Write "<a href='Admin_Label.asp?Action=Del&LabelID=" & rsLabel("LabelID")
            If ListType > 0 Then
                Response.Write "&ListType=" & ListType
            End If
            Response.Write "' onclick=""return confirm('真的要删除此标签吗？如果有文件或模板中使用此标签，请注意修改过来呀！');"">删除</a>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            iCount = iCount + 1
            If iCount >= MaxPerPage Then Exit Do
            rsLabel.MoveNext
        Loop
        Response.Write "</table></form>"
        rsLabel.Close
        Set rsLabel = Nothing
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个标签", True)
    End If
End Sub

Sub ShowJSLabel(LabelType)
    Dim TrueSiteUrl
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function addclass(){" & vbCrLf
    Response.Write "    var select=document.myform.LabelClassList;" & vbCrLf
    Response.Write "    for(i=0;i<select.length;i++){" & vbCrLf
    Response.Write "        if(document.myform.LabelClassList[i].selected==true){" & vbCrLf
    Response.Write "            document.myform.LabelClass.value=document.myform.LabelClassList[i].value;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.LabelName.value==''){" & vbCrLf
    Response.Write "     alert('标签名称不能为空！');" & vbCrLf
    Response.Write "     document.myform.LabelName.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.LabelIntro.value==''){" & vbCrLf
    If LabelType = 0 Then
        Response.Write "     alert('标签简介不能为空！');" & vbCrLf
    ElseIf LabelType = 1 Or LabelType = 3 Then
        Response.Write "     alert('查询语句不能为空！');" & vbCrLf
    Else
        Response.Write "     alert('连接地址不能为空！');" & vbCrLf
    End If
    Response.Write "     document.myform.LabelIntro.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Priority.value==''){" & vbCrLf
    Response.Write "     alert('优先等级不能为空！');" & vbCrLf
    Response.Write "     document.myform.Priority.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf

    If LabelType <> 2 Then
        Response.Write "  document.myform.LabelContent2.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
        Response.Write "  if (document.myform.LabelContent.value==''){" & vbCrLf
        Response.Write "     alert('标签内容不能为空！');" & vbCrLf
        Response.Write "     document.myform.LabelContent.focus();" & vbCrLf
        Response.Write "     return false;" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  if (Strsave==""B""){" & vbCrLf
        Response.Write "      setContent (""get"",1);" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  return true;  " & vbCrLf
        Response.Write "}" & vbCrLf

        If (LabelType = 1 Or LabelType = 3) And (Action = "AddDyna2" Or Action = "Modify") Then
            Response.Write "function addfield(fieldname,num,dbname,dbtype){" & vbCrLf
            Response.Write "    myform.LabelContent.focus();" & vbCrLf
            Response.Write "    var str = document.selection.createRange();" & vbCrLf
            Response.Write "    var link=""Admin_pfield.asp?fieldname="" + fieldname + ""&num=""+ num + ""&dbname="" + dbname +""&dbtype="" + dbtype;" & vbCrLf
            Response.Write "    var arr=showModalDialog(link,'','dialogWidth:300px; dialogHeight:180px; help: no; scroll: no; status: no');" & vbCrLf
            Response.Write "    if (arr != null){" & vbCrLf
            Response.Write "        str.text = arr;" & vbCrLf
            Response.Write "        document.myform.LabelContent2.value=document.myform.LabelContent.value;" & vbCrLf
            Response.Write "        editor.HtmlEdit.document.body.innerHTML=document.myform.LabelContent2.value;" & vbCrLf
            Response.Write "    }" & vbCrLf
            Response.Write "}" & vbCrLf
            Response.Write "function addfield2(fiele1){" & vbCrLf
            Response.Write "    myform.LabelContent.focus();" & vbCrLf
            Response.Write "    var str = document.selection.createRange();" & vbCrLf
            Response.Write "    if (fiele1 != null){" & vbCrLf
            Response.Write "        str.text = ""{input("" + fiele1 + "")}"";" & vbCrLf
            Response.Write "        document.myform.LabelContent2.value=document.myform.LabelContent.value;" & vbCrLf
            Response.Write "        editor.HtmlEdit.document.body.innerHTML=document.myform.LabelContent2.value;" & vbCrLf
            Response.Write "    }" & vbCrLf
            Response.Write "}" & vbCrLf
        End If
        Response.Write "</script>" & vbCrLf

        Response.Write "<script language=""VBScript"">" & vbCrLf
        Response.Write "    Dim regEx, Match, Matches, StrBody,strTemp,strMatch,arrMatch,i,Strsave" & vbCrLf
        Response.Write "    Dim Content,arrContent" & vbCrLf
        Response.Write "    Set regEx = New RegExp" & vbCrLf
        Response.Write "    regEx.IgnoreCase = True" & vbCrLf
        Response.Write "    regEx.Global = True" & vbCrLf
        Response.Write "    Strsave=""A""" & vbCrLf
        '=================================================
        '作  用：排序html
        '=================================================
        Response.Write "Function  Resumeblank(byval Content)" & vbCrLf
        Response.Write " Dim strHtml,strHtml2,Num,Numtemp,Strtemp" & vbCrLf
        Response.Write "   strHtml=Replace(Content, ""<DIV"", ""<div"")" & vbCrLf
        Response.Write "   strHtml=Replace(strHtml, ""</DIV>"", ""</div>"")" & vbCrLf
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
        Response.Write "      Next" & vbCrLf
        Response.Write "      Resumeblank=strHtml2" & vbCrLf
        Response.Write "    End function" & vbCrLf
        Response.Write "    function setContent(zhi,TpyeTemplate)" & vbCrLf
        Response.Write "      if zhi=""get"" then" & vbCrLf
        Response.Write "        if Strsave=""A"" then Exit Function" & vbCrLf
        Response.Write "        Strsave=""A""" & vbCrLf
        Response.Write "        Content= editor.HtmlEdit.document.body.innerHTML" & vbCrLf
        Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\}['|""""]\>""" & vbCrLf
        Response.Write "        Set Matches = regEx.Execute(Content)" & vbCrLf
        Response.Write "        For Each Match In Matches" & vbCrLf
        Response.Write "            regEx.Pattern = ""\{\$(.*?)\}""" & vbCrLf
        Response.Write "            Set strTemp = regEx.Execute(replace(Match.Value,"" "",""""))" & vbCrLf
        Response.Write "            For Each Match2 In strTemp" & vbCrLf
        Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
        Response.Write "                Content = Replace(Content, Match.Value, ""<!--"" & strTemp2 & ""-->"")" & vbCrLf
        Response.Write "            Next" & vbCrLf
        Response.Write "        Next" & vbCrLf
        Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\$\>""" & vbCrLf
        Response.Write "        Set Matches = regEx.Execute(Content)" & vbCrLf
        Response.Write "        For Each Match In Matches" & vbCrLf
        Response.Write "        regEx.Pattern = ""\#\[(.*?)\]\#""" & vbCrLf
        Response.Write "        Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
        Response.Write "            For Each Match2 In strTemp" & vbCrLf
        Response.Write "                strTemp2 = Replace(Match2.Value, ""&amp;"", ""&"")" & vbCrLf
        Response.Write "                strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
        Response.Write "                strTemp2 = Replace(strTemp2,""&13;&10;"",vbCrLf)" & vbCrLf
        Response.Write "                strTemp2 = Replace(strTemp2,""&9;"",vbTab)" & vbCrLf
        Response.Write "                strTemp2 = Replace(strTemp2,""′"",""'"")" & vbCrLf
        Response.Write "                strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
        Response.Write "                strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
        Response.Write "                Content = Replace(Content, Match.Value, strTemp2)" & vbCrLf
        Response.Write "            Next" & vbCrLf
        Response.Write "         Next" & vbCrLf
        Response.Write "        Content=Replace(Content, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
        Response.Write "        Content=Replace(Content, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf
        Response.Write "        Content=Resumeblank(Content)" & vbCrLf
        Response.Write "        Content=Replace(Content,""{$InstallDir}{$rsClass_ClassUrl}"",""{$rsClass_ClassUrl}"")" & vbCrLf
        Response.Write "        regEx.Pattern = ""\{\$InstallDir\}editor.asp(.[^\<]*?)\#""" & vbCrLf
        Response.Write "        Content = regEx.Replace(Content, ""#"")" & vbCrLf
        Response.Write "        document.myform.LabelContent.value=Content" & vbCrLf
        Response.Write "    Else" & vbCrLf
        Response.Write "        if Strsave=""B"" then Exit Function" & vbCrLf
        Response.Write "        Strsave=""B""" & vbCrLf
        Response.Write "        Content= document.myform.LabelContent.value" & vbCrLf
        Response.Write "        if Content="""" then " & vbCrLf
        Response.Write "            alert ""您删除了代码框网页，请您务必填写网页 ！""" & vbCrLf
        Response.Write "            Exit function" & vbCrLf
        Response.Write "           " & vbCrLf
        Response.Write "        End if" & vbCrLf
        Response.Write "        if Instr(TemplateContent,""<body>"") <> 0 then" & vbCrLf
        Response.Write "            alert ""您加载的自定义标签包含<body> ！这是不能牵套在自定义标签内的！""" & vbCrLf
        Response.Write "            Exit function" & vbCrLf
        Response.Write "        End if" & vbCrLf
        Response.Write "        Content = Replace(Content, ""<!--{$"", ""{$"")" & vbCrLf
        Response.Write "        Content = Replace(Content, ""}-->"", ""}"")" & vbCrLf
        Response.Write "        '图片替换JS" & vbCrLf
        Response.Write "        regEx.Pattern = ""(\<Script)([\s\S]*?)(\<\/Script\>)""" & vbCrLf
        Response.Write "        Set Matches = regEx.Execute(Content)" & vbCrLf
        Response.Write "        For Each Match In Matches" & vbCrLf
        Response.Write "            strTemp = Replace(Match.Value, ""<"", ""[!"")" & vbCrLf
        Response.Write "            strTemp = Replace(strTemp, "">"", ""!]"")" & vbCrLf
        Response.Write "            strTemp = Replace(strTemp, ""'"", ""′"")" & vbCrLf
        Response.Write "            strTemp = ""<IMG alt='#"" & strTemp & ""#' src=""""" & InstallDir & "editor/images/jscript.gif"""" border=0 $>""" & vbCrLf
        Response.Write "            Content = Replace(Content, Match.Value, strTemp)" & vbCrLf
        Response.Write "        Next" & vbCrLf
        Response.Write "        '图片替换超级标签" & vbCrLf
        Response.Write "        regEx.Pattern = ""(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct)\((.*?)\)\}""" & vbCrLf
        Response.Write "        Content = regEx.Replace(Content, ""<IMG src=""""" & InstallDir & "editor/images/label.gif"""" border=0 zzz='$1($2)}'>"")" & vbCrLf
        Response.Write "        regEx.Pattern = ""\{\$InstallDir\}""" & vbCrLf
        Response.Write "        Set Matches = regEx.Execute(Content)" & vbCrLf
        Response.Write "        For Each Match In Matches" & vbCrLf
        Response.Write "            Content = Replace(Content, Match.Value, ""http://" & TrueSiteUrl & InstallDir & """)" & vbCrLf
        Response.Write "        Next" & vbCrLf
        Response.Write "        editor.HtmlEdit.document.body.innerHTML=Content" & vbCrLf
        Response.Write "        editor.showBorders()" & vbCrLf
        Response.Write "    End if" & vbCrLf
        Response.Write "    End function" & vbCrLf
        Response.Write "    function setstatus()" & vbCrLf '为323 版兼容editor.asp 无效过程
        Response.Write "    end function" & vbCrLf
        Response.Write "</script>" & vbCrLf
    Else
        Response.Write "  return true;  " & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "</script>" & vbCrLf
    End If
End Sub

Sub Add(AddType)
    Call ShowJSLabel(AddType)
    Response.Write "<form action='Admin_Label.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    If AddType = 0 Then
        Response.Write "      <td align='center'><strong>添 加 静 态 标 签</strong></td>"
    ElseIf AddType = 2 Then
        Response.Write "      <td align='center'><strong>添 加 采 集 标 签 <font color=#aaffaa>（本标签类型消耗资源较大，推荐独立服务器用户使用）</font></strong></td>"
    End If
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签名称：</strong></td>"
    Response.Write "          <td>{$MY_<input name='LabelName' type='text' id='LabelName' size='30' maxlength='50'>}</td>"
    Response.Write "          <td width='10'></td>"
    Response.Write "          <td><font color='#FF0000'>* 输入名称（英文要注意大小写）即可，不用输入定界符。</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签分类：</strong></td>"
    Response.Write "          <td colspan='3'><input name='LabelClass' type='text' id='LabelClass' size='30' maxlength='50'> " & getlabelclass(AddType) & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    If AddType = 0 Then
        Response.Write "         <tr><td width='100' align='center'><strong>标签简介：</strong></td>"
        Response.Write "         <td><textarea name='LabelIntro' cols='96' rows='4' id='LabelIntro'></textarea></td></tr>"
        Response.Write "  </table></tr>" & vbCrLf
    ElseIf AddType = 2 Then
        Response.Write "         <tr><td width='100' align='center'><strong>采集方式：</strong></td>"
        Response.Write "         <td><INPUT TYPE='radio' NAME='CaiType' value='0' checked onClick=""adv.style.display='none'"">快速 <INPUT TYPE='radio' NAME='CaiType' value='1' onClick=""adv.style.display=''"">高级</td></tr>"
        Response.Write "         <tr><td width='100' align='center'><strong>连接地址：</strong></td>"
        If Trim(Request("PageUrl")) = "" Then
            Response.Write "         <td><textarea name='LabelIntro' cols='96' rows='4' id='LabelIntro'>http://</textarea></td></tr>"
        Else
            Response.Write "         <td><textarea name='LabelIntro' cols='96' rows='4' id='LabelIntro'>" & Trim(Request("PageUrl")) & "</textarea></td></tr>"
        End If
        Response.Write "  </table></tr>" & vbCrLf
        Response.Write "  <script language=""JavaScript"">" & vbCrLf
        Response.Write "  <!--" & vbCrLf
        Response.Write "  function setFileFileds(num){" & vbCrLf
        Response.Write "      for(var i=1,str="""";i<=9;i++){" & vbCrLf
        Response.Write "          eval(""objFiles"" + i +"".style.display='none';"")" & vbCrLf
        Response.Write "      }" & vbCrLf
        Response.Write "      for(var i=1,str="""";i<=num;i++){" & vbCrLf
        Response.Write "          eval(""objFiles"" + i +"".style.display='';"")" & vbCrLf
        Response.Write "      }" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  //-->" & vbCrLf
        Response.Write "  </script>" & vbCrLf

        Response.Write "<tbody id='adv' style='display:none'><tr class=""tdbg""><td>" & vbCrLf
        Response.Write "  <table border='0' cellpadding='0' cellspacing='0' width='100%' >" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 编码格式：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td>GB2312：<INPUT TYPE='radio' NAME='Code' value=0 checked> UTF-8：<INPUT TYPE='radio' NAME='Code' value=1> Big5：<INPUT TYPE='radio' NAME='Code' value=2><font color=red> * </font>&nbsp;</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 开始字符：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td> <TEXTAREA NAME='LableStart' ROWS='' COLS='' style='width:400px;height:50px'></TEXTAREA><font color=red> * </font></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 结束字符：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td> <TEXTAREA NAME='LableEnd' ROWS='' COLS='' style='width:400px;height:50px'></TEXTAREA><font color=red> * </font></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 替换项目：&nbsp;</strong></td>"
        Response.Write "      <td>" & vbCrLf
        Response.Write "      <select name=""ReplaceNum"" onChange=""setFileFileds(this.value)"">" & vbCrLf
        Response.Write "         <option value=""0"">0</option>" & vbCrLf
        Response.Write "         <option value=""1"">1</option>" & vbCrLf
        Response.Write "         <option value=""2"">2</option>" & vbCrLf
        Response.Write "         <option value=""3"">3</option>" & vbCrLf
        Response.Write "         <option value=""4"">4</option>" & vbCrLf
        Response.Write "         <option value=""5"">5</option>" & vbCrLf
        Response.Write "         <option value=""6"">6</option>" & vbCrLf
        Response.Write "         <option value=""7"">7</option>" & vbCrLf
        Response.Write "         <option value=""8"">8</option>" & vbCrLf
        Response.Write "         <option value=""9"">9</option>" & vbCrLf
        Response.Write "      </select>" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'>&nbsp;</td>" & vbCrLf
        Response.Write "      <td>" & vbCrLf
        Response.Write "      <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>" & vbCrLf
        Dim i
        For i = 1 To 9
            Response.Write "  <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td class=""tdbg""  id=""objFiles" & i & """ valign='top' style=""display:'none'"">" & vbCrLf
            Response.Write i
            Response.Write "        将字符：<TEXTAREA NAME='ReplaceQuilt" & i & "' ROWS='' COLS='' style='width:250px;height:30px'></TEXTAREA>"
            Response.Write "        替换为：<TEXTAREA NAME='ReplaceWith" & i & "' ROWS='' COLS='' style='width:250px;height:30px'></TEXTAREA>"
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
        Next
        Response.Write "     </table>" & vbCrLf
        Response.Write "     </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 链接后缀：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td> <input name=""UpFileType"" type=""text"" id=""UpFileType"" size=""80"" maxlength=""50"" Value=""gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rmvb|html|asp|shtml|jsp|shtml|htm|php|cgi""> <font color=red> * </font> <font color='blue'>注：用|分割</font><br>" & vbCrLf
        Response.Write "      <font color='blue'>说明:将采集链接的相对地址转换为绝对地址,请在上面输入要转换链接的后缀。</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong>过滤选项：&nbsp;</strong></td>"
        Response.Write "    <td>"
        Response.Write "      &nbsp;&nbsp;<input name=""Script_Iframe"" type=""checkbox"" id=""Script_Iframe""  value=""1"">Iframe" & vbCrLf
        Response.Write "      <input name=""Script_Object"" type=""checkbox"" id=""Script_Object""  value=""1"">Object" & vbCrLf
        Response.Write "      <input name=""Script_Script"" type=""checkbox"" id=""Script_Script""  value=""1"">Script" & vbCrLf
        Response.Write "      <input name=""Script_Class"" type=""checkbox"" id=""Script_Class""  value=""1"">Style" & vbCrLf
        Response.Write "      <input name=""Script_Div"" type=""checkbox"" id=""Script_Div""  value=""1"">Div" & vbCrLf
        Response.Write "      <input name=""Script_Table"" type=""checkbox"" id=""Script_Table""  value=""1"">Table" & vbCrLf
        Response.Write "      <input name=""Script_Tr"" type=""checkbox"" id=""Script_tr""  value=""1"">Tr" & vbCrLf
        Response.Write "      <input name=""Script_td"" type=""checkbox"" id=""Script_td""  value=""1"">Td" & vbCrLf
        Response.Write "      <br>" & vbCrLf
        Response.Write "      &nbsp;&nbsp;<input name=""Script_Span"" type=""checkbox"" id=""Script_Span""  value=""1"">Span" & vbCrLf
        Response.Write "      &nbsp;&nbsp;<input name=""Script_Img"" type=""checkbox"" id=""Script_Img""  value=""1"">Img&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "      <input name=""Script_Font"" type=""checkbox"" id=""Script_Font""  value=""1"">FONT&nbsp;&nbsp;" & vbCrLf
        Response.Write "      <input name=""Script_A"" type=""checkbox"" id=""Script_A""  value=""1"">A&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "      <input name=""Script_Html"" type=""checkbox"" id=""Script_Html""  value=""1"">Html" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr></table></td></tr></tbody>" & vbCrLf
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>优 先 级：</strong></td>"
    Response.Write "          <td><input name='Priority' type='text' id='Priority' size='5' maxlength='5'></td>"
    Response.Write "          <td width='10'></td>"
    Response.Write "          <td><font color='#FF0000'>数字越小，优先级越高。当标签中再嵌套调用其他标签时，就需要决定标签的优先级。<br>系统按照如下顺序来替换标签：自定义标签-->系统通用标签-->频道标签</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If AddType = 0 Then
        Response.Write "    <tr class='title' height='22'>"
        Response.Write "      <td  align='center'><strong>标 签 内 容</strong></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "     <td >&nbsp;&nbsp;"
        Response.Write "        <textarea name='LabelContent' class='body2'   ROWS='10' COLS='108' onMouseUp=""setContent('get',1)"">请输入您自定义的html代码</textarea>"
        Response.Write "     </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "     <td >&nbsp;"
        Response.Write "        <textarea name='LabelContent2'  style='display:none' >请输入您自定义的html代码</textarea>"
        Response.Write "        <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=1&TemplateType=0&tContentid=LabelContent2' frameborder='1' scrolling='no' width='780' height='400' ></iframe>"
        Response.Write "     </td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'>"
    Response.Write "        <input name='LabelType' type='hidden' id='LabelType' value=" & AddType & ">"
    Response.Write "        <input name='Scode' type='hidden' id='Scode' value='" & CheckSecretCode("start") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 添 加 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub AddDyna()
    Dim LabelName1, lnum, AddType, dbname1, dbname2, i, rs, sqlstr, rstSchema
    LabelName1 = Trim(Request("LabelName1"))
    lnum = PE_CLng(Trim(Request("lnum")))
    AddType = PE_CLng(Trim(Request("addtype")))
    If AddType <> 3 Then AddType = 1
    dbname1 = Trim(Request("dbname1"))
    If dbname1 <> "" Then
        dbname1 = ReplaceBadChar(dbname1)
    End If
    dbname2 = ReplaceBadChar(Trim(Request("dbname2")))
    If dbname2 <> "" Then
        dbname2 = ReplaceBadChar(dbname2)
    End If

    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function addclass(){" & vbCrLf
    Response.Write "    var select=document.myform.LabelClassList;" & vbCrLf
    Response.Write "    for(i=0;i<select.length;i++){" & vbCrLf
    Response.Write "        if(document.myform.LabelClassList[i].selected==true){" & vbCrLf
    Response.Write "            document.myform.LabelClass.value=document.myform.LabelClassList[i].value;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.LabelName.value==''){" & vbCrLf
    Response.Write "     alert('标签名称不能为空！');" & vbCrLf
    Response.Write "     document.myform.LabelName.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.LabelIntro.value==''){" & vbCrLf
    Response.Write "     alert('查询语句不能为空！');" & vbCrLf
    Response.Write "     document.myform.LabelIntro.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Priority.value==''){" & vbCrLf
    Response.Write "     alert('优 先 等级不能为空！');" & vbCrLf
    Response.Write "     document.myform.Priority.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function changedb(){" & vbCrLf
    Response.Write "    var dbname=document.myform.dbname1.value;" & vbCrLf
    Response.Write "    var dbname2=document.myform.dbname2.value;" & vbCrLf
    Response.Write "    var Labelname=document.myform.LabelName.value;" & vbCrLf
    Response.Write "    var Listnum=document.myform.pagenum.value;" & vbCrLf
    Response.Write "    var addtype=document.myform.labeltype;" & vbCrLf
    Response.Write "    for(i=0;i<addtype.length;i++){" & vbCrLf
    Response.Write "        if(document.myform.labeltype[i].checked==true){" & vbCrLf
    Response.Write "            var addtype2=document.myform.labeltype[i].value" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    window.location.href=""Admin_Label.asp?Action=AddDyna&lnum="" + Listnum + ""&addtype="" + addtype2 + ""&dbname1="" + dbname + ""&dbname2="" + dbname2 + ""&LabelName1="" + Labelname + """";" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function addfield(){" & vbCrLf
    Response.Write "    document.myform.LabelIntro.value='';" & vbCrLf
    Response.Write "    var select=document.myform.field;" & vbCrLf
    Response.Write "    var select2=document.myform.field2;" & vbCrLf
    Response.Write "    for(i=0;i<select.length;i++){" & vbCrLf
    Response.Write "        if(document.myform.field[i].selected==true){" & vbCrLf
    Response.Write "            if(document.myform.dbname2.value==''){" & vbCrLf
    Response.Write "                if (document.myform.LabelIntro.value==''){" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=document.myform.field[i].value;" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=document.myform.LabelIntro.value+"",""+document.myform.field[i].value;" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                if (document.myform.LabelIntro.value==''){" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=document.myform.dbname1.value + ""."" + document.myform.field[i].value;" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=document.myform.LabelIntro.value + "","" + document.myform.dbname1.value + ""."" + document.myform.field[i].value;" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(document.myform.dbname2.value==''){" & vbCrLf
    Response.Write "        if(document.myform.pagenum.value>0){" & vbCrLf
    Response.Write "            document.myform.LabelIntro.value=""select "" + document.myform.LabelIntro.value + "" from " & dbname1 & """;" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            document.myform.LabelIntro.value=""select top 10 "" + document.myform.LabelIntro.value + "" from " & dbname1 & """;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        for(i=0;i<select2.length;i++){" & vbCrLf
    Response.Write "            if(document.myform.field2[i].selected==true){" & vbCrLf
    Response.Write "                if (document.myform.LabelIntro.value==''){" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=document.myform.dbname2.value + ""."" + document.myform.field2[i].value;" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=document.myform.LabelIntro.value + "","" + document.myform.dbname2.value + ""."" + document.myform.field2[i].value;" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if(document.myform.dbname1.value==''){" & vbCrLf
    Response.Write "            if(document.myform.pagenum.value>0){" & vbCrLf
    Response.Write "                document.myform.LabelIntro.value=""select "" + document.myform.LabelIntro.value + "" from " & dbname2 & """;" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                document.myform.LabelIntro.value=""select top 10 "" + document.myform.LabelIntro.value + "" from " & dbname2 & """;" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            if(document.myform.bg1.value==''){" & vbCrLf
    Response.Write "                if(document.myform.pagenum.value>0){" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=""select "" + document.myform.LabelIntro.value + "" from " & dbname1 & "," & dbname2 & """;" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=""select top 10 "" + document.myform.LabelIntro.value + "" from " & dbname1 & "," & dbname2 & """;" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                if(document.myform.pagenum.value>0){" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=""select "" + document.myform.LabelIntro.value + "" from " & dbname1 & "," & dbname2 & " where "";" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    document.myform.LabelIntro.value=""select top 10 "" + document.myform.LabelIntro.value + "" from " & dbname1 & "," & dbname2 & " where "";" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "                document.myform.LabelIntro.value=document.myform.LabelIntro.value + """ & dbname1 & "."" + document.myform.bg1.value + "" = "" + """ & dbname2 & "."" + document.myform.bg2.value;" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function checkfield(){" & vbCrLf
    Response.Write "    var strtmpp = ""<table border='1' cellpadding='2' cellspacing='1'  width='600' class='border'><tr align='center'>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist2(0)'>{$Now}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist2(1)'>{$NowDay}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist2(2)'>{$NowMonth}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist2(3)'>{$NowYear}</td></tr><tr align='center'>"";" & vbCrLf
    Response.Write "    var fieldtemp = document.myform.FieldList.value.split(""\n"");" & vbCrLf
    Response.Write "        for(i=0;i<fieldtemp.length;i++){" & vbCrLf
    Response.Write "            strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist("" + i + "")'>"" + fieldtemp[i] + ""</td>"";" & vbCrLf
    Response.Write "            if(((i+1)%6) == 0){" & vbCrLf
    Response.Write "                strtmpp = strtmpp + ""</tr><tr align='center'>"";" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        strtmpp = strtmpp + ""</table>"";" & vbCrLf
    Response.Write "        document.getElementById (""flist2"").innerHTML=strtmpp;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function addlist(input){" & vbCrLf
    Response.Write "    myform.LabelIntro.focus();" & vbCrLf
    Response.Write "    var str = document.selection.createRange();" & vbCrLf
    Response.Write "    if (input != null){" & vbCrLf
    Response.Write "        str.text = ""{input("" + input + "")}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function addlist2(input){" & vbCrLf
    Response.Write "    myform.LabelIntro.focus();" & vbCrLf
    Response.Write "    var str = document.selection.createRange();" & vbCrLf
    Response.Write "    if (input == 0){" & vbCrLf
    Response.Write "        str.text = ""{$Now}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 1){" & vbCrLf
    Response.Write "        str.text = ""{$NowDay}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 2){" & vbCrLf
    Response.Write "        str.text = ""{$NowMonth}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 3){" & vbCrLf
    Response.Write "        str.text = ""{$NowYear}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form action='Admin_Label.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td align='center'><strong>添 加 动 态 标 签（第一步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签名称：</strong></td>"
    Response.Write "          <td width='240'>{$MY_<input name='LabelName' type='text' id='LabelName' size='30' maxlength='50' value=" & LabelName1 & ">}</td>"
    Response.Write "          <td width='10'></td>"
    Response.Write "          <td><font color='#FF0000'>* 输入名称（英文要注意大小写）即可，不用输入定界符。</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签分类：</strong></td>"
    Response.Write "          <td colspan='3'><input name='LabelClass' type='text' id='LabelClass' size='30' maxlength='50'> " & getlabelclass(AddType) & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签类型：</strong></td>"
    Response.Write "          <td><Input type='radio' name='labeltype' value=1"
    If AddType = 1 Then Response.Write " checked"
    Response.Write " onClick=""flist.style.display='none';"">标准动态标签 <Input type='radio' name='labeltype' value=3"
    If AddType = 3 Then Response.Write " checked"
    Response.Write " onClick=""flist.style.display='';"">函数型动态标签</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>分页数量：</strong></td>"
    Response.Write "          <td width='45'><input name='pagenum' type='text' id='pagenum' size='3' maxlength='10' value=" & lnum & "></td>"
    Response.Write "          <td width='10'></td><td><font color='#FF0000'>* 动态标签分页显示的每页显示数,为0时则不分页。</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>自动刷新：</strong></td>"
    Response.Write "          <td width='45'><input name='rtime' type='text' id='rtime' size='3' maxlength='3' value=0></td>"
    Response.Write "          <td width='10'></td><td><font color='#FF0000'>* 标签内容自动刷新频率，必须大于10秒并启用分页功能才能启动。</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "      <td width='100' align='center'><strong>主表：</strong></td>"
    Response.Write "      <td><select name='dbname1' style='width:250px;' onChange=""changedb()"">"
    Response.Write "<option value=''>请选择一个表</option>"
    Set rstSchema = Conn.OpenSchema(20)
    Do Until rstSchema.EOF
    If Left(rstSchema("TABLE_NAME"), 2) <> "MS" And rstSchema("TABLE_NAME") <> "" And rstSchema("TABLE_NAME") <> "PE_Admin" And rstSchema("TABLE_NAME") <> "PE_Config" Then
        Response.Write "<option value='" & rstSchema("TABLE_NAME") & "'"
        If dbname1 = rstSchema("TABLE_NAME") Then
        Response.Write " selected"
        End If
        Response.Write ">" & rstSchema("TABLE_NAME") & "</option>"
    End If
    rstSchema.MoveNext
    Loop
    Response.Write "</select></td>"
    Response.Write "      <td width='100' align='center'><strong>从表：</strong></td>"
    Response.Write "      <td><select name='dbname2' style='width:250px;' onChange=""changedb()"">"
    Response.Write "<option value=''>请选择一个表</option>"
    Set rstSchema = Conn.OpenSchema(20)
    Do Until rstSchema.EOF
    If Left(rstSchema("TABLE_NAME"), 2) <> "MS" And rstSchema("TABLE_NAME") <> "" And rstSchema("TABLE_NAME") <> "PE_Admin" And rstSchema("TABLE_NAME") <> "PE_Config" Then
        Response.Write "<option value='" & rstSchema("TABLE_NAME") & "'"
        If dbname2 = rstSchema("TABLE_NAME") Then
        Response.Write " selected"
        End If
        Response.Write ">" & rstSchema("TABLE_NAME") & "</option>"
    End If
    rstSchema.MoveNext
    Loop
    Response.Write "</select></td></tr>"
    If dbname1 <> "" And dbname2 <> "" Then
        Response.Write "        <tr><td align='center'><strong>约束字段：</strong></td>"
        Response.Write "          <td><select name='bg1' style='width:250px;'>"
        Response.Write "<option value=''>选择主表字段</option>"
        If dbname1 <> "" Then
        sqlstr = "select * from " & dbname1
        Set rs = Conn.Execute(sqlstr)
        For i = 0 To rs.Fields.Count - 1
            Response.Write "<option value='" & rs(i).name & "'>" & rs(i).name & "</option>"
        Next
        Else
        Response.Write "<option value='0'>请先选择一个表</option>"
        End If
        Response.Write "</select></td><td align='center'><strong><< 等于 >></strong></td><td><select name='bg2' style='width:250px;'>"
        Response.Write "<option value=''>选择从表字段</option>"
        If dbname2 <> "" Then
        sqlstr = "select * from " & dbname2
        Set rs = Conn.Execute(sqlstr)
        For i = 0 To rs.Fields.Count - 1
            Response.Write "<option value='" & rs(i).name & "'>" & rs(i).name & "</option>"
        Next
        Else
        Response.Write "<option value='0'>请先选择一个表</option>"
        End If
        Response.Write "         </select></td><td><font color='#FF0000'>请选择跨表查询的约束条件。</font></td>"
        Response.Write "</tr>"
    End If
    Response.Write "        <tr><td width='100' align='center'><strong>选择字段：</strong><br><br><font color='#FF0000'>请选择需要调用的字段名称,按Ctrl或Shift键多选</font></td>"
    Response.Write "         <td width='100'><select name='field' size='1' multiple style='height:200px;width:250px;' onChange='addfield()'>"
    If dbname1 <> "" Then
    sqlstr = "select * from " & dbname1
    Set rs = Conn.Execute(sqlstr)
    For i = 0 To rs.Fields.Count - 1
        Response.Write "<option value='" & rs(i).name & "'>" & rs(i).name & "</option>"
    Next
    Else
    Response.Write "<option value='0'>请先选择一个表</option>"
    End If
    Response.Write "</select></td>"
    Response.Write "<td align='center'><strong>>>>></strong></td>"
    Response.Write "<td><select name='field2' size='2' multiple style='height:200px;width:250px;' onChange='addfield()'>"
    If dbname2 <> "" Then
    sqlstr = "select * from " & dbname2
    Set rs = Conn.Execute(sqlstr)
    For i = 0 To rs.Fields.Count - 1
        Response.Write "<option value='" & rs(i).name & "'>" & rs(i).name & "</option>"
    Next
    Else
    Response.Write "<option value='0'>请先选择一个表</option>"
    End If
    Response.Write "</select></td>"
    Response.Write "         <td></td></tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tbody id='flist'"
    If AddType <> 3 Then Response.Write " style=""display:none"""
    Response.Write "><tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>参数说明：</strong></td>"
    Response.Write "          <td width='80'><textarea name='FieldList' cols='40' rows='5' id='FieldList' onkeydown=""checkfield();""></textarea></td>"
    Response.Write "          <td width='10'></td><td><font color='#FF0000'>* 输入函数列表参数,每行一个。</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr></tbody>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "         <td width='100' align='center'><strong>查询语句：</strong></td>"
    If lnum > 0 Then
        Response.Write "         <td><div id=""flist2"">"
        Response.Write "<table border='1' cellpadding='2' cellspacing='1'  width='600' class='border'><tr align='center'>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(0)'>{$Now}</td>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(1)'>{$NowDay}</td>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(2)'>{$NowMonth}</td>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(3)'>{$NowYear}</td></tr></table>"
        Response.Write "</div><textarea name='LabelIntro' cols='83' rows='6' id='LabelIntro'>select * from " & dbname1 & "</textarea></td>"
    Else
        Response.Write "         <td><div id=""flist2"">"
        Response.Write "<table border='1' cellpadding='2' cellspacing='1'  width='600' class='border'><tr align='center'>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(0)'>{$Now}</td>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(1)'>{$NowDay}</td>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(2)'>{$NowMonth}</td>"
        Response.Write "<td style='cursor:hand;' onclick='addlist2(3)'>{$NowYear}</td></tr></table>"
        Response.Write "</div><textarea name='LabelIntro' cols='83' rows='6' id='LabelIntro'>select top 10 * from " & dbname1 & "</textarea></td>"
    End If
    Response.Write "       </tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>优 先 级：</strong></td>"
    Response.Write "          <td width='50'><input name='Priority' type='text' id='Priority' size='5' maxlength='5'></td>"
    Response.Write "          <td width='10'></td>"
    Response.Write "          <td><font color='#FF0000'>数字越小，优先级越高。当标签中再嵌套调用其他标签时，就需要决定标签的优先级。<br>系统按照如下顺序来替换标签：自定义标签-->系统通用标签-->频道标签</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='AddDyna2'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub AddDyna2()
    Dim LabelName, LabelClass, LabelType, PageNum, RTime, FieldList, LabelIntro, LabelIntro2, Priority, strtmp, dbname1, dbname2, dbtype
    Dim i, rs
    LabelName = "MY_" & Trim(Request.Form("LabelName"))
    LabelClass = Trim(Request.Form("LabelClass"))
    LabelType = PE_CLng(Trim(Request.Form("labeltype")))
    PageNum = PE_CLng(Trim(Request.Form("pagenum")))
    RTime = PE_CLng(Trim(Request.Form("rtime")))
    FieldList = Request.Form("FieldList")
    dbname1 = Trim(Request.Form("dbname1"))
    dbname2 = Trim(Request.Form("dbname2"))
    LabelIntro = Trim(Request.Form("LabelIntro"))
    If Left(LCase(LabelIntro), 6) <> "select" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>只能做选择查询！</li>"
        Exit Sub
    Else
        If InStr(LCase(LabelIntro), "where") > 0 Then
            regEx.Pattern = "(.*?)\where"
            Set Matches = regEx.Execute(LabelIntro)
            For Each Match In Matches			
                LabelIntro2 = Match.value
                Exit For 
            Next
			
            LabelIntro2 = Trim(Replace(LCase(LabelIntro2), "where", ""))
            regEx.Pattern = "\{input\((.*?)\)\}"
            Set Matches = regEx.Execute(LabelIntro2)
            LabelIntro2 = regEx.Replace(LabelIntro2, "1")
        Else
            regEx.Pattern = "\{input\((.*?)\)\}"
            Set Matches = regEx.Execute(LabelIntro)
            LabelIntro2 = regEx.Replace(LabelIntro, "1")
        End If
    End If
    Priority = Trim(Request.Form("Priority"))

    If dbname1 = "PE_Article" Or dbname2 = "PE_Article" Or InStr(LabelIntro, "PE_Article") > 0 Then
        dbtype = 1
    ElseIf dbname1 = "PE_Soft" Or dbname2 = "PE_Soft" Or InStr(LabelIntro, "PE_Soft") > 0 Then
        dbtype = 2
    ElseIf dbname1 = "PE_Photo" Or dbname2 = "PE_Photo" Or InStr(LabelIntro, "PE_Photo") > 0 Then
        dbtype = 3
    ElseIf dbname1 = "PE_Product" Or dbname2 = "PE_Product" Or InStr(LabelIntro, "PE_Product") > 0 Then
        dbtype = 5
    Else
        dbtype = 0
    End If

    strtmp = "<table border='1' cellpadding='2' cellspacing='1'  width='695' class='border'><tr align='center'>" & vbCrLf
    On Error Resume Next
    Set rs = Conn.Execute(LabelIntro2)

    If Err.Number <> 0 Then
        Set rs = Nothing
        FoundErr = True
        ErrMsg = ErrMsg & "<li>SQL查询失败，查询代码：" & LabelIntro2 & "错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    Else
        For i = 0 To rs.Fields.Count - 1
            strtmp = strtmp & "<td onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""cursor:hand;"" onclick=""addfield('" & rs(i).name & "'," & i & "," & dbtype & "," & rs(i).Type & ")"">" & rs(i).name & "</td>" & vbCrLf
            If (i + 1) Mod 5 = 0 Then
                strtmp = strtmp & "</tr><tr align='center'>"
            End If
        Next
        Set rs = Nothing
    End If
    strtmp = strtmp & "</tr></table>"

    Call ShowJSLabel(1)
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function addlist3(input){" & vbCrLf
    Response.Write "    myform.LabelContent.focus();" & vbCrLf
    Response.Write "    var str = document.selection.createRange();" & vbCrLf
    Response.Write "    if (input == 0){" & vbCrLf
    Response.Write "        str.text = ""{$Now}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 1){" & vbCrLf
    Response.Write "        str.text = ""{$NowDay}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 2){" & vbCrLf
    Response.Write "        str.text = ""{$NowMonth}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 3){" & vbCrLf
    Response.Write "        str.text = ""{$NowYear}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 4){" & vbCrLf
    Response.Write "        str.text = ""{$AutoID}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (input == 5){" & vbCrLf
    Response.Write "        str.text = ""{$totalPut}"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "        document.myform.LabelContent2.value=document.myform.LabelContent.value;" & vbCrLf
    Response.Write "        editor.HtmlEdit.document.body.innerHTML=document.myform.LabelContent2.value;" & vbCrLf
    Response.Write "}" & vbCrLf
     Response.Write "</script>" & vbCrLf
    Response.Write "<form action='Admin_Label.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    If LabelType = 3 Then
        Response.Write "      <td align='center'><strong>添 加 函 数 型 动 态 标 签（第二步）</strong></td>"
    Else
        Response.Write "      <td align='center'><strong>添 加 动 态 标 签（第二步）</strong></td>"
    End If
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签名称：</strong></td>"
    Response.Write "          <td>{$" & LabelName & "}</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>查询语句：</strong></td>"
    Response.Write "          <td><textarea name='LabelIntro' cols='96' rows='6' id='LabelIntro' readonly>" & LabelIntro & "</textarea></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>字段列表：</strong></td>" & vbCrLf
    Response.Write "          <td>" & strtmp & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"

    strtmp = "<table border='1' cellpadding='2' cellspacing='1'  width='695' class='border'><tr align='center'>" & vbCrLf
    strtmp = strtmp & "<td style='cursor:hand;' onclick='addlist3(0)'>{$Now}</td>"
    strtmp = strtmp & "<td style='cursor:hand;' onclick='addlist3(1)'>{$NowDay}</td>"
    strtmp = strtmp & "<td style='cursor:hand;' onclick='addlist3(2)'>{$NowMonth}</td>"
    strtmp = strtmp & "<td style='cursor:hand;' onclick='addlist3(3)'>{$NowYear}</td>"
    strtmp = strtmp & "<td style='cursor:hand;' onclick='addlist3(4)'>{$AutoID}</td>"
    strtmp = strtmp & "<td style='cursor:hand;' onclick='addlist3(5)'>{$totalPut}</td></tr>"
    If LabelType = 3 And FieldList <> "" Then
        strtmp = strtmp & "<tr align='center'>"
        Dim arrFieldList, FieldList2
        arrFieldList = Split(FieldList, vbCrLf)
        For i = 0 To UBound(arrFieldList)
            If Trim(arrFieldList(i)) <> "" Then
                strtmp = strtmp & "<td onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""cursor:hand;"" onclick=""addfield2(" & i & ")"">" & arrFieldList(i) & "</td>" & vbCrLf
                If (i + 1) Mod 4 = 0 Then
                    strtmp = strtmp & "</tr><tr align='center'>"
                End If
            End If
            FieldList2 = FieldList2 & arrFieldList(i) & "|||"
        Next
        strtmp = strtmp & "</tr>"
    End If
    strtmp = strtmp & "</table>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>参数列表：</strong></td>" & vbCrLf
    Response.Write "          <td>" & strtmp & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='title' height='22'>"
    Response.Write "      <td  align='center'><strong>标 签 内 容</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "     <td >&nbsp;&nbsp;"
    Response.Write "        <textarea name='LabelContent' class='body2' ROWS='10' COLS='108' onMouseUp=""setContent('get',1)"">{Loop}{Infobegin}循环内容{Infoend}{/Loop}</textarea>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "     <td >&nbsp;"
    Response.Write "        <textarea name='LabelContent2'  style='display:none' >{Loop}{Infobegin}循环内容{Infoend}{/Loop}</textarea>"
    Response.Write "        <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=1&TemplateType=0&tContentid=LabelContent2' frameborder='1' scrolling='no' width='780' height='400' ></iframe>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'>"
    Response.Write "        <input name='LabelName' type='hidden' id='LabelName' value='" & LabelName & "'>"
    Response.Write "        <input name='LabelClass' type='hidden' id='LabelClass' value='" & LabelClass & "'>"
    Response.Write "        <input name='Priority' type='hidden' id='Priority' value=" & Priority & ">"
    Response.Write "        <input name='LabelType' type='hidden' id='LabelType' value=" & LabelType & ">"
    Response.Write "        <input name='pagenum' type='hidden' id='pagenum' value=" & PageNum & ">"
    Response.Write "        <input name='rtime' type='hidden' id='rtime' value=" & RTime & ">"
    Response.Write "        <input name='FieldList' type='hidden' id='FieldList' value=" & FieldList2 & ">"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 添 加 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim LabelID, sqlLabel, rsLabel, LabelIntro2, EditLabelContent, LabelContent, strTemp, LabelNameTemp
    LabelID = PE_CLng(Trim(Request("LabelID")))
    If LabelID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    sqlLabel = "select * from PE_Label where LabelID=" & LabelID
    Set rsLabel = Conn.Execute(sqlLabel)
    If rsLabel.BOF And rsLabel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的标签！</li>"
        rsLabel.Close
        Set rsLabel = Nothing
        Exit Sub
    End If
        
    '解决文本框重复问题
    LabelContent = rsLabel("LabelContent")
    regEx.Pattern = "(\<\/textarea\>)"
    LabelContent = regEx.Replace(LabelContent, "[/textarea]")
    
    EditLabelContent = rsLabel("LabelContent")
    EditLabelContent = Replace(EditLabelContent, "<!--{$", "{$")
    EditLabelContent = Replace(EditLabelContent, "}-->", "}")
     
    '图片替换JS
    regEx.Pattern = "(\<Script)([\s\S]*?)(\<\/Script\>)"
    Set Matches = regEx.Execute(EditLabelContent)
    For Each Match In Matches
        strTemp = Replace(Match.value, "<", "[!")
        strTemp = Replace(strTemp, ">", "!]")
        strTemp = Replace(strTemp, "'", """")
        strTemp = "<IMG alt='#" & strTemp & "#' src=""" & InstallDir & "editor/images/jscript.gif"" border=0 $>"
        EditLabelContent = Replace(EditLabelContent, Match.value, strTemp)
    Next
        
    '图片替换超级标签
    regEx.Pattern = "(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct)\((.*?)\)\}"
    EditLabelContent = regEx.Replace(EditLabelContent, "<IMG src=""" & InstallDir & "editor/images/label.gif"" border=0 zzz='$1($2)}'>")

    If rsLabel("LabelType") = 1 Or rsLabel("LabelType") = 3 Then
        LabelIntro2 = rsLabel("LabelIntro")
        If InStr(LCase(LabelIntro2), "where") > 0 Then
            regEx.Pattern = "(.*?)\where"
            Set Matches = regEx.Execute(LabelIntro2)
            For Each Match In Matches
                LabelIntro2 = Match.value
				Exit for
            Next
            LabelIntro2 = Trim(Replace(LCase(LabelIntro2), "where", ""))
            regEx.Pattern = "\{input\((.*?)\)\}"
            Set Matches = regEx.Execute(LabelIntro2)
            LabelIntro2 = regEx.Replace(LabelIntro2, "1")
        Else
            regEx.Pattern = "\{input\((.*?)\)\}"
            Set Matches = regEx.Execute(LabelIntro2)
            LabelIntro2 = regEx.Replace(LabelIntro2, "1")
        End If

        Dim i, rs, dbtype
        If InStr(rsLabel("LabelIntro"), "PE_Article") > 0 Then
            dbtype = 1
        ElseIf InStr(rsLabel("LabelIntro"), "PE_Soft") > 0 Then
            dbtype = 2
        ElseIf InStr(rsLabel("LabelIntro"), "PE_Photo") > 0 Then
            dbtype = 3
        ElseIf InStr(rsLabel("LabelIntro"), "PE_Product") > 0 Then
            dbtype = 5
        Else
            dbtype = 0
        End If

        strTemp = "<table border='1' cellpadding='2' cellspacing='1'  width='695' class='border'><tr align='center'>" & vbCrLf
        On Error Resume Next
        Set rs = Conn.Execute(LabelIntro2)
        If Err.Number <> 0 Then
            Set rs = Nothing
            FoundErr = True
            ErrMsg = ErrMsg & "<li>SQL查询失败，错误原因：" & Err.Description
            Err.Clear
            Exit Sub
        Else
            Err.Clear
            For i = 0 To rs.Fields.Count - 1
                strTemp = strTemp & "<td onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""cursor:hand;"" onclick=""addfield('" & rs(i).name & "'," & i & "," & dbtype & "," & rs(i).Type & ")"">" & rs(i).name & "</td>" & vbCrLf
                If (i + 1) Mod 5 = 0 Then
                    strTemp = strTemp & "</tr><tr align='center'>"
                End If
            Next
            Set rs = Nothing
        End If
        strTemp = strTemp & "</tr></table>"
    End If

    Call ShowJSLabel(rsLabel("LabelType"))
    Response.Write "<script language = 'JavaScript'>" & vbCrLf	
    Response.Write "function checkfield(){" & vbCrLf
    Response.Write "    var strtmpp = ""<Div id ='flist2'><table border='1' cellpadding='2' cellspacing='1'  width='695' class='border'><tr align='center'>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist3(0)'>{$Now}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist3(1)'>{$NowDay}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist3(2)'>{$NowMonth}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist3(3)'>{$NowYear}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist3(4)'>{$AutoID}</td>"";" & vbCrLf
    Response.Write "    strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addlist3(5)'>{$totalPut}</td>"";" & vbCrLf
    Response.Write "    var fieldtemp = document.myform.FieldList.value.split(""\n"");" & vbCrLf
    Response.Write "                strtmpp = strtmpp + ""</tr><tr align='center'>"";" & vbCrLf
    Response.Write "        for(i=0;i<fieldtemp.length;i++){" & vbCrLf
    Response.Write "            strtmpp = strtmpp + ""<td style='cursor:hand;' onclick='addfield2("" + i + "")'>"" + fieldtemp[i] + ""</td>"";" & vbCrLf    
    Response.Write "            if(((i+1)%6) == 0){" & vbCrLf
    Response.Write "                strtmpp = strtmpp + ""</tr><tr align='center'>"";" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        strtmpp = strtmpp + ""</table><div>"";" & vbCrLf
    Response.Write "        document.getElementById (""flist2"").innerHTML=strtmpp;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form action='Admin_Label.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title' height='22'>"
    If rsLabel("LabelType") = 0 Then
        Response.Write "      <td  align='center'><strong>修 改 静 态 标 签</strong></td>"
    ElseIf rsLabel("LabelType") = 1 Then
        Response.Write "      <td  align='center'><strong>修 改 动 态 标 签</strong></td>"
    ElseIf rsLabel("LabelType") = 2 Then
        Response.Write "      <td  align='center'><strong>修 改 采 集 标 签 <font color=#aaffaa>（本标签类型较为消耗ＣＰＵ时间，推荐独立服务器用户使用）</font></strong></td>"
    ElseIf rsLabel("LabelType") = 3 Then
        Response.Write "      <td  align='center'><strong>修 改 动 态 函 数 标 签</strong></td>"
    End If
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签名称：</strong></td>"
    If Left(rsLabel("LabelName"), 3) = "MY_" Then
        LabelNameTemp = Right(rsLabel("LabelName"), Len(rsLabel("LabelName")) - 3)
    Else
        LabelNameTemp = rsLabel("LabelName")
    End If
    Response.Write "          <td>{$MY_<input name='LabelName' type='text' id='LabelName' size='30' maxlength='50' value='" & LabelNameTemp & "'>}"
    Response.Write "          <td width='10'></td>"
    Response.Write "          <td><font color='#FF0000'>* 输入名称（英文要注意大小写）即可，不用输入定界符。</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>标签分类：</strong></td>"
    Response.Write "          <td colspan='3'><input name='LabelClass' type='text' id='LabelClass' size='30' maxlength='50' value='" & rsLabel("LabelClass") & "'> " & getlabelclass(rsLabel("LabelType")) & "</td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If rsLabel("LabelType") = 1 Or rsLabel("LabelType") = 3 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
        Response.Write "        <tr>"
        Response.Write "          <td width='100' align='center'><strong>分页数量：</strong></td>"
        Response.Write "          <td width='45'><input name='pagenum' type='text' id='pagenum' size='3' maxlength='5' value=" & rsLabel("PageNum") & "></td>"
        Response.Write "          <td width='10'></td><td><font color='#FF0000'>* 动态标签分页显示的每页显示数,为0时则不分页。</font></td>"
        Response.Write "        </tr>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
        Response.Write "        <tr>"
        Response.Write "          <td width='100' align='center'><strong>自动刷新：</strong></td>"
        Response.Write "          <td width='45'><input name='rtime' type='text' id='rtime' size='3' maxlength='3' value=" & rsLabel("reFlashTime") & "></td>"
        Response.Write "          <td width='10'></td><td><font color='#FF0000'>* 标签内容自动刷新频率，必须大于10秒并启用分页功能才能启动。</font></td>"
        Response.Write "        </tr>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    If rsLabel("LabelType") = 2 Then
        Response.Write "        <tr>"
        Response.Write "          <td width='100' align='center'><strong>采集方式：</strong></td>"
        Response.Write "          <td><INPUT TYPE='radio' NAME='CaiType' value='0'"
        If rsLabel("AreaCollectionID") = 0 Then Response.Write " checked"
        Response.Write " onClick=""adv.style.display='none'"">快速 <INPUT TYPE='radio' NAME='CaiType' value='1'"
        If rsLabel("AreaCollectionID") = 1 Then Response.Write " checked"
        Response.Write " onClick=""adv.style.display=''"">高级</td>" & vbCrLf
        Response.Write "        </tr>"
    End If
    Response.Write "        <tr>"
    If rsLabel("LabelType") = 0 Then
        Response.Write "         <td width='100' align='center'><strong>标签简介：</strong></td>"
    ElseIf rsLabel("LabelType") = 1 Or rsLabel("LabelType") = 3 Then
        Response.Write "         <td width='100' align='center'><strong>查询语句：</strong></td>"
    ElseIf rsLabel("LabelType") = 2 Then
        Response.Write "         <td width='100' align='center'><strong>连接地址：</strong></td>"
    End If
    Response.Write "         <td><textarea name='LabelIntro' cols='96' rows='4' id='LabelIntro'>" & rsLabel("LabelIntro") & "</textarea></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
	
    If rsLabel("LabelType") = 3  Then
        Response.Write "<tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
        Response.Write "        <tr>"
        Response.Write "          <td width='100' align='center'><strong>参数说明：</strong></td>"
        Response.Write "          <td width='80'><textarea name='FieldList' cols='40' rows='5' id='FieldList'onkeydown=""checkfield();"">"
        Dim arrFieldList
        arrFieldList = Split(rsLabel("fieldlist"), "|||")	
	For i = 0 To UBound(arrFieldList)
            If Trim(arrFieldList(i)) <> "" and i<>UBound(arrFieldList) Then
                Response.Write arrFieldList(i) & vbCrLf
            else
                Response.Write arrFieldList(i)
            End If
        Next
	Response.Write "</textarea></td>"
        Response.Write "          <td width='10'></td><td><font color='#FF0000'>* 输入函数列表参数,每行一个。</font></td>"
        Response.Write "        </tr>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr></tbody>"
    End If

    If rsLabel("LabelType") = 1 Or rsLabel("LabelType") = 3 Then
        Response.Write "<script language = 'JavaScript'>" & vbCrLf
        Response.Write "function addlist3(input){" & vbCrLf
        Response.Write "    myform.LabelContent.focus();" & vbCrLf
        Response.Write "    var str = document.selection.createRange();" & vbCrLf
        Response.Write "    if (input == 0){" & vbCrLf
        Response.Write "        str.text = ""{$Now}"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (input == 1){" & vbCrLf
        Response.Write "        str.text = ""{$NowDay}"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (input == 2){" & vbCrLf
        Response.Write "        str.text = ""{$NowMonth}"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (input == 3){" & vbCrLf
        Response.Write "        str.text = ""{$NowYear}"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (input == 4){" & vbCrLf
        Response.Write "        str.text = ""{$AutoID}"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (input == 5){" & vbCrLf
        Response.Write "        str.text = ""{$totalPut}"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "        document.myform.LabelContent2.value=document.myform.LabelContent.value;" & vbCrLf
        Response.Write "        editor.HtmlEdit.document.body.innerHTML=document.myform.LabelContent2.value;" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "</script>" & vbCrLf
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
        Response.Write "        <tr>"
        Response.Write "          <td width='100' align='center'><strong>字段列表：</strong></td>" & vbCrLf
        Response.Write "          <td>" & strTemp & "</td>"
        Response.Write "        </tr>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        strTemp = "<div id='flist2'><table border='1'id= cellpadding='2' cellspacing='1'  width='695' class='border'><tr align='center'>" & vbCrLf
        strTemp = strTemp & "<td style='cursor:hand;' onclick='addlist3(0)'>{$Now}</td>"
        strTemp = strTemp & "<td style='cursor:hand;' onclick='addlist3(1)'>{$NowDay}</td>"
        strTemp = strTemp & "<td style='cursor:hand;' onclick='addlist3(2)'>{$NowMonth}</td>"
        strTemp = strTemp & "<td style='cursor:hand;' onclick='addlist3(3)'>{$NowYear}</td>"
        strTemp = strTemp & "<td style='cursor:hand;' onclick='addlist3(4)'>{$AutoID}</td>"
        strTemp = strTemp & "<td style='cursor:hand;' onclick='addlist3(5)'>{$totalPut}</td></tr>"
        If rsLabel("LabelType") = 3 And rsLabel("fieldlist") <> "" Then
            strTemp = strTemp & "<tr align='center'>"
            For i = 0 To UBound(arrFieldList)
               If Trim(arrFieldList(i)) <> "" Then
                    strTemp = strTemp & "<td onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""cursor:hand;"" onclick=""addfield2(" & i & ")"">" & arrFieldList(i) & "</td>" & vbCrLf
                    If (i + 1) Mod 6 = 0 Then
                        strTemp = strTemp & "</tr><tr align='center'>"
                    End If
                End If
            Next
            strTemp = strTemp & "</tr>"
        End If
        strTemp = strTemp & "</table><div>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
        Response.Write "        <tr>"
        Response.Write "          <td width='100' align='center'><strong>参数列表：</strong></td>" & vbCrLf
        Response.Write "          <td>" & strTemp & "</td>"
        Response.Write "        </tr>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
    ElseIf rsLabel("LabelType") = 2 Then
        Dim rsArea, Code, StringReplace, LableStart, LableEnd, UpFileType, FilterProperty
        Dim Script_Property, ReplaceNum

        If rsLabel("AreaCollectionID") > 0 Then
            Set rsArea = Conn.Execute("select Top 1 * from PE_AreaCollection where AreaID=" & rsLabel("AreaCollectionID") & " and Type=1")
            If Not rsArea.EOF Then
                Code = rsArea("Code")
                StringReplace = rsArea("StringReplace")
                LableStart = rsArea("LableStart")
                LableEnd = rsArea("LableEnd")
                FilterProperty = rsArea("FilterProperty")
                UpFileType = rsArea("UpFileType")
            End If
        Else
            Code = 0
            FilterProperty = "0|0|0|0|0|0|0|0|0|0|0|0|0"
            UpFileType = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rmvb|html|asp|shtml|jsp|shtml|htm|php|cgi"
        End If

        Response.Write "  <script language=""JavaScript"">" & vbCrLf
        Response.Write "  <!--" & vbCrLf
        Response.Write "  function setFileFileds(num){" & vbCrLf
        Response.Write "      for(var i=1,str="""";i<=9;i++){" & vbCrLf
        Response.Write "          eval(""objFiles"" + i +"".style.display='none';"")" & vbCrLf
        Response.Write "      }" & vbCrLf
        Response.Write "      for(var i=1,str="""";i<=num;i++){" & vbCrLf
        Response.Write "          eval(""objFiles"" + i +"".style.display='';"")" & vbCrLf
        Response.Write "      }" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  //-->" & vbCrLf
        Response.Write "  </script>" & vbCrLf
        If rsLabel("AreaCollectionID") > 0 Then
            Response.Write "<tbody id='adv' style='display:'>" & vbCrLf
        Else
            Response.Write "<tbody id='adv' style='display:none'>" & vbCrLf
        End If
        Response.Write "<tr class=""tdbg""><td>" & vbCrLf
        Response.Write "  <table border='0' cellpadding='0' cellspacing='0' width='100%' >" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 编码格式：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td>GB2312：<INPUT TYPE='radio' NAME='Code' value=0"
        If Code = 0 Then Response.Write " checked"
        Response.Write "> UTF-8：<INPUT TYPE='radio' NAME='Code' value=1"
        If Code = 1 Then Response.Write " checked"
        Response.Write "> Big5：<INPUT TYPE='radio' NAME='Code' value=2"
        If Code = 2 Then Response.Write " checked"
        Response.Write "><font color=red> * </font>&nbsp;</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 开始字符：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td> <TEXTAREA NAME='LableStart' ROWS='' COLS='' style='width:400px;height:50px'>" & LableStart & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 结束字符：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td> <TEXTAREA NAME='LableEnd' ROWS='' COLS='' style='width:400px;height:50px'>" & LableEnd & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf

        Dim arrAreaCode2, arrAreaCode, AreaCode1, AreaCode2

        arrAreaCode2 = Split(StringReplace, "$$$")
        ReplaceNum = UBound(arrAreaCode2) + 1
        Response.Write "  <tr> " & vbCrLf
        Response.Write "    <td width='100' align='center'><strong> 代码预览：&nbsp;</strong></td>" & vbCrLf
        Response.Write "    <td> <TEXTAREA NAME='preview' ROWS='' COLS='' style='width:614px;height:150px'>" & Server.HTMLEncode(GetBody(GetHttpPage(rsLabel("LabelIntro"), PE_CLng(Code)), LableStart, LableEnd, True, True)) & "</TEXTAREA><font color=red> * </font></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf

        Response.Write "    <tr><td width='100' align='center'><strong> 替换项目：&nbsp;</strong></td>"
        Response.Write "      <td>" & vbCrLf
        Response.Write "      <select name=""ReplaceNum"" onChange=""setFileFileds(this.value)"">" & vbCrLf
        Response.Write "         <option value=""0"" " & IsOptionSelected(ReplaceNum, 0) & ">0</option>" & vbCrLf
        Response.Write "         <option value=""1"" " & IsOptionSelected(ReplaceNum, 1) & ">1</option>" & vbCrLf
        Response.Write "         <option value=""2"" " & IsOptionSelected(ReplaceNum, 2) & ">2</option>" & vbCrLf
        Response.Write "         <option value=""3"" " & IsOptionSelected(ReplaceNum, 3) & ">3</option>" & vbCrLf
        Response.Write "         <option value=""4"" " & IsOptionSelected(ReplaceNum, 4) & ">4</option>" & vbCrLf
        Response.Write "         <option value=""5"" " & IsOptionSelected(ReplaceNum, 5) & ">5</option>" & vbCrLf
        Response.Write "         <option value=""6"" " & IsOptionSelected(ReplaceNum, 6) & ">6</option>" & vbCrLf
        Response.Write "         <option value=""7"" " & IsOptionSelected(ReplaceNum, 7) & ">7</option>" & vbCrLf
        Response.Write "         <option value=""8"" " & IsOptionSelected(ReplaceNum, 8) & ">8</option>" & vbCrLf
        Response.Write "         <option value=""9"" " & IsOptionSelected(ReplaceNum, 9) & ">9</option>" & vbCrLf
        Response.Write "      </select>" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'>&nbsp;</td>" & vbCrLf
        Response.Write "      <td>" & vbCrLf
        Response.Write "      <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>" & vbCrLf
        For i = 0 To UBound(arrAreaCode2)
            arrAreaCode = Split(arrAreaCode2(i), "|||")
            AreaCode1 = arrAreaCode(0)
            AreaCode2 = arrAreaCode(1)

            Response.Write "  <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td class=""tdbg""  id=""objFiles" & i + 1 & """ valign='top' style=""display:''"">" & vbCrLf
            Response.Write i + 1
            Response.Write "        将字符：<TEXTAREA NAME='ReplaceQuilt" & i + 1 & "' ROWS='' COLS='' style='width:250px;height:30px'>" & AreaCode1 & "</TEXTAREA>"
            Response.Write "        替换为：<TEXTAREA NAME='ReplaceWith" & i + 1 & "' ROWS='' COLS='' style='width:250px;height:30px'>" & AreaCode2 & "</TEXTAREA>"
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
        Next
        ReplaceNum = ReplaceNum + 1
        For i = ReplaceNum To 9
            Response.Write "  <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "    <td class=""tdbg""  id=""objFiles" & i & """ valign='top' style=""display:'none'"">" & vbCrLf
            Response.Write i
            Response.Write "        将字符：<TEXTAREA NAME='ReplaceQuilt" & i & "' ROWS='' COLS='' style='width:250px;height:30px'></TEXTAREA>"
            Response.Write "        替换为：<TEXTAREA NAME='ReplaceWith" & i & "' ROWS='' COLS='' style='width:250px;height:30px'></TEXTAREA>"
            Response.Write "    </td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
        Next
        Response.Write "     </table>" & vbCrLf
        Response.Write "     </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr><td width='100' align='center'><strong> 链接后缀：&nbsp;</strong></td>" & vbCrLf
        Response.Write "      <td> <input name=""UpFileType"" type=""text"" id=""UpFileType"" size=""80"" maxlength=""50"" value=" & UpFileType & "> <font color=red> * </font> <font color='blue'>注：用|分割</font><br>" & vbCrLf
        Response.Write "      <font color='blue'>说明:将采集链接的相对地址转换为绝对地址,请在上面输入要转换链接的后缀。</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Script_Property = Split(FilterProperty, "|")
        Response.Write "    <tr><td width='100' align='center'><strong>过滤选项：&nbsp;</strong></td>"
        Response.Write "    <td>"
        Response.Write "      &nbsp;&nbsp;<input name=""Script_Iframe"" type=""checkbox"" id=""Script_Iframe""  value=""1"" "
        If Script_Property(0) = "1" Then Response.Write " checked"
        Response.Write ">Iframe" & vbCrLf
        Response.Write "      <input name=""Script_Object"" type=""checkbox"" id=""Script_Object""  value=""1"" "
        If Script_Property(1) = "1" Then Response.Write " checked"
        Response.Write ">Object" & vbCrLf
        Response.Write "      <input name=""Script_Script"" type=""checkbox"" id=""Script_Script""  value=""1"" "
        If Script_Property(2) = "1" Then Response.Write " checked"
        Response.Write ">Script" & vbCrLf
        Response.Write "      <input name=""Script_Class"" type=""checkbox"" id=""Script_Class""  value=""1"" "
        If Script_Property(3) = "1" Then Response.Write " checked"
        Response.Write ">Style" & vbCrLf
        Response.Write "      <input name=""Script_Div"" type=""checkbox"" id=""Script_Div""  value=""1"" "
        If Script_Property(4) = "1" Then Response.Write " checked"
        Response.Write ">Div" & vbCrLf
        Response.Write "      <input name=""Script_Table"" type=""checkbox"" id=""Script_Table""  value=""1"" "
        If Script_Property(5) = "1" Then Response.Write " checked"
        Response.Write ">Table" & vbCrLf
        Response.Write "      <input name=""Script_Tr"" type=""checkbox"" id=""Script_tr""  value=""1"" "
        If Script_Property(6) = "1" Then Response.Write " checked"
        Response.Write ">Tr" & vbCrLf
        Response.Write "      <input name=""Script_td"" type=""checkbox"" id=""Script_td""  value=""1"" "
        If Script_Property(7) = "1" Then Response.Write " checked"
        Response.Write ">Td" & vbCrLf
        Response.Write "      <br>" & vbCrLf
        Response.Write "      &nbsp;&nbsp;<input name=""Script_Span"" type=""checkbox"" id=""Script_Span""  value=""1"" "
        If Script_Property(8) = "1" Then Response.Write " checked"
        Response.Write ">Span" & vbCrLf
        Response.Write "      &nbsp;&nbsp;<input name=""Script_Img"" type=""checkbox"" id=""Script_Img""  value=""1"" "
        If Script_Property(9) = "1" Then Response.Write " checked"
        Response.Write ">Img&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "      <input name=""Script_Font"" type=""checkbox"" id=""Script_Font""  value=""1"" "
        If Script_Property(10) = "1" Then Response.Write " checked"
        Response.Write ">FONT&nbsp;&nbsp;" & vbCrLf
        Response.Write "      <input name=""Script_A"" type=""checkbox"" id=""Script_A""  value=""1"" "
        If Script_Property(11) = "1" Then Response.Write " checked"
        Response.Write ">A&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "      <input name=""Script_Html"" type=""checkbox"" id=""Script_Html""  value=""1"" "
        If Script_Property(12) = "1" Then Response.Write " checked"
        Response.Write ">Html" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr></table></td></tr></tbody>" & vbCrLf
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' >"
    Response.Write "        <tr>"
    Response.Write "          <td width='100' align='center'><strong>优 先 级：</strong></td>"
    Response.Write "          <td><input name='Priority' type='text' id='Priority' size='5' maxlength='5' value='" & rsLabel("Priority") & "'>"
    Response.Write "          <td width='10'></td>"
    Response.Write "          <td><font color='#FF0000'>数字越小，优先级越高。当标签中再嵌套调用其他标签时，就需要决定标签的优先级。<br>系统按照如下顺序来替换标签：自定义标签-->系统通用标签-->频道标签</font></td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If rsLabel("LabelType") <> 2 Then
        Response.Write "    <tr class='title' height='22'>"
        Response.Write "      <td  align='center'><strong>标 签 内 容</strong></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "     <td >&nbsp;&nbsp;"
        Response.Write "        <textarea name='LabelContent' class='body2' ROWS='10' COLS='108' onMouseUp=""setContent('get',1)"">" & LabelContent & "</textarea>"
        Response.Write "     </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "     <td >&nbsp;"
        Response.Write "        <textarea name='LabelContent2'  style='display:none' >" & Server.HTMLEncode(EditLabelContent) & "</textarea>"
        Response.Write "        <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=1&TemplateType=0&tContentid=LabelContent2' frameborder='1' scrolling='no' width='780' height='400' ></iframe>"
        Response.Write "     </td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40'  align='center'><input name='LabelID' type='hidden' id='LabelID' value='" & LabelID & "'>"
    Response.Write "        <input name='LabelType' type='hidden' id='LabelType' value=" & rsLabel("LabelType") & ">"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 保存修改结果 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsLabel.Close
    Set rsLabel = Nothing
End Sub

Sub Save()
    Dim LabelID, LabelName, LabelClass, LabelIntro, LabelIntro2, Priority, LabelContent, LabelType, PageNum, RTime, SystemLabelName
    Dim rsLabel, sqlLabel, trs, i, AreaCollectionID, FieldList, Scode
    LabelID = PE_CLng(Trim(Request.Form("LabelID")))
    LabelName = Trim(Request.Form("LabelName"))
    LabelClass = Trim(Request.Form("LabelClass"))
    LabelIntro = Trim(Request.Form("LabelIntro"))
    Priority = Trim(Request.Form("Priority"))
    LabelContent = Trim(Request.Form("LabelContent"))
    LabelType = PE_CLng(Trim(Request.Form("LabelType")))
    PageNum = PE_CLng(Trim(Request.Form("pagenum")))
    RTime = PE_CLng(Trim(Request.Form("rtime")))
    Scode = Trim(Request.Form("Scode"))
	FieldList = Request.Form("FieldList")
    If Action = "SaveModify" Then
        If LabelID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定LabelID</li>"
            Exit Sub
        End If
    End If

    If LabelName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>标签名称不能为空！</li>"
    Else
        LabelName = ReplaceBadChar(LabelName)
        If Left(LabelName, 3) <> "MY_" Then
        LabelName = "MY_" & LabelName
        End If
        If Action = "SaveModify" Then
            Set trs = Conn.Execute("select * from PE_Label where LabelID<>" & LabelID & " and LabelName='" & LabelName & "'")
        Else
            Set trs = Conn.Execute("select * from PE_Label where LabelName='" & LabelName & "'")
        End If
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>指定的标签名称已经存在！</li>"
        End If
        Set trs = Nothing
    End If
    
    If LabelType = "" Then
        LabelType = 0
    Else
        LabelType = PE_CLng(LabelType)
    End If

    If LabelType = 1 Or LabelType = 3 Then
        If LabelIntro = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>查询语句不能为空！</li>"
        Else
            If Left(LCase(LabelIntro), 6) <> "select" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>只能使用查询语句！</li>"
            End If
            If InStr(LCase(LabelIntro), "pe_admin") > 0 Or InStr(LCase(LabelIntro), "pe_config") > 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>出于安全目的本功能禁止对管理员及系统设置的查询！</li>"
            End If
        End If
 
        If InStr(LCase(LabelIntro), "where") > 0 Then
            regEx.Pattern = "(.*?)\where"
            Set Matches = regEx.Execute(LabelIntro)
            For Each Match In Matches
                 LabelIntro2 = Match.value
                 Exit For
            Next
            LabelIntro2 = Trim(Replace(LCase(LabelIntro2), "where", ""))
            regEx.Pattern = "\{input\((.*?)\)\}"
            Set Matches = regEx.Execute(LabelIntro2)
            LabelIntro2 = regEx.Replace(LabelIntro2, "1")
        Else
            regEx.Pattern = "\{input\((.*?)\)\}"
            Set Matches = regEx.Execute(LabelIntro)
            LabelIntro2 = regEx.Replace(LabelIntro, "1")
        End If



        On Error Resume Next
        Set trs = Conn.Execute(LabelIntro2)
        If Err.Number <> 0 Then
            Set trs = Nothing
            FoundErr = True
            ErrMsg = ErrMsg & "<li>SQL查询失败，错误原因：" & Err.Description
            Err.Clear
            Exit Sub
        End If
        Set trs = Nothing
    ElseIf LabelType = 2 Then
        If LabelIntro = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>连接地址不能为空！</li>"
        Else
            If GetHttpPage(LabelIntro, 0) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>在获取:" & LabelIntro & "网页源码时发生错误。</li>"
            End If
        End If
    End If
    
    If LabelContent = "" And LabelType <> 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>标签内容不能为空！</li>"
    End If
    
    If Priority = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>优先级不能为空！</li>"
    Else
        Priority = PE_CLng(Priority)
        If Priority <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>优先级必须大于0！</li>"
        End If
    End If
    
    If InStr(LabelContent, LabelName&"(") > 0 or InStr(LabelContent, LabelName&"}") > 0 Then '自定义标签{$MY_标签名}的标签内容可以包含{$MY_标签名**}这样的标签.
        FoundErr = True
        ErrMsg = ErrMsg & "<li>自定义标签不能自己包括自己！</li>"
    End If
    
    If InStr(LabelContent, "<body>") > 0 Or InStr(LabelContent, "<html>") > 0 Or InStr(LabelContent, "</html>") > 0 Or InStr(LabelContent, "</body>") > 0 Or InStr(LabelContent, "<head>") > 0 Or InStr(LabelContent, "</head>") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>自定义标签不能包含&lt;html&gt;&nbsp;,&lt;body&gt;,&lt;/body&gt;,&lt;/html&gt;等！！！</li>"
    End If

    Dim NullBody, strTemp, strTemp2, Match2
   '使用正则 分别过滤调编辑模板中的图片
     
    regEx.Pattern = "(\<body)(.[^\<]*)(\>)"
    Set Matches = regEx.Execute(LabelContent)
    For Each Match In Matches
        NullBody = Match.value
    Next

    
    If NullBody <> "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>自定义标签不能包含&lt;html&gt;&nbsp;,&lt;body&gt;,&lt;/body&gt;,&lt;/html&gt;等！！！</li>"
    End If
    
    If FoundErr = True Then Exit Sub

    Dim rscai, Code, StringReplace, LableStart, LableEnd, UpFileType, FilterProperty
    Dim Script_Iframe, Script_Object, Script_Script, Script_Class, Script_Div, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, Script_Table, Script_Tr, Script_Td
    Dim ReplaceNum, AreaCode

    If Action = "SaveModify" Then
        If LabelType = 2 And Trim(Request.Form("CaiType")) = "1" Then
            Code = PE_CLng(Request.Form("Code"))
            StringReplace = Trim(Request.Form("StringReplace"))
            LableStart = Trim(Request.Form("LableStart"))
            LableEnd = Trim(Request.Form("LableEnd"))
            UpFileType = Trim(Request.Form("UpFileType"))

            Script_Iframe = Trim(Request.Form("Script_Iframe"))
            Script_Object = Trim(Request.Form("Script_Object"))
            Script_Script = Trim(Request.Form("Script_Script"))
            Script_Class = Trim(Request.Form("Script_Class"))
            Script_Div = Trim(Request.Form("Script_Div"))
            Script_Span = Trim(Request.Form("Script_Span"))
            Script_Img = Trim(Request.Form("Script_Img"))
            Script_Font = Trim(Request.Form("Script_Font"))
            Script_A = Trim(Request.Form("Script_A"))
            Script_Html = Trim(Request.Form("Script_Html"))
            Script_Table = Trim(Request.Form("Script_Table"))
            Script_Tr = Trim(Request.Form("Script_Tr"))
            Script_Td = Trim(Request.Form("Script_Td"))

            FilterProperty = Script_Iframe & "|" & Script_Object & "|" & Script_Script & "|" & Script_Class & "|" & Script_Div & "|" & Script_Table & "|" & Script_Tr & "|" & Script_Td & "|" & Script_Span & "|" & Script_Img & "|" & Script_Font & "|" & Script_A & "|" & Script_Html
            ReplaceNum = PE_CLng(Trim(Request.Form("ReplaceNum")))

            If Code = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>区域项目采集编码不能为空</li>"
            End If
            If LableStart = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>截取代码开始不能为空</li>"
            End If
            If LableEnd = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>截取代码结束不能为空</li>"
            End If
            If UpFileType = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>截取内容链接的后缀名不能为空</li>"
            End If
            If FoundErr = True Then
                Exit Sub
            End If
            If FoundErr <> True Then
                AreaCode = GetHttpPage(LabelIntro, PE_CLng(Code)) '获得列表源代码
                If AreaCode <> "" Then
                    AreaCode = GetBody(AreaCode, LableStart, LableEnd, True, True) '获得列表代码
                    If AreaCode = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>在截取区域代码的时发生错误。</li>"
                    End If
                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>在获取:" & LabelIntro & "网页源码时发生错误。</li>"
                End If
            End If

            If ReplaceNum <> 0 Then
                For i = 1 To ReplaceNum
                    If i <> 1 Then
                        StringReplace = StringReplace & "$$$"
                    End If
                    StringReplace = StringReplace & Trim(Request("ReplaceQuilt" & i)) & "|||" & Trim(Request("ReplaceWith" & i))
                Next
            End If
            sqlLabel = "select * from PE_Label where LabelID=" & LabelID
            Set rsLabel = Server.CreateObject("ADODB.Recordset")
                rsLabel.Open sqlLabel, Conn, 1, 3
                If rsLabel("AreaCollectionID") > 0 Then
                    sqlLabel = "SELECT TOP 1 * FROM PE_AreaCollection Where AreaID=" & rsLabel("AreaCollectionID")
                    Set rscai = Server.CreateObject("adodb.recordset")
                    rscai.Open sqlLabel, Conn, 1, 3
                Else
                    sqlLabel = "SELECT TOP 1 * FROM PE_AreaCollection"
                    Set rscai = Server.CreateObject("adodb.recordset")
                    rscai.Open sqlLabel, Conn, 1, 3
                    rscai.addnew
                End If
                rscai("AreaName") = LabelName
                rscai("AreaFile") = LabelName
                rscai("AreaIntro") = LabelName
                rscai("Code") = Code
                rscai("StringReplace") = StringReplace
                rscai("AreaUrl") = LabelIntro
                rscai("LableStart") = LableStart
                rscai("LableEnd") = LableEnd
                rscai("FilterProperty") = FilterProperty
                rscai("UpFileType") = UpFileType
                rscai("AreaPassed") = True
                rscai("Type") = 1
                rscai.Update
                rscai.Close
                Set rscai = Nothing

                rsLabel("LabelName") = LabelName
                rsLabel("LabelClass") = LabelClass
                rsLabel("LabelIntro") = LabelIntro
                rsLabel("LabelContent") = LabelContent
                rsLabel("Priority") = Priority
                If rsLabel("AreaCollectionID") = 0 Then
                    Set rscai = Conn.Execute("select max(AreaID) from PE_AreaCollection")
                    rsLabel("AreaCollectionID") = rscai(0)
                End If
                rsLabel.Update
            rsLabel.Close
        Else
            sqlLabel = "select * from PE_Label where LabelID=" & LabelID
            Set rsLabel = Server.CreateObject("ADODB.Recordset")
                rsLabel.Open sqlLabel, Conn, 1, 3
                rsLabel("LabelName") = LabelName
                rsLabel("LabelClass") = LabelClass
                rsLabel("LabelIntro") = LabelIntro
                rsLabel("PageNum") = PageNum
                rsLabel("reFlashTime") = RTime
                rsLabel("LabelContent") = LabelContent
                rsLabel("Priority") = Priority
                If LabelType = 3 Then
                    FieldList = Request.Form("FieldList")
                    Dim arrFieldList, FieldList2
                    arrFieldList = Split(FieldList, vbCrLf)
                    For i = 0 To UBound(arrFieldList)
                        If Trim(arrFieldList(i)) <> "" Then
                            FieldList2 = FieldList2 & arrFieldList(i) & "|||"
                        End If
                    Next
                    rsLabel("fieldlist") = FieldList2
                End If
                rsLabel.Update
            rsLabel.Close
            If LabelType = 2 And Trim(Request.Form("CaiType")) = "0" And AreaCollectionID > 0 Then Conn.Execute ("delete from PE_AreaCollection where AreaID=" & AreaCollectionID)
        End If
        Set rsLabel = Nothing
        Call WriteSuccessMsg("修改自定义标签成功！", ComeUrl & "")
    Else
        If LabelType > 0 Then
            Set rscai = Conn.Execute("Select count(LabelID) From PE_Label Where LabelType=" & LabelType)
            If rscai(0) > 30 And SystemDatabaseType = "ACCESS" Then
                Set rscai = Nothing
                FoundErr = True
                ErrMsg = ErrMsg & "<li>您添加的本类型标签已经超过服务器负载能力，请删除不常用的标签再添加！</li>"
                Exit Sub
            End If
        End If
        If LabelType = 2 And Trim(Request.Form("CaiType")) = "1" Then '增加采集标签入库
            Code = PE_CLng(Request.Form("Code"))
            StringReplace = Trim(Request.Form("StringReplace"))
            LableStart = Trim(Request.Form("LableStart"))
            LableEnd = Trim(Request.Form("LableEnd"))
            UpFileType = Trim(Request.Form("UpFileType"))

            Script_Iframe = Trim(Request.Form("Script_Iframe"))
            Script_Object = Trim(Request.Form("Script_Object"))
            Script_Script = Trim(Request.Form("Script_Script"))
            Script_Class = Trim(Request.Form("Script_Class"))
            Script_Div = Trim(Request.Form("Script_Div"))
            Script_Span = Trim(Request.Form("Script_Span"))
            Script_Img = Trim(Request.Form("Script_Img"))
            Script_Font = Trim(Request.Form("Script_Font"))
            Script_A = Trim(Request.Form("Script_A"))
            Script_Html = Trim(Request.Form("Script_Html"))
            Script_Table = Trim(Request.Form("Script_Table"))
            Script_Tr = Trim(Request.Form("Script_Tr"))
            Script_Td = Trim(Request.Form("Script_Td"))

            FilterProperty = Script_Iframe & "|" & Script_Object & "|" & Script_Script & "|" & Script_Class & "|" & Script_Div & "|" & Script_Table & "|" & Script_Tr & "|" & Script_Td & "|" & Script_Span & "|" & Script_Img & "|" & Script_Font & "|" & Script_A & "|" & Script_Html
            ReplaceNum = PE_CLng(Trim(Request.Form("ReplaceNum")))

            If Code = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>区域项目采集编码不能为空</li>"
            End If
            If LableStart = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>截取代码开始不能为空</li>"
            End If
            If LableEnd = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>截取代码结束不能为空</li>"
            End If
            If UpFileType = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>截取内容链接的后缀名不能为空</li>"
            End If
            If FoundErr = True Then
                Exit Sub
            End If

            If FoundErr <> True Then
                AreaCode = GetHttpPage(LabelIntro, PE_CLng(Code)) '获得列表源代码
                If AreaCode <> "" Then
                    AreaCode = GetBody(AreaCode, LableStart, LableEnd, True, True) '获得列表代码
                    If AreaCode = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>在截取区域代码的时发生错误。</li>"
                    End If
                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>在获取:" & LabelIntro & "网页源码时发生错误。</li>"
                End If
            End If

            If ReplaceNum <> 0 Then
                For i = 1 To ReplaceNum
                    If i <> 1 Then
                        StringReplace = StringReplace & "$$$"
                    End If
                    StringReplace = StringReplace & Trim(Request("ReplaceQuilt" & i)) & "|||" & Trim(Request("ReplaceWith" & i))
                Next
            End If
            sqlLabel = "SELECT TOP 1 * FROM PE_AreaCollection"
            Set rscai = Server.CreateObject("adodb.recordset")
            rscai.Open sqlLabel, Conn, 1, 3
            rscai.addnew
            rscai("AreaName") = LabelName
            rscai("AreaFile") = LabelName
            rscai("AreaIntro") = LabelName
            rscai("Code") = Code
            rscai("StringReplace") = StringReplace
            rscai("AreaUrl") = LabelIntro
            rscai("LableStart") = LableStart
            rscai("LableEnd") = LableEnd
            rscai("FilterProperty") = FilterProperty
            rscai("UpFileType") = UpFileType
            rscai("AreaPassed") = True
            rscai("Type") = 1
            rscai.Update
            rscai.Close
            Set rscai = Conn.Execute("select max(AreaID) from PE_AreaCollection")
            AreaCollectionID = rscai(0)
            Set rscai = Nothing
        Else
            AreaCollectionID = 0
        End If
        sqlLabel = "select top 1 * from PE_Label"
        Set rsLabel = Server.CreateObject("ADODB.Recordset")
        rsLabel.Open sqlLabel, Conn, 1, 3
        rsLabel.addnew
        rsLabel("LabelName") = LabelName
        rsLabel("LabelClass") = LabelClass
        rsLabel("LabelIntro") = LabelIntro
        rsLabel("PageNum") = PageNum
        rsLabel("reFlashTime") = RTime
        rsLabel("fieldlist") = FieldList	
        rsLabel("LabelContent") = LabelContent
        rsLabel("Priority") = Priority
        rsLabel("LabelType") = LabelType
        rsLabel("AreaCollectionID") = AreaCollectionID
        rsLabel.Update
        rsLabel.Close
        Set rsLabel = Nothing
        Call WriteSuccessMsg("保存自定义标签成功！", ComeUrl & "")
    End If
End Sub

Sub DelLabel()
    Dim LabelID, sqlLabel, rsLabel, tLabelContent, ListType
    LabelID = PE_CLng(Trim(Request("LabelID")))
    ListType = PE_CLng(Trim(Request("ListType")))
    If LabelID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数丢失！</li>"
        Exit Sub
    End If
    sqlLabel = "select * from PE_Label where LabelID=" & LabelID
    Set rsLabel = Server.CreateObject("ADODB.Recordset")
    rsLabel.Open sqlLabel, Conn, 1, 3
    If rsLabel.BOF And rsLabel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的标签！</li>"
        rsLabel.Close
        Set rsLabel = Nothing
        Exit Sub
    End If
    rsLabel.Delete
    rsLabel.Update
    rsLabel.Close
    Set rsLabel = Nothing
    Call CloseConn
    If ListType > 0 Then
        Response.Redirect "Admin_Label.asp?ListType=" & ListType
    Else
        Response.Redirect "Admin_Label.asp"
    End If
End Sub

'=================================================
'过程名：Import
'作  用：导入标签第一步
'=================================================
Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Label.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>标签导入（第一步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的标签数据库的文件名： "
    Response.Write "        <input name='LabelMdb' type='text' id='LabelMdb' value='../Temp/PE_Label.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='import2'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：import2
'作  用：导入标签第二步
'=================================================
Sub import2()
    On Error Resume Next

    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    
    '获得导入模板数据库路径
    mdbname = Replace(Trim(Request.Form("LabelMdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("LabelMdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "＄", "/") '防止外部链接安全问题

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入标签数据库名"
        Exit Sub
    End If

    '建立导入标签数据库
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If

    Response.Write "<form name='myform' method='post' action='Admin_Label.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>标签导入（第二步）</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>将要导入的标签分类</strong></td>"
    Response.Write "          </tr>"
    Response.Write "           <tr>"
    Response.Write "            <td>"
    
    '显示标签
    Response.Write "              <select name='LabelID' size='2' multiple style='height:300px;width:250px;'>"
    
    sql = "select LabelID,LabelName,LabelClass,LabelType from PE_Label order by LabelType,LabelClass,LabelID desc"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "                <option value='0'>没有任何标签</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Select Case rs("LabelType")
            Case 1
               Response.Write "            <option value='" & rs("LabelID") & "'>动态标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            Case 2
               Response.Write "            <option value='" & rs("LabelID") & "'>采集标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            Case 3
               Response.Write "            <option value='" & rs("LabelID") & "'>函数标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            Case Else
               Response.Write "            <option value='" & rs("LabelID") & "'>静态标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            End Select
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"
    Response.Write "                  </td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='10'></td></tr>"
    Response.Write "                  <tr>"
    Response.Write "                    <td height='25' align='center'><b> 提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='20'></td></tr>"
    Response.Write "                  <tr><td colspan='3' height='25' align='center'><input type='submit' name='Submit' value=' 导入标签 ' onClick=""document.myform.Action.value='Doimport';"""
    Response.Write "                 </td></tr>"
    Response.Write "               </table>"
    Response.Write "               <input name='LabelMdb' type='hidden' id='LabelMdb' value='" & mdbname & "'>"
    Response.Write "               <input name='Action' type='hidden' id='Action' value='Doimport'>"
    Response.Write "               <br>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "       </table>"
    Response.Write "</form>"
End Sub

'=================================================
'过程名：DoImport
'作  用：导入标签保存
'=================================================
Sub DoImport()
    On Error Resume Next
    
    Dim crs, mdbname, tconn
    Dim LabelID, rs, sql, rsLabel, Table_PE_lable
    LabelID = Trim(Request.Form("LabelID"))
    mdbname = Replace(Trim(Request.Form("LabelMdb")), "'", "")
    If IsValidID(LabelID) = False Then
        LabelID = ""
    End If
    
    If LabelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>您尚未选择导入标签！</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导出标签数据库名"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
        
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If

    Set crs = tconn.Execute("select * from PE_Label where LabelID in (" & LabelID & ") order by LabelID")
    Set rs = Server.CreateObject("adodb.recordset")
    Do While Not crs.EOF
        rs.Open "select * from PE_Label", Conn, 1, 3
        rs.addnew
        rs("LabelName") = crs("LabelName")
        rs("LabelClass") = crs("LabelClass")
        rs("LabelType") = crs("LabelType")
        rs("PageNum") = crs("PageNum")
        rs("reFlashTime") = crs("reFlashTime")
        rs("fieldlist") = crs("fieldlist")
        rs("LabelIntro") = crs("LabelIntro")
        rs("Priority") = crs("Priority")
        rs("LabelContent") = crs("LabelContent")
        rs("AreaCollectionID") = crs("AreaCollectionID")
        rs.Update
        rs.Close
        crs.MoveNext
    Loop
    Set rs = Nothing
    crs.Close
    Set crs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("已经成功从指定的数据库中导入选中的标签！", ComeUrl & "?Action=Import2&LabelMdb=" & Replace(mdbname, "/", "＄") & "")
End Sub

'=================================================
'过程名：Export
'作  用：导出标签
'=================================================
Sub Export()
    Dim rs, sql
    Dim trs, iCount
 
    Response.Write "<form name='myform' method='post' action='Admin_Label.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>标签导出</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='LabelID' size='2' multiple style='height:300px;width:450px;'>"
    
    sql = "select LabelID,LabelName,LabelClass,LabelType from PE_Label Order by LabelType,LabelClass,LabelID desc"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>没有任何标签</option>"
        '关闭提交按钮
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Select Case rs("LabelType")
            Case 1
               Response.Write "            <option value='" & rs("LabelID") & "'>动态标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            Case 2
               Response.Write "            <option value='" & rs("LabelID") & "'>采集标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            Case 3
               Response.Write "            <option value='" & rs("LabelID") & "'>函数标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            Case Else
               Response.Write "            <option value='" & rs("LabelID") & "'>静态标签[" & rs("LabelClass") & "] -- {$" & rs("LabelName") & "}</option>"
            End Select
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
    Response.Write "        <td colspan='2'>目标数据库：<input name='LabelMdb' type='text' id='LabelMdb' value='../Temp/PE_Label.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> 先清空目标数据库</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='执行导出操作' onClick=""document.myform.Action.value='Doexport';"">"
    Response.Write "              <input name='Action' type='hidden' id='Action' value='Doexport'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.LabelID.length;i++){" & vbCrLf
    Response.Write "    document.myform.LabelID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.LabelID.length;i++){" & vbCrLf
    Response.Write "    document.myform.LabelID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'过程名：DoExport
'作  用：导出标签
'=================================================
Sub DoExport()
    On Error Resume Next
    Dim mdbname, tconn, trs
    Dim LabelID, rs, FormatConn

    LabelID = Trim(Request.Form("LabelID"))
    FormatConn = Request.Form("FormatConn")
    mdbname = Replace(Trim(Request.Form("LabelMdb")), "'", "")
    If IsValidID(LabelID) = False Then
        LabelID = ""
    End If

    If LabelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要导出的标签</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导出标签数据库名</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If

    Err.Clear
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
        Err.Clear
        Exit Sub
    End If

    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_Label")
    End If

    Set rs = Conn.Execute("select * from PE_Label where LabelID in (" & LabelID & ")  order by LabelID")
    Set trs = Server.CreateObject("adodb.recordset")
    Do While Not rs.EOF
        trs.Open "select * from PE_Label", tconn, 1, 3
        trs.addnew
        trs("LabelName") = rs("LabelName")
        trs("LabelClass") = rs("LabelClass")
        trs("LabelType") = rs("LabelType")
        trs("PageNum") = rs("PageNum")
        trs("reFlashTime") = rs("reFlashTime")
        trs("fieldlist") = rs("fieldlist")
        trs("LabelIntro") = rs("LabelIntro")
        trs("Priority") = rs("Priority")
        trs("LabelContent") = rs("LabelContent")
        trs("AreaCollectionID") = rs("AreaCollectionID")
        trs.Update
        trs.Close
        rs.MoveNext
    Loop
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("已经成功将所选中的自定义标签设置导出到指定的数据库中！", ComeUrl)
End Sub

Function IsOptionSelected(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function

Function getlabelclass(itype)
    If itype = "" Then
        getlabelclass = ""
    Else
        Dim strtmp, rsClass
        strtmp = "<select name='LabelClassList' onChange='addclass()'><option value=''>新增分类</option>"
        Set rsClass = Conn.Execute("select LabelClass from PE_Label Where LabelType=" & itype & " GROUP BY LabelClass")
        Do While Not rsClass.EOF
            If Trim(rsClass(0) & "") <> "" Then
                strtmp = strtmp & "<option value='" & rsClass(0) & "'>" & rsClass(0) & "</option>"
            End If
            rsClass.MoveNext
        Loop
        Set rsClass = Nothing
        strtmp = strtmp & "</select>"
        getlabelclass = strtmp
    End If
End Function



%>
