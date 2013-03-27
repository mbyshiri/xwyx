<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

FileName = "Admin_City.asp"
strFileName = FileName & "?Field=" & strField & "&keyword=" & Keyword

Response.Write "<html><head><title>网站邮政编码管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("邮 政 编 码 管 理", 10032)
Response.Write "<tr class='tdbg'><td width='70'>管理导航：</td><td>"
Response.Write " <a href='Admin_City.asp'>全部邮政编码</a> | "
Response.Write " <a href='Admin_City.asp?Action=PostCodeAdd'>添加邮政编码</a>"
Response.Write "</table>" & vbCrLf

Select Case Action
Case "PostCodeAdd"
    Call PostCodeAdd
Case "SavePostCodeAdd"
    Call SavePostCodeAdd
Case "PostCodeDel"
    Call PostCodeDel
Case "PostCodeEdit"
    Call PostCodeEdit
Case "SavePostCodeEdit"
    Call SavePostCodeEdit
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsPostCode, sqlPostCode
    Dim strAction, strSql
    If Keyword <> "" Then
        strAction = "搜索结果："
        Select Case strField
        Case "Province"
            strSql = " and Province like '%" & Keyword & "%' "
            strAction = strAction & "省份中含有关键字<font color='red'>" & Keyword & "</font>的记录"
        Case "City"
            strSql = " and City like '%" & Keyword & "%' "
            strAction = strAction & "城市中含有关键字<font color='red'>" & Keyword & "</font>的记录"
        Case "Area"
            strSql = " and Area like '%" & Keyword & "%' "
            strAction = strAction & "县区中含有关键字<font color='red'>" & Keyword & "</font>的记录"
        Case "PostCode"
            strSql = " and PostCode like '%" & Keyword & "%' "
            strAction = strAction & "邮政编码中含有关键字<font color='red'>" & Keyword & "</font>的记录"
        Case "AreaCode"
            strSql = " and AreaCode like '%" & Keyword & "%' "
            strAction = strAction & "区号中含有关键字<font color='red'>" & Keyword & "</font>的记录"
        Case Else
            strSql = " and Area like '%" & Keyword & "%' "
            strAction = strAction & "地区中含有关键字<font color='red'>" & Keyword & "</font>的记录"
        End Select
    Else
        strAction = "全部邮政编码"
    End If
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：邮政编码管理&nbsp;&gt;&gt;&nbsp;"
    Response.Write strAction
    Response.Write "</td></tr></table>"
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "      <td>"
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' align='center' height='22'>"
    Response.Write "          <td><strong>所属省</strong></td>"
    Response.Write "          <td><strong>所属城市</strong></td>"
    Response.Write "          <td width='150'><strong>地区</strong></td>"
    Response.Write "          <td width='140'><strong>邮政编码</strong></td>"
    Response.Write "          <td width='110'><strong>区号</strong></td>"
    Response.Write "          <td width='150'><strong>操作</strong></td>"
    Response.Write "        </tr>"

    sqlPostCode = "select * from PE_City where 1=1 "

    If Keyword <> "" Then
        sqlPostCode = sqlPostCode & strSql
    End If
    sqlPostCode = sqlPostCode & " order by AreaID asc"
    Set rsPostCode = Server.CreateObject("adodb.recordset")
    rsPostCode.Open sqlPostCode, Conn, 1, 1
    If rsPostCode.EOF Then
        Response.Write "<tr class='tdbg' align='center' height='50'><td colspan='10'>无此信息！</td></tr>"
    Else
        totalPut = rsPostCode.RecordCount
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
                rsPostCode.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Dim PostCodeNum
        PostCodeNum = 0
        Do While Not rsPostCode.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td>" & rsPostCode("Province") & "</td>"
            Response.Write "        <td>" & rsPostCode("City") & "</td>"
            Response.Write "        <td width='150'>&nbsp;&nbsp;&nbsp;" & rsPostCode("Area") & "</td>"
            Response.Write "        <td width='140' align='center'>" & rsPostCode("PostCode") & "</td>"
            Response.Write "        <td width='110' align='center'>" & rsPostCode("AreaCode") & "</td>"
            Response.Write "        <td width='150' align='center'><a href='Admin_City.asp?Action=PostCodeEdit&AreaID=" & rsPostCode("AreaID") & "'>编辑</a> | <a href='Admin_City.asp?Action=PostCodeDel&AreaID=" & rsPostCode("AreaID") & "' onclick=""return confirm('确定要删除此条记录吗？');"">删除</a></td>"
            Response.Write "      </tr>"
            PostCodeNum = PostCodeNum + 1
            If PostCodeNum >= MaxPerPage Then Exit Do
            rsPostCode.MoveNext
        Loop
    End If
    rsPostCode.Close
    Set rsPostCode = Nothing
    
    Response.Write "</table>"

    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "条记录", True)

    Response.Write PostCodeSearch
End Sub

Sub PostCodeAdd()
    Response.Write "    <form method='post' action='" & FileName & "' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
    Response.Write "         <tr class='title'>" & vbCrLf
    Response.Write "            <td height='22' colspan='2'> " & vbCrLf
    Response.Write "               <div align='center'><strong>添加邮政编码</strong></div>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "          </tr>    " & vbCrLf
    Response.Write "         <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>所属省份：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Province' type='text' id='Province' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>所属城市：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='City' type='text' id='City' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>所属县区：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Area' type='text' id='Area' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>邮政编码：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='PostCode' type='text' id='PostCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>地区区号：</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='AreaCode' type='text' id='AreaCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>     " & vbCrLf
    Response.Write "                     <td colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "                     " & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SavePostCodeAdd'>        <input  type='submit' name='Submit' value=' 添 加 '  style='cursor:hand;'>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_City.asp'"" style='cursor:hand;'>" & vbCrLf
    Response.Write "                     </td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </form>" & vbCrLf
    Response.Write PostCodeSearch
End Sub


Sub SavePostCodeAdd()
    Dim Province, City, Area, PostCode, AreaCode, sql, rs
    Province = ReplaceBadChar(Trim(Request.Form("Province")))
    City = ReplaceBadChar(Trim(Request.Form("City")))
    Area = ReplaceBadChar(Trim(Request.Form("Area")))
    PostCode = ReplaceBadChar(Trim(Request.Form("PostCode")))
    AreaCode = ReplaceBadChar(Trim(Request.Form("AreaCode")))
    If Province = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写省份！"
        Exit Sub
    End If
    If City = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写城市！"
        Exit Sub
    End If
    If Area = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写县区！"
        Exit Sub
    End If
    If PostCode = "" Or Not IsTrueCode(PostCode, "PostCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写正确的邮政编码！"
        Exit Sub
    End If
    If AreaCode = "" Or Not IsTrueCode(AreaCode, "AreaCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写正确的区号！"
        Exit Sub
    End If
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_City where Province='" & Province & "' and City='" & City & "' and Area='" & Area & "'"
    rs.Open sql, Conn, 1, 3
    If rs.EOF And rs.BOF Then
        rs.AddNew
        rs("Country") = "中华人民共和国"
        rs("Province") = Province
        rs("City") = City
        rs("Area") = Area
        rs("PostCode") = PostCode
        rs("AreaCode") = AreaCode
        rs.Update
        Call WriteSuccessMsg("邮政编码添加成功！", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您输入的地区已经存在。</li>"
    End If
    rs.Close
    Set rs = Nothing
End Sub


Sub PostCodeEdit()
    Dim AreaID, PostCode, sql, rs

    AreaID = PE_CLng(Trim(Request("AreaID")))

    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_City where AreaID=" & AreaID & ""
    rs.Open sql, Conn, 1, 3
    If rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "不存在该记录！"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    Response.Write "    <form method='post' action='" & FileName & "' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
    Response.Write "         <tr class='title'>" & vbCrLf
    Response.Write "            <td height='22' colspan='2'> " & vbCrLf
    Response.Write "               <div align='center'><strong>修改邮政编码</strong></div>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>所属省份：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Province' value='" & rs("Province") & "' type='text' id='Province' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>所属城市：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='City' value='" & rs("City") & "' type='text' id='City' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "               <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>所属县区：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='Area' value='" & rs("Area") & "' type='text' id='Area' size='40' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>邮政编码：</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='PostCode' value='" & rs("PostCode") & "' type='text' id='PostCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='280' class='tdbg5' align='right'><strong>地区区号：</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='AreaCode' value='" & rs("AreaCode") & "' type='text' id='AreaCode' size='25' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>     " & vbCrLf
    Response.Write "                     <td colspan='2' align='center' class='tdbg'><input name='AreaID' type='hidden' id='AreaID' value='" & AreaID & "'>" & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SavePostCodeEdit'>        <input  type='submit' name='Submit' value='保存修改结果'  style='cursor:hand;'>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_City.asp'"" style='cursor:hand;'>" & vbCrLf
    Response.Write "                     </td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </form>" & vbCrLf
    rs.Close
    Set rs = Nothing
    Response.Write PostCodeSearch
End Sub

Sub SavePostCodeEdit()
    Dim AreaID, Province, City, Area, PostCode, AreaCode, sql, rs
    AreaID = PE_CLng(Trim(Request.Form("AreaID")))
    Province = ReplaceBadChar(Trim(Request.Form("Province")))
    City = ReplaceBadChar(Trim(Request.Form("City")))
    Area = ReplaceBadChar(Trim(Request.Form("Area")))
    PostCode = ReplaceBadChar(Trim(Request.Form("PostCode")))
    AreaCode = ReplaceBadChar(Trim(Request.Form("AreaCode")))
    If Province = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写省份！"
    End If
    If City = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写城市！"
    End If
    If PostCode = "" Or Not IsTrueCode(PostCode, "PostCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写正确的邮政编码！"
    End If
    If AreaCode = "" Or Not IsTrueCode(AreaCode, "AreaCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写正确的区号！"
    End If
    If FoundErr = True Then Exit Sub

    Dim trs
    Set trs = Conn.Execute("select top 1 AreaID from PE_City where AreaID<>" & AreaID & " and Province='" & Province & "' and City='" & City & "' and Area='" & Area & "'")
    If Not (trs.BOF And trs.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您输入的区域已经存在！</li>"
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub

    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_City where AreaID=" & AreaID & ""
    rs.Open sql, Conn, 1, 3
    If Not (rs.EOF And rs.BOF) Then
        rs("Province") = Province
        rs("City") = City
        rs("Area") = Area
        rs("PostCode") = PostCode
        rs("AreaCode") = AreaCode
        rs.Update
        Call WriteSuccessMsg("邮政编码修改成功！", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "修改失败，原始数据丢失。"
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub PostCodeDel()
    Dim AreaID, RowCount

    AreaID = PE_CLng(Trim(Request("AreaID")))

    Conn.Execute ("delete from PE_City where AreaID=" & AreaID & ""), RowCount
    If RowCount = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "记录删除失败。"
    Else
        Call WriteSuccessMsg("记录删除成功！", ComeUrl)
    End If
End Sub

Function IsTrueCode(thisCode, CodeType)
    Dim temp
    IsTrueCode = False
    If CodeType = "PostCode" Then
        regEx.Pattern = "^\d{6}$"
    Else
        regEx.Pattern = "^\d{3,7}$"
    End If

    IsTrueCode = regEx.Test(thisCode)
End Function

Function PostCodeSearch()
    Dim strHtml
    strHtml = "<br>"
    strHtml = strHtml & "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    strHtml = strHtml & "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    strHtml = strHtml & "  <tr class='tdbg'>"
    strHtml = strHtml & "   <td width='130' align='right'><strong>邮政编码搜索：</strong></td>"
    strHtml = strHtml & "   <td>"
    strHtml = strHtml & "<select name='Field' size='1'>"
    strHtml = strHtml & "<option value='Province'>所属省份</option>"
    strHtml = strHtml & "<option value='City'>所属城市</option>"
    strHtml = strHtml & "<option value='Area' selected>县区</option>"
    strHtml = strHtml & "<option value='PostCode'>邮政编码</option>"
    strHtml = strHtml & "<option value='AreaCode'>地区区号</option>"
    strHtml = strHtml & "</select>"
    strHtml = strHtml & "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    strHtml = strHtml & "<input type='submit' name='Submit'  value='搜索'>"
    strHtml = strHtml & "</td></tr></table></form>"
    PostCodeSearch = strHtml
End Function
%>
