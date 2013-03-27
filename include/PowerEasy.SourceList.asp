<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub UserList()
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title' height='22'>" & vbCrLf
    Response.Write "    <td valign='top'><b>已经选定的用户名：</b></td>" & vbCrLf
    Response.Write "    <td align='right'><a href='javascript:window.returnValue=myform.UserList.value;window.close();'>返回&gt;&gt;</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td><input type='text' name='UserList' size='60' maxlength='200' readonly='readonly'></td>" & vbCrLf
    Response.Write "    <td align='center'><input type='button' name='del1' onclick='del(1)' value='删除最后'> <input type='button' name='del2' onclick='del(0)' value='删除全部'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'><td>" & GetUserGroup() & "</td></tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>列表：</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='查找'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan=2>"
    Dim i, rsUser, sql
    sql = "select UserName from PE_User Where 1=1"
    If PE_CLng(Group) > 0 Then
        sql = sql & " and GroupID=" & PE_CLng(Group)
    End If
    If Keyword <> "" Then
        sql = sql & " and UserName like '%" & Keyword & "%'"
    End If
    sql = sql & " order by Userid"

    Set rsUser = Server.CreateObject("adodb.recordset")
    rsUser.Open sql, Conn, 1, 1
    If rsUser.BOF And rsUser.EOF Then
        totalPut = 0
        Response.Write "<li>没有任何用户</li>"
    Else
        totalPut = rsUser.RecordCount
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
            rsUser.Move (CurrentPage - 1) * MaxPerPage
        Else
                    CurrentPage = 1
                End If
        End If
        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'><tr>"
        Do While Not rsUser.EOF
            If AllUserList = "" Then
                AllUserList = rsUser("UserName")
            Else
                AllUserList = AllUserList & "," & rsUser("UserName")
            End If
            Response.Write "<td align='center'><a href='#' onclick='add(""" & rsUser("UserName") & """)'>" & rsUser("UserName") & "</a></td>"
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            If (i Mod 8) = 0 And i > 1 Then Response.Write "</tr><tr>"
            rsUser.MoveNext
        Loop
        Response.Write "</tr></table>"
    End If
    rsUser.Close
    Set rsUser = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td align='center' colspan=2><a href='#' onclick='add(""" & AllUserList & """)'>增加以上所有用户名</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个用户名", True)
    Call ShowJS("用户名")
End Sub

Sub AgentList()
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>列表：</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='查找'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan=2>"
    Dim i, rsUser, sql
    sql = "select U.UserName,G.GroupName from PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID Where G.GroupType=4"
    If Keyword <> "" Then
        sql = sql & " and U.UserName like '%" & Keyword & "%'"
    End If
    sql = sql & " order by U.UserID"
    
    Set rsUser = Server.CreateObject("adodb.recordset")
    rsUser.Open sql, Conn, 1, 1
    If rsUser.BOF And rsUser.EOF Then
        totalPut = 0
        Response.Write "<li>没有任何代理商</li>"
    Else
        totalPut = rsUser.RecordCount
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
            rsUser.Move (CurrentPage - 1) * MaxPerPage
        Else
                    CurrentPage = 1
                End If
        End If
        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'><tr>"
        Do While Not rsUser.EOF
            Response.Write "<td align='center'><a href='#' onclick=""window.returnValue='" & rsUser("UserName") & "';window.close();"")'>" & rsUser("UserName") & "</a></td>"
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            If (i Mod 8) = 0 And i > 1 Then Response.Write "</tr><tr>"
            rsUser.MoveNext
        Loop
        Response.Write "<td align='center'><a href='#' onclick=""window.returnValue='';window.close();"")'>无</a></td>"
        Response.Write "</tr></table>"
    End If
    rsUser.Close
    Set rsUser = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个代理商", True)
End Sub

Sub Key()
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title' height='22'>" & vbCrLf
    Response.Write "    <td valign='top'><b>已经选定的关键字：</b></td>" & vbCrLf
    Response.Write "    <td align='right'><a href='javascript:window.close();'>返回&gt;&gt;</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td><input type='text' name='KeyList' size='60' maxlength='200' readonly='readonly'></td>" & vbCrLf
    Response.Write "    <td align='center'><input type='button' name='del1' onclick='del(1)' value='删除最后'> <input type='button' name='del2' onclick='del(0)' value='删除全部'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' class='title'><td>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=AddTime'><FONT style='font-size:12px'" & vbCrLf
    If Group = "AddTime" Then Response.Write "color='red'>按发布时间排序</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Hits'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Hits" Then Response.Write "color='red'>按使用频率排序</FONT></a>" & vbCrLf
    Response.Write "         | </td></tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>列表：</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='查找'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan=2>"
    
    
    Dim i, rsKey, sql
    If Group = "AddTime" Or Group = "" Then
        sql = "select * from PE_NewKeys Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then
            sql = sql & " and KeyText like '%" & Keyword & "%'"
        End If
        sql = sql & " order by LastUseTime Desc"
    ElseIf Group = "Hits" Then
        sql = "select * from PE_NewKeys Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then
            sql = sql & " and KeyText like '%" & Keyword & "%'"
        End If
        sql = sql & " order by Hits Desc"
    End If
    Set rsKey = Server.CreateObject("adodb.recordset")
    rsKey.Open sql, Conn, 1, 1
    If rsKey.BOF And rsKey.EOF Then
        totalPut = 0
        Response.Write "<li>没有关键子</li>"
    Else
        totalPut = rsKey.RecordCount
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
                rsKey.Move (CurrentPage - 1) * MaxPerPage
            Else
                        CurrentPage = 1
                    End If
            End If
        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'><tr>"
        Do While Not rsKey.EOF
            If AllKeyList = "" Then
                AllKeyList = rsKey("KeyText")
            Else
                AllKeyList = AllKeyList & "|" & rsKey("KeyText")
            End If
            Response.Write "<td align='center'><a href='#' onclick='add(""" & rsKey("KeyText") & """)'>" & rsKey("KeyText") & "</a></td>"
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            If (i Mod 8) = 0 And i > 1 Then Response.Write "</tr><tr>"
            rsKey.MoveNext
        Loop
        Response.Write "</tr></table>"
    End If
    rsKey.Close
    Set rsKey = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td align=center colspan=2><a href='#' onclick='add(""" & AllKeyList & """)'>增加以上所有关键字</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个关键字", True)
    
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "myform.KeyList.value=opener.myform.Keyword.value;" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    if(opener.myform.Keyword.value=="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        opener.myform.Keyword.value=obj;" & vbCrLf
    Response.Write "        myform.KeyList.value=opener.myform.Keyword.value;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var singleKey=obj.split(""|"");" & vbCrLf
    Response.Write "    var ignoreKey="""";" & vbCrLf
    Response.Write "    for(i=0;i<singleKey.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(checkKey(opener.myform.Keyword.value,singleKey[i]))" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            ignoreKey=ignoreKey+singleKey[i]+"" """ & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            opener.myform.Keyword.value=opener.myform.Keyword.value+""|""+singleKey[i];" & vbCrLf
    Response.Write "            myform.KeyList.value=opener.myform.Keyword.value;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ignoreKey!="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert(ignoreKey+"" 关键字已经存在，此操作已经忽略！"");" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function del(num)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (num==0 || opener.myform.Keyword.value=="""" || opener.myform.Keyword.value==""|"")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        opener.myform.Keyword.value="""";" & vbCrLf
    Response.Write "        myform.KeyList.value="""";" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    var strDel=opener.myform.Keyword.value;" & vbCrLf
    Response.Write "    var s=strDel.split(""|"");" & vbCrLf
    Response.Write "    opener.myform.Keyword.value=strDel.substring(0,strDel.length-s[s.length-1].length-1);" & vbCrLf
    Response.Write "    myform.KeyList.value=opener.myform.Keyword.value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function checkKey(Keylist,thisKey)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if (Keylist==thisKey){" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    var s=Keylist.split(""|"");" & vbCrLf
    Response.Write "    for (j=0;j<s.length;j++){" & vbCrLf
    Response.Write "        if(s[j]==thisKey)" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Author()
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title' height='22'>" & vbCrLf
    Response.Write "    <td valign='top'><b>已经选定的作者：</b></td>" & vbCrLf
    Response.Write "    <td align='right'><a href='javascript:window.close();'>返回&gt;&gt;</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td><input type='text' name='AuthorList' size='60' maxlength='200' readonly='readonly'></td>" & vbCrLf
    Response.Write "    <td align='center'><input type='button' name='del1' onclick='del(1)' value='删除最后'> <input type='button' name='del2' onclick='del(0)' value='删除全部'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' class='title'><td>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Time'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Time" Then Response.Write " color='red'"
    Response.Write ">最近常用</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=All'><FONT style='font-size:12px'" & vbCrLf
    If Group = "All" Then Response.Write " color='red'"
    Response.Write ">全部作者</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Site'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Site" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType4", "本站特约") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=MLand'><FONT style='font-size:12px'" & vbCrLf
    If Group = "MLand" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType1", "大陆作者") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Gt'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Gt" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType2", "港台作者") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=OutSea'><FONT style='font-size:12px'" & vbCrLf
    If Group = "OutSea" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType3", "海外作者") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Other'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Other" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowAuthor/AuthorType5", "其他作者") & "</FONT></a>" & vbCrLf
    Response.Write "         | </td></tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>列表：</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='查找'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    Dim i, rsAuthor, sql
    Select Case Group
    Case "Time"
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ", LastUseTime Desc")
    Case "All"
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc")
    Case "MLand"
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0) and AuthorType=1"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc")
    Case "Gt"
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0) and AuthorType=2"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc")
    Case "OutSea"
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0) and AuthorType=3"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc")
    Case "Site"
        sql = "select AuthorName,Sex,Intro from PE_Author Where(ChannelID=" & ChannelID & " or ChannelID=0) and AuthorType=4"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc")
    Case "Other"
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0) and AuthorType=0"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc")
    Case Else
        sql = "select AuthorName,Sex,Intro from PE_Author Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and AuthorName like '%" & Keyword & "%'")
        sql = sql & (" and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",LastUseTime Desc")
    End Select
    Set rsAuthor = Server.CreateObject("adodb.recordset")
    rsAuthor.Open sql, Conn, 1, 1
    If rsAuthor.BOF And rsAuthor.EOF Then
        totalPut = 0
        Response.Write "<li>没有作者</li>"
    Else
        totalPut = rsAuthor.RecordCount
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
                rsAuthor.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>"
        Response.Write "<tr align='center'><td width='150' >姓名</td><td width='35'>性别</td><td>简介</td></tr>"
        Do While Not rsAuthor.EOF
            If AllKeyList = "" Then
                AllKeyList = rsAuthor("AuthorName")
            Else
                AllKeyList = AllKeyList & "|" & rsAuthor("AuthorName")
            End If
            Response.Write "<tr onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""><td align='center'><a href='#' onclick='add(""" & rsAuthor("AuthorName") & """)'>" & rsAuthor("AuthorName") & "</a></td><td align='center'>"
            If rsAuthor("Sex") = 0 Then
                Response.Write "女</td>"
            Else
                Response.Write "男</td>"
            End If
            If IsNull(rsAuthor("Intro")) Then
                Response.Write "<td>无</td></tr>"
            Else
                Response.Write "<td>" & Left(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), 40) & "</td></tr>"
            End If
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsAuthor.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsAuthor.Close
    Set rsAuthor = Nothing
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个作者", True)

    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "myform.AuthorList.value=opener.myform.Author.value;" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    if(opener.myform.Author.value=="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        opener.myform.Author.value=obj;" & vbCrLf
    Response.Write "        myform.AuthorList.value=opener.myform.Author.value;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var singleKey=obj.split(""|"");" & vbCrLf
    Response.Write "    var ignoreKey="""";" & vbCrLf
    Response.Write "    for(i=0;i<singleKey.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(checkKey(opener.myform.Author.value,singleKey[i]))" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            ignoreKey=ignoreKey+singleKey[i]+"" """ & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            opener.myform.Author.value=opener.myform.Author.value+""|""+singleKey[i];" & vbCrLf
    Response.Write "            myform.AuthorList.value=opener.myform.Author.value;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ignoreKey!="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert(ignoreKey+"" 该作者已经存在，此操作已经忽略！"");" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function del(num)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (num==0 || opener.myform.Author.value=="""" || opener.myform.Author.value==""|"")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        opener.myform.Author.value="""";" & vbCrLf
    Response.Write "        myform.AuthorList.value="""";" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    var strDel=opener.myform.Author.value;" & vbCrLf
    Response.Write "    var s=strDel.split(""|"");" & vbCrLf
    Response.Write "    opener.myform.Author.value=strDel.substring(0,strDel.length-s[s.length-1].length-1);" & vbCrLf
    Response.Write "    myform.AuthorList.value=opener.myform.Author.value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function checkKey(Keylist,thisKey)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if (Keylist==thisKey){" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    var s=Keylist.split(""|"");" & vbCrLf
    Response.Write "    for (j=0;j<s.length;j++){" & vbCrLf
    Response.Write "        if(s[j]==thisKey)" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub CopyFrom()
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    opener.myform.CopyFrom.value=obj;" & vbCrLf
    Response.Write "    window.close();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' class='title'><td>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Time'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Time" Then Response.Write " color='red'"
    Response.Write ">最近常用</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=All'><FONT style='font-size:12px'" & vbCrLf
    If Group = "All" Then Response.Write " color='red'"
    Response.Write ">全部来源</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Site'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Site" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "友情站点") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=MLand'><FONT style='font-size:12px'" & vbCrLf
    If Group = "MLand" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "中文站点") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=OutSea'><FONT style='font-size:12px'" & vbCrLf
    If Group = "OutSea" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "外文站点") & "</FONT></a>" & vbCrLf
    Response.Write "     | <a href='" & FileName & "?ChannelID=" & ChannelID & "&TypeSelect=" & TypeSelect & "&Group=Other'><FONT style='font-size:12px'" & vbCrLf
    If Group = "Other" Then Response.Write " color='red'"
    Response.Write ">" & XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "其他来源") & "</FONT></a>" & vbCrLf
    Response.Write "         | </td></tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr  height='22' class='title'>" & vbCrLf
    Response.Write "    <td><b><font color=red>" & strTypeName & "</font>列表：</b></td><td align=right><input name='KeyWord' type='text' size='20' value=" & Keyword & ">&nbsp;&nbsp;<input type='submit' value='查找'></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td valign='top' height='100' colspan='2'>"
    
    
    Dim i, rsCopyFrom, sql
    Select Case Group
    Case "Time"
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",LastUseTime Desc"
    Case "All"
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc"
    Case "Site"
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0) and SourceType=1"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc"
    Case "MLand"
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0) and SourceType=2"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc"
    Case "OutSea"
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0) and SourceType=3"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc"
    Case "Other"
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0) and SourceType=0"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc"
    Case Else
        sql = "select * from PE_CopyFrom Where (ChannelID=" & ChannelID & " or ChannelID=0)"
        If Keyword <> "" Then sql = sql & (" and SourceName like '%" & Keyword & "%'")
        sql = sql & " and Passed=" & PE_True & " order by onTop " & PE_OrderType & ",ID Desc"
    End Select
    Set rsCopyFrom = Server.CreateObject("adodb.recordset")
    rsCopyFrom.Open sql, Conn, 1, 1
    If rsCopyFrom.BOF And rsCopyFrom.EOF Then
        totalPut = 0
        Response.Write "<li>没有来源</li>"
    Else
        totalPut = rsCopyFrom.RecordCount
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
                rsCopyFrom.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Response.Write "<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>"
        Response.Write "<tr align='center'><td width='100' >名称</td><td width='100'>联系人</td><td>简介</td></tr>"
        Do While Not rsCopyFrom.EOF
            If AllKeyList = "" Then
                AllKeyList = rsCopyFrom("SourceName")
            Else
                AllKeyList = AllKeyList & "|" & rsCopyFrom("SourceName")
            End If
            Response.Write "<tr><td align='center'><a href='#' onclick='add(""" & rsCopyFrom("SourceName") & """)'>" & rsCopyFrom("SourceName") & "</a></td><td>" & rsCopyFrom("ContacterName") & "</td>"
            If IsNull(rsCopyFrom("Intro")) Then
                Response.Write "<td>无</td></tr>"
            Else
                Response.Write "<td>" & Left(nohtml(PE_HtmlDecode(rsCopyFrom("Intro"))), 50) & "</td></tr>"
            End If
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsCopyFrom.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
    
    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write ShowSourcePage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个来源", True)

End Sub

Function GetUserGroup()
    Dim strGroup, rsGroup, i
    i = 0
    strGroup = "<table><tr>"
    Set rsGroup = Conn.Execute("select GroupID,GroupName from PE_UserGroup order by GroupType,GroupID")
    Do While Not rsGroup.EOF
        strGroup = strGroup & "<td>|  <a href='#' onclick='myform.Group.value=" & rsGroup(0) & ";myform.submit();'><FONT style='font-size:12px'"
        If Group = rsGroup(0) Then
            strGroup = strGroup & "color='red'"
        End If
        strGroup = strGroup & ">" & rsGroup(1) & "</FONT></a></td>"
        rsGroup.MoveNext
        i = i + 1
        If i Mod 5 = 0 Then strGroup = strGroup & "</tr><tr>"
    Loop
    Set rsGroup = Nothing
    strGroup = strGroup & "</table><input type='hidden' name='Group' value='0'>"
    GetUserGroup = strGroup
End Function

Function ShowSourcePage(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowSourcePage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
        
    strTemp = "<div class=""show_page"">"
    If ShowTotal = True Then
        strTemp = strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "首页 上一页&nbsp;"
    Else
        strTemp = strTemp & "<a href='#' onclick='myform.page.value=1;myform.submit();'>首页</a>&nbsp;"
        strTemp = strTemp & "<a href='#' onclick='myform.page.value=" & (CurrentPage - 1) & ";myform.submit();'>上一页</a>&nbsp;"
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "下一页 尾页"
    Else
        strTemp = strTemp & "<a href='#' onclick='myform.page.value=" & (CurrentPage + 1) & ";myform.submit();'>下一页</a>&nbsp;"
        strTemp = strTemp & "<a href='#' onclick='myform.page.value=" & TotalPage & ";myform.submit();'>尾页</a>"
    End If
    strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>页 "
    If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;<input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress='if (event.keyCode==13) myform.submit();'>" & strUnit & "/页"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/页"
    End If
    If ShowAllPages = True Then
        strTemp = strTemp & "&nbsp;&nbsp;转到第<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) myform.submit();"" onmousewheel=""if ((parseInt(this.value) + parseInt(event.wheelDelta/120))>0&&(parseInt(this.value) + parseInt(event.wheelDelta/120))<=" & TotalPage & ") this.value=parseInt(this.value) + parseInt(event.wheelDelta/120);"">页"
    End If
    strTemp = strTemp & "</div>"
    ShowSourcePage = strTemp
End Function

Sub ShowJS(strName)
    Response.Write "<script language=""javascript"">" & vbCrLf
    If Trim(Request("UserList")) <> "" Then
        Response.Write "myform.UserList.value='" & Trim(Request("UserList")) & "';" & vbCrLf
    Else
        Response.Write "myform.UserList.value='" & Trim(Request("DefaultValue")) & "';" & vbCrLf
    End If
    Response.Write "var oldUser='';" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==''){return false;}" & vbCrLf
    Response.Write "    if(myform.UserList.value=='')" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        myform.UserList.value=obj;" & vbCrLf
    Response.Write "        window.returnValue=myform.UserList.value;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var singleUser=obj.split(',');" & vbCrLf
    Response.Write "    var ignoreUser='';" & vbCrLf
    Response.Write "    for(i=0;i<singleUser.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(checkUser(myform.UserList.value,singleUser[i]))" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            ignoreUser=ignoreUser+singleUser[i]+"" """ & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            myform.UserList.value=myform.UserList.value+','+singleUser[i];" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ignoreUser!='')" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert(ignoreUser+'" & strName & "已经存在，此操作已经忽略！');" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    window.returnValue=myform.UserList.value;" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function del(num)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (num==0 || myform.UserList.value=='' || myform.UserList.value==',')" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        myform.UserList.value='';" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    var strDel=myform.UserList.value;" & vbCrLf
    Response.Write "    var s=strDel.split(',');" & vbCrLf
    Response.Write "    myform.UserList.value=strDel.substring(0,strDel.length-s[s.length-1].length-1);" & vbCrLf
    Response.Write "    window.returnValue=myform.UserList.value;" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function checkUser(UserList,thisUser)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if (UserList==thisUser){" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    var s=UserList.split(',');" & vbCrLf
    Response.Write "    for (j=0;j<s.length;j++){" & vbCrLf
    Response.Write "        if(s[j]==thisUser)" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub
%>
