<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ClassID, Passed
Dim tClass, ClassName, ReadMe, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ChildID, tID, tChild

Sub Execute()
    ChannelID = PE_CLng(Trim(Request("ChannelID")))
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要查看的频道ID！</li>"
        Response.Write ErrMsg
        Exit Sub
    End If


    ClassID = PE_CLng(Trim(Request("ClassID")))
    Passed = Trim(Request("Passed"))
    Session("Passed") = Passed
    FileName = "User_Comment.asp?ChannelID=" & ChannelID
    strFileName = "User_Comment.asp?ChannelID=" & ChannelID & "&Field=" & strField & "&keyword=" & Keyword

    Select Case Action
    Case "Modify"
        Call Modify
    Case "SaveModify"
        Call SaveModify
    Case "Del"
        Call DelComment
    Case Else
        Call main
    End Select
     
    If FoundErr = True Then
       Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub


Sub main()
    Dim rs, sql
    If ClassID > 0 Then
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的栏目</li>"
            Exit Sub
        Else
            ClassName = tClass("ClassName")
            RootID = tClass("RootID")
            ParentID = tClass("ParentID")
            Depth = tClass("Depth")
            ParentPath = tClass("ParentPath")
            Child = tClass("Child")
            arrChildID = tClass("arrChildID")
        End If
    End If

    Select Case ModuleType
    Case 1
        sql = "select I.Title as ObjectTitle,I.IncludePic"
    Case 2
        sql = "select I.SoftName as ObjectTitle"
    Case 3
        sql = "select I.PhotoName as ObjectTitle"
    Case 5
        sql = "select I.ProductName as ObjectTitle"
    Case 6
        sql = "Select I.SupplyTitle as ObjectTitle,I.SupplyID as ObjectID,C.CommentID,C.UserType,C.UserName,C.Email,C.Oicq,C.Homepage,C.Icq,C.Msn,C.IP,C.Content,C.WriteTime,C.ReplyName,C.ReplyContent,C.ReplyTime,C.Score,C.Passed From PE_Comment C Inner Join PE_Supply I On C.InfoID=I.SupplyID Where I.UserName='" & UserName & "'"

    End Select
    If ModuleType <> 6 Then
        sql = sql & ",I." & ModuleName & "ID as ObjectID,C.CommentID,C.UserType,C.UserName,C.Email,C.Oicq,C.Homepage,C.Icq,C.Msn,C.IP"
        sql = sql & ",C.Content,C.WriteTime,C.ReplyName,C.ReplyContent,C.ReplyTime,C.Score,C.Passed"
        sql = sql & " from PE_Comment C inner join " & SheetName & " I on C.InfoID=I." & ModuleName & "ID"
        sql = sql & " where I.ChannelID=" & ChannelID
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "CommentContent"
            sql = sql & " and C.Content like '%" & Keyword & "%' "
        Case "CommentTime"
            If IsDate(Trim(Request("keyword"))) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>输入的关键字不是有效日期！</li>"
                Exit Sub
            Else
                If SystemDatabaseType = "SQL" Then
                    sql = sql & " and C.WriteTime = '" & Trim(Request("keyword")) & "' "
                Else
                    sql = sql & " and C.WriteTime = #" & Trim(Request("keyword")) & "# "
                End If
            End If
        Case Else
            sql = sql & " and C.Content like '%" & Keyword & "%' "
        End Select
    End If
    If Passed = "True" Then
        sql = sql & " and C.Passed =" & PE_True & ""
    ElseIf Passed = "False" Then
        sql = sql & " and C.Passed =" & PE_False & ""
    End If

    If ClassID > 0 Then
        If Child > 0 Then
            sql = sql & " and I.ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and I.ClassID=" & ClassID
        End If
    End If
    If ModuleType <> 6 Then
        sql = sql & " and C.UserName='" & UserName & "' order by " & ModuleName & "ID desc"
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn, 1, 1
    
    Call ShowJS_Main("评论")
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetRootClass() & "</td>"
    Response.Write "  </tr>" & GetChild_Root() & ""
    Response.Write "</table>"
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "<form name='myform' method='post' action='" & strFileName & "' onsubmit='return ConfirmDel();'>"
    Response.Write "  <tr>"
    Response.Write "    <td align='center'>"
    Response.Write "      <table border='0' cellpadding='2' width='100%' cellspacing='0'>"
    Response.Write "        <tr><td>"
    Response.Write "评论管理&nbsp;&gt;&gt;&nbsp;"
    If ClassID > 0 Then
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Server.CreateObject("adodb.recordset")
            rsPath.Open sqlPath, Conn, 1, 1
            Do While Not rsPath.EOF
                Response.Write "<a href='" & FileName & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        Response.Write "<a href='" & FileName & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If Keyword = "" Then
        Response.Write "所有评论"
    Else
        Select Case strField
            Case "CommentContent"
                Response.Write "评论内容中含有 <font color=red>" & Keyword & "</font> 的评论"
            Case "CommentName"
                Response.Write "评论人中含有 <font color=red>" & Keyword & "</font> 的评论"
            Case Else
                Response.Write "评论中含有 <font color=red>" & Keyword & "</font> 的评论"
        End Select
    End If
    Response.Write "          </td><td width='150' align='right'>"
    If rs.BOF And rs.EOF Then
        Response.Write "共找到 0 篇评论</td></tr></table>"
    Else
        totalPut = rs.RecordCount
        Response.Write "共找到 " & totalPut & " 篇评论</td></tr></table>"
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
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim CommentNum, rsCommentUser
        CommentNum = 0
        Dim PrevID, iTemp
        PrevID = rs("ObjectID")
        Do While Not rs.EOF
            If rs("ObjectID") <> PrevID Then Response.Write "</table></td></tr></table><br>"
            If CommentNum = 0 Or rs("ObjectID") <> PrevID Then
                iTemp = 1
                Response.Write "      <table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
                Response.Write "        <tr class='title'>"
                Response.Write "          <td width='80%' height='22'>"
                If ModuleType = 1 Then
                    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
                    XmlDoc.async = False
                    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
                    Select Case rs("IncludePic")
                        Case 1
                            Response.Write "<font color=blue>" & XmlText("Article", "ArticlePro1", "[图文]") & "</font>"
                        Case 2
                            Response.Write "<font color=blue>" & XmlText("Article", "ArticlePro2", "[组图]") & "</font>"
                        Case 3
                            Response.Write "<font color=blue>" & XmlText("Article", "ArticlePro3", "[推荐]") & "</font>"
                        Case 4
                            Response.Write "<font color=blue>" & XmlText("Article", "ArticlePro4", "[注意]") & "</font>"
                    End Select
                    Set XmlDoc = Nothing
                End If
                Response.Write rs("ObjectTitle")
                Response.Write "          </td>"
                Response.Write "          <td width='20%' align='right'></td>"
                Response.Write "        </tr>"
                Response.Write "        <tr>"
                Response.Write "          <td colspan='2'>"
                Response.Write "            <table border='0' cellspacing='1' width='100%' cellpadding='0' style='word-break:break-all'>"
            End If
            Response.Write "              <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "                <td width='20' align='center'>" & iTemp & "</td>"
            Response.Write "                <td><a href='#' Title='" & Left(rs("Content"), 200) & "'>评论内容：" & Left(rs("Content"), 30) & "</a></td>"
            Response.Write "                <td width='70' align='center'>评分：" & rs("Score") & "</td>"
            Response.Write "                <td width='160' align='center'>时间：" & rs("WriteTime") & "</td>"
            Response.Write "                <td width='60' align='center'>"
            If rs("Passed") = True Then
                Response.Write "已审核"
            Else
                Response.Write "<font color='red'>未审核</font>"
            End If
            Response.Write "</td>"
            Response.Write "                <td width='120' align='center'>"
            Response.Write "&nbsp;&nbsp;&nbsp;"

            If rs("Passed") <> True Then
                Response.Write "<a href='" & strFileName & "&Action=Modify&CommentID=" & rs("CommentID") & "'>修改</a>&nbsp;"
                Response.Write "<a href='" & strFileName & "&Action=Del&CommentID=" & rs("CommentID") & "' onclick=""return confirm('确定要删除此评论吗？');"">删除</a>&nbsp;"
            End If
            Response.Write "                </td>"
            Response.Write "              </tr>"
            If rs("ReplyContent") <> "" Then
                Response.Write "            <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
                Response.Write "              <td align='center'>&nbsp;</td>"
                Response.Write "                <td colspan='2'>"
                Response.Write "                回复：" & rs("ReplyContent") & ""
                Response.Write "                </td>"
                Response.Write "                <td align='center'>管理员：" & rs("ReplyName") & "</td>"
                Response.Write "                <td align='center' colspan='2'>回复时间：" & rs("ReplyTime") & "</td>"
                Response.Write "              </tr>"
            End If
            CommentNum = CommentNum + 1
            If CommentNum >= MaxPerPage Then Exit Do
            PrevID = rs("ObjectID")
            iTemp = iTemp + 1
            rs.MoveNext
        Loop
        Response.Write "            </table>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "      </table>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个评论", True)
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>评论搜索：</strong></td><td>"
    Response.Write "<table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='User_Comment.asp'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='CommentContent' selected>评论内容</option>"
    Response.Write "<option value='CommentTime'>评论时间</option>"
    Response.Write "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='搜索'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></form></table>"

    Response.Write "  </td></tr>"
    Response.Write "</table>"
End Sub

Sub Modify()
    Dim rs, sql
    Dim CommentID
    CommentID = Trim(Request("CommentID"))
    If IsValidID(CommentID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定评论ID</li>"
        Exit Sub
    Else
        CommentID = PE_CLng(CommentID)
    End If
    sql = "Select * from PE_Comment where UserName='" & UserName & "' and CommentID=" & CommentID
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1
    If rs.BOF Or rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的评论！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' style='word-break:break-all;Width:fixed'>"
    Response.Write "  <form name='myform' method='post' action='" & strFileName & "'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td height='22' colspan='4'> <strong>修 改 评 论 </strong>&nbsp;&nbsp;"

    Response.Write "    </td>"
    Response.Write "  </tr>"

    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>评论人姓名：</td>"
    Response.Write "    <td class='tdbg' colspan='3'>"
    Response.Write "      <input name='ShowUserName' type='text' id='UserName' value='" & rs("UserName") & "' disabled>"
    Response.Write "      <input name='UserName' type='hidden' id='UserName' value='" & rs("UserName") & "'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>评论时间：</td>"
    Response.Write "    <td class='tdbg' width='200'>"
    Response.Write "      <input name='WriteTime' type='text' id='WriteTime' value='" & rs("WriteTime") & "'>"
    Response.Write "    </td>"
    Response.Write "    <td class='tdbg' align='right' width='101'>评论人IP：</td>"
    Response.Write "    <td class='tdbg' width='475'>"
    Response.Write "      <input name='IP' type='text' id='IP' maxlength='15' value='" & rs("IP") & "'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"

    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>评 分：</td>"
    Response.Write "    <td class='tdbg' colspan='3'>"
    Response.Write "      <input type='radio' name='Score' value='1' "
    If rs("Score") = 1 Then Response.Write " checked"
    Response.Write "      >1分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='2' "
    If rs("Score") = 2 Then Response.Write " checked"
    Response.Write "      >2分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='3' "
    If rs("Score") = 3 Then Response.Write " checked"
    Response.Write "      >3分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='4' "
    If rs("Score") = 4 Then Response.Write " checked"
    Response.Write "      >4分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='5' "
    If rs("Score") = 5 Then Response.Write " checked"
    Response.Write "      >5分 </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>评论内容：</td>"
    Response.Write "    <td class='tdbg' colspan='3'>"
    Response.Write "      <textarea name='Content' cols='56' rows='8' id='Content'>" & rs("Content") & "</textarea>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td height='30' colspan='4' class='tdbg'>"
    Response.Write "      <input name='ComeUrl' type='hidden' id='ComeUrl' value='" & ComeUrl & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "      <input name='CommentID' type='hidden' id='CommentID' value='" & rs("CommentID") & "'>"
    Response.Write "      <input name='UserType' type='hidden' id='UserType' value='" & rs("UserType") & "'>"
    Response.Write "      <input  type='submit' name='Submit' value=' 保存修改结果 '>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub DelComment()
    Dim CommentID
    CommentID = Trim(Request("CommentID"))
    If IsValidID(CommentID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定评论ID</li>"
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    Conn.Execute "delete from PE_Comment where UserName='" & UserName & "' and CommentID=" & PE_CLng(CommentID)
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub SaveModify()
    Dim rsComment, sql, ClassID, tClass, CommentID
    Dim CommentUserType, CommentUserName, CommentUserSex, CommentUserEmail, CommentUserOicq
    Dim CommentUserIcq, CommentUserMsn, CommentUserHomepage, CommentUserScore, CommentUserContent
    Dim CommentUserIP, CommentWritetime
    CommentID = Trim(Request("CommentID"))
    If IsValidID(CommentID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定评论ID</li>"
        Exit Sub
    Else
        CommentID = PE_CLng(CommentID)
    End If
    CommentUserName = Trim(Request("UserName"))
    If CommentUserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入姓名</li>"
        Exit Sub
    End If
    CommentUserType = PE_CLng(Request("UserType"))
    If CommentUserType = 0 Then
        CommentUserSex = Trim(Request("Sex"))
        CommentUserOicq = Trim(Request("Oicq"))
        CommentUserIcq = Trim(Request("Icq"))
        CommentUserMsn = Trim(Request("Msn"))
        CommentUserEmail = Trim(Request("Email"))
        CommentUserHomepage = Trim(Request("Homepage"))
        If CommentUserHomepage = "http://" Or IsNull(CommentUserHomepage) Then CommentUserHomepage = ""
    End If
    CommentUserIP = Trim(Request.Form("IP"))
    CommentWritetime = PE_CDate(Trim(Request.Form("WriteTime")))
    CommentUserScore = PE_CLng(Request.Form("Score"))
    CommentUserContent = Trim(Request.Form("Content"))
    If CommentUserContent = "" Or CommentUserIP = "" Or CommentUserScore = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请完整输入评论内容、评论时间、评论人IP等信息</li>"
    End If
    CommentUserContent = PE_HTMLEncode(CommentUserContent)

    If FoundErr = True Then
        Exit Sub
    End If

    sql = "Select * from PE_Comment where UserName='" & UserName & "' and CommentID=" & CommentID
    Set rsComment = Server.CreateObject("Adodb.RecordSet")
    rsComment.Open sql, Conn, 1, 3
    If rsComment.BOF Or rsComment.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的评论！</li>"
    Else
        rsComment("UserType") = CommentUserType
        rsComment("UserName") = CommentUserName
        rsComment("Sex") = CommentUserSex
        rsComment("Oicq") = CommentUserOicq
        rsComment("Icq") = CommentUserIcq
        rsComment("Msn") = CommentUserMsn
        rsComment("Email") = CommentUserEmail
        rsComment("Homepage") = CommentUserHomepage
        rsComment("IP") = CommentUserIP
        rsComment("WriteTime") = CommentWritetime
        rsComment("Score") = CommentUserScore
        rsComment("Content") = ReplaceText(CommentUserContent, 3)
        rsComment.Update
    End If
    rsComment.Close
    Set rsComment = Nothing
    Call CloseConn
    Response.Redirect strFileName
End Sub

%>
