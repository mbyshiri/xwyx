<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<%
Const NeedCheckComeUrl = True  '�Ƿ���Ҫ����ⲿ����
Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Response.Write "<html><head><title>�ʼ����Ĺ���</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'>" & vbCrLf
Response.Write "    <td height='22' colspan='2' align='center'><strong>�� �� �� �� �� ��</strong>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Action = Trim(request("Action"))
Select Case Action
Case "Send"
    Call SendMaillist
Case "Preview"
    Call PreviewMail
Case "user"
    Call UserList
Case "SetChannel"
    Call SetChannel
Case "SaveSet"
    Call SaveSet
Case Else
    Call main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsChannelList, sqlChannelList
    sqlChannelList = "select M.ChannelID,M.IsUse, C.ChannelName,M.UserID from PE_MailChannel M left join PE_Channel C On  C.ChannelID = M.ChannelID  order by OrderID"
    Set rsChannelList = Conn.Execute(sqlChannelList)
    Response.Write "  <form name='myform' method='post' onSubmit='return CheckForm();' action=''>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td align='center'><strong>Ƶ������</strong></td>"
    Response.Write "    <td width='100' align='center'><strong>��������</strong></td>"
    Response.Write "    <td width='100' align='center'><strong>�Ƿ�����</strong></td>"
    Response.Write "    <td align='center'><strong>����</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    Do While Not rsChannelList.EOF
        Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td align='center'>" & rsChannelList("ChannelName") & "</td>"
        Response.Write "    <td align='center'>"
        Dim MailNum, arrMailNum
        If rsChannelList("UserID") = "" Or IsNull(rsChannelList("UserID")) Then
            MailNum = 0
        Else
            arrMailNum = Split(rsChannelList("UserID"), ",")
            MailNum = UBound(arrMailNum) + 1
        End If
        Response.Write "   " & MailNum & " </td>"
        Response.Write "    <td align='center'>"
        If rsChannelList("IsUse") = PE_CBool(PE_True) Then Response.Write "��"
        Response.Write "    </td>"
        Response.Write "    <td align='center'>"
        Response.Write "    <a href='Admin_Mail.asp?Action=user&iChannelID=" & rsChannelList("ChannelID") & "'>�г�������</a>"
        Response.Write "    &nbsp;|&nbsp;"
        Response.Write "    <a href='Admin_Mail.asp?Action=Preview&iChannelID=" & rsChannelList("ChannelID") & "'>�ʼ�����Ԥ��</a>"
        Response.Write "    &nbsp;|&nbsp;"
        Response.Write "   <a href='Admin_Mail.asp?Action=Send&iChannelID=" & rsChannelList("ChannelID") & "'> ���Ͷ����ʼ�</a>"
        Response.Write "    &nbsp;|&nbsp;"
        Response.Write "   <a href='Admin_Mail.asp?Action=SetChannel&iChannelID=" & rsChannelList("ChannelID") & "'> Ƶ������</a>"
        Response.Write "    </td>"
        Response.Write "</tr>"
        rsChannelList.MoveNext
    Loop
    Response.Write "</table>"
    rsChannelList.Close
    Set rsChannelList = Nothing
End Sub

Sub SendMaillist()
    Dim iChannelID, rsChannel
    iChannelID = PE_Clng(Trim(request("iChannelID")))
    If iChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���͵�Ƶ��ID</li>"
        Exit Sub
    Else
        iChannelID = PE_Clng(iChannelID)
    End If
    Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & iChannelID)
    If rsChannel.bof And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����Ƶ����</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Sub
    End If
    Dim i, j, k, rs
    i = 0
    j = 0
    k = 0
    Dim usql, rsu, UserName, umail
    usql = "select * from PE_MailChannel where ChannelID =" & iChannelID
    Set rsu = Server.CreateObject("adodb.recordset")
    rsu.Open usql, Conn, 1, 3
    If rsu.bof And rsu.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Ƶ����û�ж�����</li>"
        rsu.Close
        Set rsu = Nothing
    Else
        UserID = rsu("UserID")
        Dim PE_Mail
        Set PE_Mail = New SendMail
        If UserID = "" Or IsNull(UserID) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����б�Ϊ��</li>"
        Else
            Response.Write "<li>���ڷ����У���ȴ�</li>"
            Dim ArrUserID, intTemp, sqlMail, rsMail

            ArrUserID = Split(UserID, ",")
            For intTemp = 0 To UBound(ArrUserID)
                sqlMail = "select * from PE_User where [UserID] =" &  PE_Clng(ArrUserID(intTemp))
                Set rsMail = Server.CreateObject("adodb.recordset")
                rsMail.Open sqlMail, Conn, 1, 1
                If rsMail.bof And rsMail.EOF Then
                    Dim arrUser, UserNum, tempUserID

                    arrUser = ""
                    tempUserID = Split(UserID, ",")
                    UserNum = 0
                    If UserNum <> UBound(tempUserID) Then
                        If arrUser = "" Then
                            arrUser = UserID
                        Else
                            If tempUsertempID(0) <> TempArr(intTemp) Then
                                tempUser = arrUser & "," & ArrUserID(0)
                            End If
                        End If
                        UserNum = UserNum + 1
                    End If
                    rsu("UserID") = arrUser
                    rsu.Update
                Else
                    umail = Trim(rsMail("Email"))
                    If IsValidEmail(umail) = True Then
                        ErrMsg = PE_Mail.Send(umail, rsMail("UserName"), "�ʼ�����", " " & Content & " ", SiteName, WebmasterEmail, 3)
                        If ErrMsg = "" Then
                            i = i + 1
                            Response.Write "<li>�ɹ����û�" & rsMail("UserName") & "��" & umail & "�������ʼ���</li>"
                        Else
                            j = j + 1
                            Response.Write "<li><font color='red'>���û�" & rsMail("UserName") & "��" & umail & "�������ʼ�ʧ�ܣ�</font></li>"
                        End If
                        Response.Flush
                    Else
                        k = k + 1
                    End If
                    rsMail.Close
                    Set rsMail = Nothing
                End If
            Next
            Response.Write "<li>���γɹ������ʼ���" & i & "��</li>"
            If j > 0 Then Response.Write "<li>�����ʼ�ʧ�ܣ�" & j & "��<li>"
            If k > 0 Then Response.Write "<li>δ�����ʼ���" & j & "�⣨�ʼ���ַ����<li>"
            Response.Write "<br><br><a href='Admin_Mail.asp'><<  �����ʼ����Ĺ���</a>"
            Set PE_Mail = Nothing
            rsu.Close
            Set rsu = Nothing
        End If
    End If
End Sub

Sub PreviewMail()
    Response.Write "" & Content & " "
End Sub

Sub UserList()
    Dim iChannelID, rsChannel
    iChannelID = PE_Clng(Trim(request("iChannelID")))
    strFileName = "Admin_Mail.asp?action=user&iChannelID=" & iChannelID
    If iChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���͵�Ƶ��ID</li>"
        Exit Sub
    Else
        iChannelID = PE_Clng(iChannelID)
    End If
    Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & iChannelID)
    If rsChannel.bof And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����Ƶ����</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Sub
    End If
    Dim usql, rsu
    usql = "select * from PE_MailChannel  where ChannelID =" & iChannelID
    Set rsu = Server.CreateObject("adodb.recordset")
    rsu.Open usql, Conn, 1, 1
    If rsu.bof And rsu.EOF Then
        Call WriteErrMsg("<li>��Ƶ����û�ж����ߣ�</li>", "Admin_Mail.asp")
        rsu.Close
        Set rsu = Nothing
        Exit Sub
    End If
    totalPut = UBound(Split(rsu("UserID"), ",")) + 1
    If (SearchType = 1 Or SearchType = 2) And totalPut > 100 Then
        totalPut = 100
    End If
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
    Dim rsUserList, sqlUserList
    If rsu("UserID") = "" Or IsNull(rsu("UserID")) Then
        Call WriteErrMsg("<li>��Ƶ����û�ж����ߣ�</li>", "Admin_Mail.asp")
        rsu.Close
        Set rsu = Nothing
        Exit Sub
    Else
        sqlUserList = "Select top " & MaxPerPage & " * From PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID where U.UserID in (" & rsu("UserID") & ") "
    End If
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'><tr class='title' height='22' align='center'><td>������ " & rsChannel("ChannelName") & " Ƶ���Ļ�Ա�б�</td></tr></table>"

    Response.Write "<table width='100%'    border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    ' Response.Write "  <form name='myform' method='Post' action='Admin_User.asp'>"
    Response.Write "      <td >"
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' height='22' align='center'>"
    Response.Write "          <td width='70'> �û���</td>"
    Response.Write "          <td>��Ա����</td>"
    Response.Write "          <td>������Ա��</td>"
    Response.Write "          <td width='60'><a href='" & strFileName & "&MaxPerPage=" & MaxPerPage & "&OrderType=Balance'>�ʽ����<a></td>"
    Response.Write "          <td width='60'><a href='" & strFileName & "&MaxPerPage=" & MaxPerPage & "&OrderType=Point'>����" & PointName & "��</a></td>"
    Response.Write "          <td width='60'>ʣ������</td>"
    Response.Write "          <td width='60'><a href='" & strFileName & "&MaxPerPage=" & MaxPerPage & "&OrderType=UserExp'>���û���</a></td>"
    Response.Write "          <td width='120'>����¼IP<br>����¼ʱ��</td>"
    Response.Write "          <td width='40'>��¼<br>����</td>"
    Response.Write "          <td width='40'>״̬</td>"
    Response.Write "        </tr>"

    If CurrentPage > 1 Then
        sqlUserList = sqlUserList & " and U.UserID < (select min(UserID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.UserID from PE_User U where U.UserID in (" & rsu("UserID") & ")  order by U.UserID desc)) "
    End If
    sqlUserList = sqlUserList & "order by U.UserID desc"
    Set rsUserList = Server.CreateObject("Adodb.RecordSet")
    rsUserList.Open sqlUserList, Conn, 1, 1
    If rsUserList.bof And rsUserList.EOF Then
        Response.Write "<tr><td colspan='20' height='50' align='center'>���ҵ� <font color=red>0</font> ����Ա</td></tr>"
    Else
        If (SearchType = 1 Or SearchType = 2 Or SearchType = 3 Or SearchType = 4) And CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsUserList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim UserNum
        UserNum = 0
        Do While Not rsUserList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" align=center>"
            Response.Write "        <td><a href='Admin_User.asp?Action=Show&UserID=" & rsUserList("UserID") & "'>" & rsUserList("UserName") & "</a></td>"
            Response.Write "        <td>"
            If PE_Clng(rsUserList("UserType")) > 4 Then
                Response.Write arrUserType(0)
            Else
                Response.Write arrUserType(PE_Clng(rsUserList("UserType")))
            End If
            Response.Write "        </td>"
            Response.Write "        <td>" & rsUserList("GroupName") & "</td>"
            Response.Write "        <td align='right'>" & FormatNumber(PE_CDbl(rsUserList("Balance")), 2, vbTrue, vbFalse, vbTrue) & "</td>"
            Response.Write "        <td>"
            If rsUserList("UserPoint") <= 0 Then
                Response.Write "<font color=red>" & rsUserList("UserPoint") & "</font> " & PointUnit & ""
            Else
                If rsUserList("UserPoint") <= 10 Then
                    Response.Write "<font color=blue>" & rsUserList("UserPoint") & "</font> " & PointUnit & ""
                Else
                    Response.Write rsUserList("UserPoint") & " " & PointUnit & ""
                End If
            End If
            Response.Write "</td>"
            Response.Write "<td>"
            If rsUserList("ValidNum") = -1 Then
                Response.Write "������"
            Else
                ValidDays = ChkValidDays(rsUserList("ValidNum"), rsUserList("ValidUnit"), rsUserList("BeginTime"))
                If ValidDays <= 0 Then
                    Response.Write "<font color='red'>" & ValidDays & "</font> ��"
                Else
                    Response.Write ValidDays & " ��"
                End If
            End If
            Response.Write "        </td>"
            Response.Write "        <td>" & PE_Clng(rsUserList("UserExp")) & "��</td>"
            Response.Write "        <td>" & rsUserList("LastLoginIP") & "<br>" & rsUserList("LastLoginTime") & "</td>"
            Response.Write "        <td>"
            If rsUserList("LoginTimes") <> "" Then
                Response.Write rsUserList("LoginTimes")
            Else
                Response.Write "0"
            End If
            Response.Write "        </td>"
            Response.Write "        <td>"
            If rsUserList("IsLocked") = True Then
                Response.Write "<font color=red>������</font>"
            Else
                Response.Write "����"
            End If
            Response.Write "        </td>"
            Response.Write "      </tr>"

            UserNum = UserNum + 1
            If UserNum >= MaxPerPage Then Exit Do
            rsUserList.MoveNext
        Loop
    End If
    rsUserList.Close
    Set rsUserList = Nothing
    Response.Write "<br>"
    rsu.MoveNext
    rsu.Close
    Set rsu = Nothing
    Response.Write "      </table>"
    Response.Write "      </td>"
    'Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table><br>"
    Response.Write " <table width='100%'><tr><td> <a href='Admin_Mail.asp'>>>�����ʼ����Ĺ���</a></td>"
    If totalPut > 0 Then
        Response.Write "<td align=center>"
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����Ա", True)
        Response.Write "</td>"
    End If
    Response.Write "</tr></table>"
End Sub

Function Content()
    Dim iChannelID, rsChannel
    iChannelID = PE_Clng(Trim(request("iChannelID")))
    If iChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���͵�Ƶ��ID</li>"
        Exit Function
    Else
        iChannelID = PE_Clng(iChannelID)
    End If
    Set rsChannel = Conn.Execute("select * from PE_MailChannel M inner join PE_Channel C on M.ChannelID=C.ChannelID where M.ChannelID=" & iChannelID)
    If rsChannel.bof And rsChannel.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����Ƶ����</li>"

        rsChannel.Close
        Set rsChannel = Nothing
        Exit Function
    End If
    If rsChannel("IsUse") = PE_CBool(PE_False) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Ƶ��û�п����ʼ����Ĺ��ܣ�</li>"
        rsChannel.Close
        Set rsChannel = Nothing
        Exit Function
    End If
    Dim ArrClass
    ArrClass = PE_replace(rsChannel("arrClass"), "|", ",")
	If IsValidID (ArrClass) = False then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ĿID�д��������Ƶ�������ã���ͬ��ĿID֮����|������</li>"
        Exit Function		
	End If 
    Dim sql, rsArticle, strcontent, CountNum
    sql = "select top " & rsChannel("SendNum") & " * from PE_Article where Status=3 and Deleted=" & PE_False & "   and  ChannelID = " & iChannelID
    Dim tempSql
    If IsNull(rsChannel("arrClass")) Or rsChannel("arrClass") = "" Or rsChannel("arrClass") = "0" Then
    Else
        sql = sql & " and ClassID In(" & ArrClass & ")"
    End If
    sql = sql & " order by ArticleID desc "
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sql, Conn, 1, 1
    strcontent = strcontent & "<table width='580' height='116' border='0' align='center' cellpadding='0' cellspacing='0' style='border:0px;'>"
    strcontent = strcontent & "  <tr valign=middle>"
    strcontent = strcontent & "    <td align=center style='line-height: 30px; font-size:15pt; color: #ffffff; background-color:#e15e27;'>" & SiteName & "�ʼ�����" & "</td>"
    strcontent = strcontent & "  </tr>"
    strcontent = strcontent & "  <tr>"
    strcontent = strcontent & "   <td height='37' style='border-right:1px solid #e15e27; border-left:1px solid #e15e27;'><table width='100%' border='0' cellpadding='6' cellspacing='0' >"
    strcontent = strcontent & "       <tr>"
    strcontent = strcontent & "        <td bgcolor='#ffeacb'><span style='font-size:14px; line-height:160%'>����,"
    strcontent = strcontent & "         ��л���ɹ����ı�վ:" & SiteName & "<br>�����ĵ�Ƶ����  <a href=" & SiteUrl & "/" & rsChannel("ChannelDir") & "><b>" & rsChannel("ChannelName") & "</b></a>   �����������ĵ������б�</span></td>"
    strcontent = strcontent & "        </tr>"
    strcontent = strcontent & "      <tr>"
    strcontent = strcontent & "        <td><table width='100%'>"
    strcontent = strcontent & " <tr bgcolor='#fgeacb'><td width='10%'><b>���</b></td><td width='50%'><b>����</b></td><td width='20%'><b>����</b></td><td width='20%'><b>����ʱ��</b></td></tr>"
    CountNum = 1
    Do While Not rsArticle.EOF
        strcontent = strcontent & " <tr><td wdith='10%'>" & CountNum & "</td><td width='50%'><a href=" & SiteUrl & "/" & GetInfoUrl(rsArticle("ArticleID"), "Article", 1) & ">" & rsArticle("Title") & "</td><td width='20%'>" & rsArticle("Inputer") & "</td><td width='20%'>" & rsArticle("UpdateTime") & "</td><tr>"
        CountNum = CountNum + 1
        rsArticle.MoveNext
    Loop
    strcontent = strcontent & "</talbe></td>"
    strcontent = strcontent & "      </tr>  "
    strcontent = strcontent & "        <tr width='100%'>"
    strcontent = strcontent & "          <td width='100%' height='80'  colspan='5'  style='border-top:1px solid #CCCCCC;border-bottom:1px solid #e15e27;'><a href=" & SiteUrl & ">������ʱ�վ�㣡</a><br>"
    strcontent = strcontent & "            </td>"
    strcontent = strcontent & "        </tr>"
    strcontent = strcontent & "  </table></td></tr>"
    strcontent = strcontent & "</table>"
    rsArticle.Close
    Set rsArticle = Nothing
    Content = strcontent
End Function

Sub SetChannel()
    Dim iChannelID, rsChannel
    Dim rsChannelList, sqlChannelList
    iChannelID = PE_Clng(Trim(request("iChannelID")))
    sqlChannelList = "select * from PE_MailChannel M inner join PE_Channel C on M.ChannelID=C.ChannelID Where M.ChannelID=" & iChannelID
    Set rsChannelList = Conn.Execute(sqlChannelList)
    If rsChannelList.bof And rsChannelList.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ��Ƶ����</li>"
        rsChannelList.Close
        Set rsChannelList = Nothing
        Exit Sub
    End If
    Response.Write "  <form name='myform' method='post'  action='Admin_Mail.asp?Action=SaveSet'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title' height='22'>"
    Response.Write "    <td align='center'><strong>Ƶ������</strong></td>"
    Response.Write "    <td align='center'><strong>�Ƿ����ö��Ĺ���</strong></td>"
    Response.Write "    <td align='center'><strong>����������</strong></td>"
    Response.Write "    <td align='center'><strong>�����ʼ����ĵ���Ŀ</strong></td>"
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td align='center'>" & rsChannelList("ChannelName") & "</td>"
    Response.Write "    <td align='center'><INPUT name=IsUse type=CheckBox "
    If rsChannelList("Isuse") = PE_CBool(PE_True) Then Response.Write "Checked"
    Response.Write ">��Ҫ�������</td>"
    Response.Write "    <td align='center'><INPUT name=SetNum type=Text maxlength='3' size='12' value='" & rsChannelList("SendNum") & "' ></td>"
    Response.Write "    <td align='center'><INPUT name=arrClass type=Text maxlength='200' size='30' value='" & rsChannelList("ArrClass") & "' ></td>"

    Response.Write "</tr>"

    Response.Write "</table>"
    Response.Write "<p align='center'><input name='Submit'  type='submit' id='Action' value='��������'><input name='iChannelID' type='hidden' id='iChannelID' value='" & iChannelID & " '> </p>"
    Response.Write "</form>"
    rsChannelList.Close
    Set rsChannelList = Nothing
End Sub

Sub SaveSet()
    Dim iChannelID, rsChannel, IsUse, SetNum, ArrClass
    iChannelID = PE_Clng(Trim(request("iChannelID")))
    IsUse = Trim(request("IsUse"))
    If IsUse <> "" Then
        IsUse = PE_True
    Else
        IsUse = PE_False
    End If
    SetNum = PE_Clng(Trim(request("SetNum")))
    ArrClass = ReplaceBadChar(Trim(request("ArrClass")))
    Dim sqlSave, rsSave
    sqlSave = "select * from PE_MailChannel where ChannelID=" & iChannelID
    Set rsSave = Server.CreateObject("Adodb.RecordSet")
    rsSave.Open sqlSave, Conn, 1, 3
    If rsSave.bof And rsSave.EOF Then
        Response.Write "��ָ��Ƶ��ID"
        Exit Sub
        rsSave.Close
        Set rsSave = Nothing
    End If
    rsSave("arrClass") = ArrClass
    rsSave("SendNum") = SetNum
    rsSave("IsUse") = IsUse
    rsSave.Update
    rsSave.Close
    Set rsSave = Nothing
    Call WriteSuccessMsg("�������óɹ�", "Admin_Mail.asp")
End Sub
%>
