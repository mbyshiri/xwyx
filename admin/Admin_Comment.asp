<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 1   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim Passed
Dim ClassID
Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview

Passed = Trim(Request("Passed"))
If Passed = "" Then
    Passed = Session("Passed")
End If
If Passed = "" Then
    Passed = "All"
End If
Session("Passed") = Passed
FileName = "Admin_Comment.asp?ChannelID=" & ChannelID
strFileName = "Admin_Comment.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&Field=" & strField & "&keyword=" & Keyword

'ҳ��ͷ��HTML����
Response.Write "<html><head><title>���۹���</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle(ChannelShortName & "���۹���", 10010)
If Action = "" Then
    Response.Write "<form name='form' method='Post' action='" & strFileName & "'><tr class='tdbg'>"
    Response.Write "      <td width='70' height='30' ><strong>����ѡ�</strong></td><td>"
    Response.Write "  <input name='Passed' type='radio' value='All' onclick='submit();'"
    If Passed = "All" Then Response.Write " checked"
    Response.Write ">����" & ChannelShortName & "����&nbsp;&nbsp;&nbsp;&nbsp;<input name='Passed' type='radio' value='False' onclick='submit();'"
    If Passed = "False" Then Response.Write " checked"
    Response.Write ">δ��˵�" & ChannelShortName & "����&nbsp;&nbsp;&nbsp;&nbsp;<input name='Passed' type='radio' value='True' onclick='submit();'"
    If Passed = "True" Then Response.Write " checked"
    Response.Write ">����˵�" & ChannelShortName & "����"

    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input name='Passed' type='radio' value='New' onclick='submit();'"
    If Passed = "New" Then Response.Write " checked"
    Response.Write ">���µ�" & ChannelShortName & "����"

    Response.Write "</td></tr></form>" & vbCrLf
End If

Response.Write "</table>" & vbCrLf

'ִ�еĲ���
Select Case Action
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "SetPassed", "CancelPassed", "Del", "DelReply"
    Call SetProperty
Case "Del2", "DelUser"
    Call DelComment2
Case "Reply"
    Call Reply
Case "SaveReply"
    Call SaveReply
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim rs, sql
    ClassID = PE_CLng(Trim(Request("ClassID")))
    If ClassID > 0 Then
        Dim tClass
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ</li>"
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
        Set tClass = Nothing
    End If
    
    sql = "select "
    Select Case ModuleType
    Case 1
        sql = "select I.Title as Title,I.IncludePic"
    Case 2
        sql = "select I.SoftName as Title"
    Case 3
        sql = "select I.PhotoName as Title"
    Case 5
        sql = "select I.ProductName as Title"
    Case 6
        sql = "select I.SupplyTitle as Title"
    End Select
    sql = sql & ",I." & ModuleName & "ID as ObjectID,C.CommentID,C.UserType,C.UserName,C.Email,C.Oicq,C.Homepage,C.Icq,C.Msn,C.IP"
    sql = sql & ",C.Content,C.WriteTime,C.ReplyName,C.ReplyContent,C.ReplyTime,C.Score,C.Passed"
    sql = sql & " from PE_Comment C Left Join " & SheetName & " I On C.InfoID=I." & ModuleName & "ID"
    sql = sql & " where C.ModuleType=" & ModuleType & " and I.ChannelID=" & ChannelID

    If Keyword <> "" Then
        Select Case strField
        Case "CommentContent"
            sql = sql & " and C.Content like '%" & Keyword & "%' "
        Case "CommentName"
            sql = sql & " and C.UserName like '%" & Keyword & "%' "
        Case "InfoID"
            sql = sql & " and I." & ModuleName & "ID = " & PE_CLng(Keyword) & ""
        Case "CommentTime"
            If IsDate(Trim(Request("keyword"))) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����Ĺؼ��ֲ�����Ч���ڣ�</li>"
                Exit Sub
            Else
                sql = sql & " and DateDiff(" & PE_DatePart_D & ",C.WriteTime,'" & Keyword & "')=0 "
            End If
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
    If Passed = "New" Then
        sql = sql & " order by C.WriteTime desc"
    Else
        sql = sql & " order by " & ModuleName & "ID desc"
    End If

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn, 1, 1
    
    Call ShowJS_Main("����")
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
    Response.Write "        <tr>"
    If strField = "InfoID" Then
        Response.Write "          <td>�����ڵ�λ�ã�&nbsp;���۹���"
        If Not (rs.BOF And rs.EOF) Then
            Response.Write "&nbsp;&gt;&gt;&nbsp;���⣺" & rs("Title") & "</td>"
        End If
    Else
        Response.Write "          <td>" & GetCommentPath() & "</td>"
    End If
    Response.Write "          <td width='150' align='right'>"
    If rs.BOF And rs.EOF Then
        Response.Write "���ҵ� 0 ƪ����</td></tr></table>"
    Else
        totalPut = rs.RecordCount
        Response.Write "���ҵ� " & totalPut & " ƪ����</td></tr></table>"
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
        iTemp = 1
        PrevID = rs("ObjectID")

        If Passed = "New" Then
            Response.Write "      <table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
            Response.Write "        <tr class='title'>"
            Response.Write "          <td width='80%' height='22'>"
            Response.Write "<font color='#000000'>����" & ChannelShortName & "����</font>"
            Response.Write "          </td>"
            Response.Write "        </tr>"

            Response.Write "        <tr>"
            Response.Write "          <td colspan='2'>"
            Response.Write "            <table border='0' cellspacing='1' width='100%' cellpadding='0' style='word-break:break-all'>"
        End If

        Do While Not rs.EOF
            If Passed <> "New" Then
                If rs("ObjectID") <> PrevID Then Response.Write "</table></td></tr></table><br>"
                If CommentNum = 0 Or rs("ObjectID") <> PrevID Then
                    iTemp = 1
                    Response.Write "      <table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
                    Response.Write "        <tr class='title'>"
                    Response.Write "          <td width='80%' height='22'>"
                    Response.Write "<a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Show&" & ModuleName & "ID=" & rs("ObjectID") & "'>" & rs("Title") & "</a>"
                    Response.Write "          </td>"
                    Response.Write "          <td width='20%' align='right'><a href='" & strFileName & "&Action=Del2&InfoID=" & rs("ObjectID") & "'>ɾ����" & ChannelShortName & "�µ���������</a></td>"
                    Response.Write "        </tr>"
                    Response.Write "        <tr>"
                    Response.Write "          <td colspan='2'>"
                    Response.Write "            <table border='0' cellspacing='1' width='100%' cellpadding='0' style='word-break:break-all'>"
                End If
            End If
                    
            Response.Write "              <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "                <td width='30' align='center'>"
            Response.Write "                  <input name='CommentID' type='checkbox' onclick=""unselectall()"" id='CommentID' value='" & CStr(rs("CommentID")) & "'>"
            Response.Write "                </td>"
            Response.Write "                <td width='20' align='center'>" & iTemp & "</td>"
            Response.Write "                <td align='left'>"
            If rs("UserType") = 1 Then
                Response.Write "[��Ա] "
            Else
                Response.Write "[�ο�] "
            End If
            If rs("UserType") = 1 Then
                Response.Write "<a href='Admin_User.asp?UserName=" & rs("UserName") & "' target='_blank'>" & rs("UserName") & "</a>"
            Else
                Response.Write "<span title='" & nohtml("������" & rs("UserName") & vbCrLf & "���䣺" & rs("Email") & vbCrLf & "Oicq��" & rs("Oicq") & vbCrLf & " Icq��" & rs("Icq") & vbCrLf & " Msn��" & rs("Msn") & vbCrLf & " I P��" & rs("IP") & vbCrLf & "��ҳ��" & rs("Homepage")) & "' style='cursor:hand'>" & rs("UserName") & "</span>"
            End If
            Response.Write " �� " & rs("WriteTime") & " ���������������ݣ�ͬʱ���֣�" & rs("Score") & "��<br>"
            Response.Write rs("Content")
            Response.Write "                </td><td width='30' align='center'>"
            If rs("Passed") = True Then
                Response.Write "��"
            Else
                Response.Write "<font color='red'>��</font>"
            End If
            Response.Write "</td>"
            Response.Write "                <td width='150' align='center'>"
            If rs("ReplyContent") <> "" Then
                Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            Else
                Response.Write "<a href='" & strFileName & "&Action=Reply&CommentID=" & rs("CommentID") & "'>�ظ�</a>&nbsp;"
            End If
            Response.Write "<a href='" & strFileName & "&Action=Modify&CommentID=" & rs("CommentID") & "'>�޸�</a>&nbsp;"
            Response.Write "<a href='" & strFileName & "&Action=Del&CommentID=" & rs("CommentID") & "' onclick=""return confirm('ȷ��Ҫɾ����������');"">ɾ��</a>&nbsp;"
            If rs("Passed") = True Then
                Response.Write "<a href='" & strFileName & "&Action=CancelPassed&CommentID=" & rs("CommentID") & "'>ȡ��ͨ��</a>"
            Else
                Response.Write "<a href='" & strFileName & "&Action=SetPassed&CommentID=" & rs("CommentID") & "'>ͨ�����</a>"
            End If
            Response.Write "                </td>"
            Response.Write "              </tr>"
            If rs("ReplyContent") <> "" Then
                Response.Write "            <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
                Response.Write "              <td align='center'>&nbsp;</td>"
                Response.Write "              <td align='center'>&nbsp;</td>"
                Response.Write "              <td colspan='2' align='left'>[����Ա] " & rs("ReplyName") & " �� " & rs("ReplyTime") & " �ظ���<br>"
                Response.Write rs("ReplyContent")
                Response.Write "</td><td align='center'>"
                Response.Write "<a href='" & strFileName & "&Action=Reply&CommentID=" & rs("CommentID") & "'>�޸�</a>&nbsp;"
                Response.Write "<a href='" & strFileName & "&Action=DelReply&CommentID=" & rs("CommentID") & "' onclick=""return confirm('ȷ��Ҫɾ�������۵Ĺ���Ա�ظ���');"">ɾ��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
                Response.Write "</td>"
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
        Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "        <tr>"
        Response.Write "          <td width='200' height='30'>"
        Response.Write "            <input name='chkAll' type='checkbox' id='chkAll' onclick=CheckAll(this.form) value='checkbox'>"
        Response.Write "            ѡ�б�ҳ��ʾ����������"
        Response.Write "          </td>"
        Response.Write "          <td>"
        Response.Write "<input name='submit' type='submit' value='ɾ��ѡ��������'>&nbsp;&nbsp;"
        If Keyword <> "" And strField = "CommentName" Then
            Response.Write "<input name='submitUser' type='submit' value='ɾ��" & Keyword & "ȫ������'>&nbsp;&nbsp;"
            Response.Write "<input name='CommentUser' type='hidden' id='CommentUser' value='" & Keyword & "'>"
            Response.Write "<input name='Action' type='hidden' id='Action' value='DelUser'>"
        Else
            Response.Write "<input name='Action' type='hidden' id='Action' value='Del'>"
        End If
        Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='SetPassed'"" value='���ͨ��ѡ��������'>&nbsp;&nbsp;"
        Response.Write "<input name='submit2' type='submit' id='submit2' onClick=""document.myform.Action.value='CancelPassed'"" value='ȡ�����ѡ��������'>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "      </table>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "������", True)
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>����������</strong></td>"
    Response.Write "   <td>" & GetCommentSearch() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub Modify()
    Dim rs, sql
    Dim CommentID
    CommentID = PE_CLng(Trim(Request("CommentID")))
    If CommentID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
        Exit Sub
    End If
    sql = "Select * from PE_Comment where CommentID=" & CommentID
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1
    If rs.BOF Or rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ�������ۣ�</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' style='word-break:break-all;Width:fixed'>"
    Response.Write "  <form name='myform' method='post' action='" & strFileName & "'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td height='22' colspan='4'> <strong>�� �� �� �� </strong>&nbsp;&nbsp;"
    If rs("UserType") = 1 Then
        Response.Write "����Աģʽ��"
    Else
        Response.Write "���ο�ģʽ��"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    If rs("UserType") = 0 Then
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>������������</td>"
        Response.Write "    <td class='tdbg' width='200'>"
        Response.Write "      <input name='UserName' type='text' id='UserName' maxlength='16' value='" & rs("UserName") & "'>"
        Response.Write "    </td>"
        Response.Write "    <td class='tdbg' align='right' width='101'>������Oicq��</td>"
        Response.Write "    <td class='tdbg' width='475'>"
        Response.Write "      <input name='Oicq' type='text' id='UserName' maxlength='15' value='" & rs("Oicq") & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>�������Ա�</td>"
        Response.Write "    <td class='tdbg' width='200'>"
        Response.Write "      <input type='radio' name='Sex' value='1' checked style='BORDER:0px;'>"
        Response.Write "      ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "      <input type='radio' name='Sex' value='0' style='BORDER:0px;'>"
        Response.Write "      Ů </td>"
        Response.Write "    <td class='tdbg' align='right' width='101'>������ Icq��</td>"
        Response.Write "    <td class='tdbg' width='475'>"
        Response.Write "      <input name='Icq' type='text' id='UserName'  maxlength='15' value='" & rs("Icq") & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>������Email��</td>"
        Response.Write "    <td class='tdbg' width='200'>"
        Response.Write "      <input name='Email' type='text' id='UserName' maxlength='40' value='" & rs("Email") & "'>"
        Response.Write "    </td>"
        Response.Write "    <td class='tdbg' align='right' width='101'>������ Msn��</td>"
        Response.Write "    <td class='tdbg' width='475'>"
        Response.Write "      <input name='Msn' type='text' id='UserName' maxlength='40' value='" & rs("Msn") & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>����ʱ�䣺</td>"
        Response.Write "    <td class='tdbg' width='200'>"
        Response.Write "      <input name='WriteTime' type='text' id='WriteTime' value='" & rs("WriteTime") & "'>"
        Response.Write "    </td>"
        Response.Write "    <td class='tdbg' align='right' width='101'>������IP��</td>"
        Response.Write "    <td class='tdbg' width='475'>"
        Response.Write "      <input name='IP' type='text' id='IP'  maxlength='15' value='" & rs("IP") & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>��������ҳ��</td>"
        Response.Write "    <td class='tdbg' colspan='3'>"
        Response.Write "      <input name='Homepage' type='text' id='UserName' maxlength='60' value='"
        If rs("Homepage") = "" Then
            Response.Write "http://"
        Else
            Response.Write rs("Homepage")
        End If
        Response.Write "' size='66'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
    Else
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>������������</td>"
        Response.Write "    <td class='tdbg' colspan='3'>"
        Response.Write "      <input name='ShowUserName' type='text' id='UserName' value='" & rs("UserName") & "' disabled>"
        Response.Write "      <input name='UserName' type='hidden' id='UserName' value='" & rs("UserName") & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' align='right' class='tdbg'>����ʱ�䣺</td>"
        Response.Write "    <td class='tdbg' width='200'>"
        Response.Write "      <input name='WriteTime' type='text' id='WriteTime' value='" & rs("WriteTime") & "'>"
        Response.Write "    </td>"
        Response.Write "    <td class='tdbg' align='right' width='101'>������IP��</td>"
        Response.Write "    <td class='tdbg' width='475'>"
        Response.Write "      <input name='IP' type='text' id='IP' maxlength='15' value='" & rs("IP") & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
    End If
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>�� �֣�</td>"
    Response.Write "    <td class='tdbg' colspan='3'>"
    Response.Write "      <input type='radio' name='Score' value='1' "
    If rs("Score") = 1 Then Response.Write " checked"
    Response.Write "      >1��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='2' "
    If rs("Score") = 2 Then Response.Write " checked"
    Response.Write "      >2��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='3' "
    If rs("Score") = 3 Then Response.Write " checked"
    Response.Write "      >3��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='4' "
    If rs("Score") = 4 Then Response.Write " checked"
    Response.Write "      >4��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <input type='radio' name='Score' value='5' "
    If rs("Score") = 5 Then Response.Write " checked"
    Response.Write "      >5�� </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>�������ݣ�</td>"
    Response.Write "    <td class='tdbg' colspan='3'>"
    Response.Write "      <textarea name='Content' cols='56' rows='8' id='Content'>" & PE_ConvertBR(rs("Content")) & "</textarea>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td height='30' colspan='4' class='tdbg'>"
    Response.Write "      <input name='ComeUrl' type='hidden' id='ComeUrl' value='" & ComeUrl & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "      <input name='CommentID' type='hidden' id='CommentID' value='" & rs("CommentID") & "'>"
    Response.Write "      <input name='UserType' type='hidden' id='UserType' value='" & rs("UserType") & "'>"
    Response.Write "      <input  type='submit' name='Submit' value=' �����޸Ľ�� '>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub Reply()
    Dim rs, sql
    Dim CommentID
    CommentID = PE_CLng(Trim(Request("CommentID")))
    If CommentID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
        Exit Sub
    End If
    
    Select Case ModuleType
    Case 1
        sql = "select I.Title as Title"
    Case 2
        sql = "select I.SoftName as Title"
    Case 3
        sql = "select I.PhotoName as Title"
    Case 5
        sql = "select I.ProductName as Title"
    Case 6
        sql = "Select I.SupplyTitle as Title"
    End Select
    sql = sql & ",C.CommentID,C.UserName,C.IP, C.Content,C.WriteTime,C.ReplyContent"
    sql = sql & " from PE_Comment C Left Join " & SheetName & " I On C.InfoID=I." & ModuleName & "ID where C.CommentID=" & CommentID
    
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1
    If rs.BOF Or rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ�������ۣ�</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' style='word-break:break-all;Width:fixed'>"
    Response.Write "  <form method='post' action='" & strFileName & "' name='myform'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td height='22' colspan='2'> <strong>�� �� �� ��</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>����" & ChannelShortName & "���⣺</td>"
    Response.Write "    <td class='tdbg'>" & rs("Title") & "</td>"

    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>�������û�����</td>"
    Response.Write "    <td class='tdbg'>" & rs("UserName") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' align='right' class='tdbg'>�������ݣ�</td>"
    Response.Write "    <td class='tdbg'>" & rs("Content") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td align='right' class='tdbg'>�ظ����ݣ�</td>"
    Response.Write "    <td class='tdbg'><textarea name='ReplyContent' cols='50' rows='6' id='ReplyContent'>" & PE_ConvertBR(rs("ReplyContent")) & "</textarea></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td height='30' colspan='2' class='tdbg'><input name='ComeUrl' type='hidden' id='ComeUrl' value='" & ComeUrl & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='SaveReply'>"
    Response.Write "      <input name='CommentID' type='hidden' id='CommentID' value='" & rs("CommentID") & "'>"
    Response.Write "      <input  type='submit' name='Submit' value=' �� �� '> </td>"
    Response.Write "  </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub SetProperty()
    Dim CommentID
    Dim sqlProperty, rsProperty
    Dim ShowType, MoveChannelID
    CommentID = Trim(Request("CommentID"))
    If IsValidID(CommentID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If InStr(CommentID, ",") > 0 Then
        sqlProperty = "select * from PE_Comment where CommentID in (" & CommentID & ")"
    Else
        sqlProperty = "select * from PE_Comment where CommentID=" & CommentID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetPassed"
            rsProperty("Passed") = True
        Case "CancelPassed"
            rsProperty("Passed") = False
        Case "DelReply"
            rsProperty("ReplyContent") = ""
        Case "Del"
            rsProperty.Delete
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing
    
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub DelComment2()
    Dim InfoID, CommentUser
    InfoID = Trim(Request("InfoID"))
    CommentUser = Trim(Request("CommentUser"))
    If CommentUser <> "" Then
        CommentUser = ReplaceBadChar(CommentUser)
    End If
    If CommentUser = "" Then
        If InfoID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
            Exit Sub
        Else
            InfoID = PE_CLng(InfoID)
        End If
        Conn.Execute "delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID=" & InfoID
    Else
        Conn.Execute "delete from PE_Comment where ModuleType=" & ModuleType & " and UserName like '%" & CommentUser & "%' "
    End If
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub SaveModify()
    Dim rsComment, sql, CommentID
    Dim CommentUserType, CommentUserName, CommentUserSex, CommentUserEmail, CommentUserOicq
    Dim CommentUserIcq, CommentUserMsn, CommentUserHomepage, CommentUserScore, CommentUserContent
    Dim CommentUserIP, CommentWritetime
    CommentID = PE_CLng(Trim(Request("CommentID")))
    If CommentID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
        Exit Sub
    End If
    CommentUserName = Trim(Request("UserName"))
    If CommentUserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����������</li>"
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
        ErrMsg = ErrMsg & "<li>�����������������ݡ�����ʱ�䡢������IP����Ϣ</li>"
    End If
    CommentUserContent = PE_HTMLEncode(CommentUserContent)

    If FoundErr = True Then
        Exit Sub
    End If

    sql = "Select * from PE_Comment where CommentID=" & CommentID
    Set rsComment = Server.CreateObject("Adodb.RecordSet")
    rsComment.Open sql, Conn, 1, 3
    If rsComment.BOF Or rsComment.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ�������ۣ�</li>"
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
        rsComment("Content") = CommentUserContent
        rsComment.Update
    End If
    rsComment.Close
    Set rsComment = Nothing
    Call CloseConn
    Response.Redirect strFileName
End Sub

Sub SaveReply()
    Dim rs, sql
    Dim CommentID, ReplyName, ReplyContent, ReplyTime
    CommentID = PE_CLng(Trim(Request("CommentID")))
    ReplyContent = Trim(Request("ReplyContent"))
    If CommentID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ������ID</li>"
        Exit Sub
    End If
    If ReplyContent = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ظ�����</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    sql = "Select * from PE_Comment where CommentID=" & CommentID
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3
    If rs.BOF Or rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ�������ۣ�</li>"
    Else
        rs("ReplyName") = AdminName
        rs("ReplyTime") = Now()
        rs("ReplyContent") = PE_HTMLEncode(ReplyContent)
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
    Call CloseConn
    Response.Redirect strFileName
End Sub

Function GetCommentPath()
    Dim strPath
    strPath = "�����ڵ�λ�ã�&nbsp;���۹���&nbsp;&gt;&gt;&nbsp;"
    If ClassID > 0 Then
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Server.CreateObject("adodb.recordset")
            rsPath.Open sqlPath, Conn, 1, 1
            Do While Not rsPath.EOF
                strPath = strPath & "<a href='" & FileName & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        strPath = strPath & "<a href='" & FileName & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If Keyword = "" Then
        If Passed = "New" Then
            strPath = strPath & "����" & ChannelShortName & "����"
        Else
            strPath = strPath & "��������"
        End If
    Else
        Select Case strField
            Case "CommentContent"
                strPath = strPath & "���������к��� <font color=red>" & Keyword & "</font> ������"
            Case "CommentName"
                strPath = strPath & "�������к��� <font color=red>" & Keyword & "</font> ������"
            Case Else
                strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> ������"
            End Select

        End If
    GetCommentPath = strPath
End Function


Function GetCommentSearch()
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='Admin_Comment.asp'>"
    strForm = strForm & "<tr><td height='28' align='center'>"
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='CommentContent' selected>��������</option>"
    strForm = strForm & "<option value='CommentTime'>����ʱ��</option>"
    strForm = strForm & "<option value='CommentName'>������</option>"
    strForm = strForm & "</select>"
    strForm = strForm & "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    strForm = strForm & "<input type='submit' name='Submit'  value='����'>"
    strForm = strForm & "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    strForm = strForm & "</td></tr></form></table>"
    GetCommentSearch = strForm
End Function
%>
