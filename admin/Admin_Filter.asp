<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Collection"   '����Ȩ��

Dim rs, sql, rsItem, strsql, i 'ͨ�ñ���

strFileName = "Admin_Filter.asp?Action=" & Action

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>����</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf
Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
Call ShowPageTitle(" �� �� �� �� �� �� ", 10053)
Response.Write "  <tr class=""tdbg""> " & vbCrLf
Response.Write "    <td width=""70"" height=""30""><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td height=""30""><a href=Admin_Filter.asp?Action=Main>������ҳ</a> | <a href=""Admin_Filter.asp?Action=FilterAdd"">�������Ŀ</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>"

Select Case Action
Case "FilterAdd"
    Call FilterAdd              '������Ŀ���
Case "FilterModify"
    Call FilterModify           '������Ŀ�޸�
Case "SaveFileter"
    Call SaveFileter            '���������Ŀ
Case "Del"
    Call Del                    'ɾ��������Ŀ
Case "DelAll"
    Call DelAll                 '��չ�����Ŀ
Case "SetFlag"
    Call SetFlag                '�Ƿ�����
Case Else
    Call main                   '������Ŀ����
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


'=================================================
'��������main
'��  �ã��ɼ�������Ŀ�༭
'=================================================
Sub main()
    Dim FilterID, MaxPerPage
            
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))

    If MaxPerPage <= 0 Then MaxPerPage = 20

    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If

    strFileName = "Admin_Filter.asp?Action=main"
    
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function unselectall(thisform){" & vbCrLf
    Response.Write "    if(thisform.chkAll.checked){" & vbCrLf
    Response.Write "        thisform.chkAll.checked = thisform.chkAll.checked&0;" & vbCrLf
    Response.Write "    }   " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(thisform){" & vbCrLf
    Response.Write "    for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "    var e = thisform.elements[i];" & vbCrLf
    Response.Write "    if (e.Name != ""chkAll""&&e.disabled!=true)" & vbCrLf
    Response.Write "        e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form name=""form1"" method=""POST"" action=""Admin_Filter.asp"">" & vbCrLf
    Response.Write "<table class=""border"" border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & vbCrLf
    Response.Write "   <tr class=""title"" style=""padding: 0px 2px;"">" & vbCrLf
    Response.Write "     <td width=""30"" height=""22"" align=""center""><strong>ѡ��</strong></td>" & vbCrLf
    Response.Write "     <td width=""250"" height=""22"" align=""center""><strong>�����ɼ���Ŀ</strong></td>" & vbCrLf
    Response.Write "     <td width=""120"" align=""center""><strong>��������</strong></td>" & vbCrLf
    Response.Write "     <td width=""80"" align=""center""><strong>��������</strong></td>" & vbCrLf
    Response.Write "     <td width=""80"" height=""22"" align=""center""><strong>��������</strong></td>" & vbCrLf
    Response.Write "     <td width=""40"" align=""center""><strong>״̬</strong></td>" & vbCrLf
    Response.Write "     <td width=""80"" height=""22"" align=""center""><strong>����</strong></td>" & vbCrLf
    Response.Write "   </tr>" & vbCrLf
    
    sql = "SELECT F.*, I.ItemName FROM PE_Filters F LEFT JOIN PE_Item I ON F.ItemID = I.ItemID ORDER BY F.FilterID DESC"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class=""tdbg""><td colspan='7' height='50' align='center'>ϵͳ�����޹�����Ŀ��</td></tr></table>"
    Else
        totalPut = rs.RecordCount
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

        Dim VisitorNum
        VisitorNum = 0

        Do While Not rs.EOF
            Response.Write "    <tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""padding: 0px 2px;""> " & vbCrLf
            Response.Write "      <td width=""30"" align=""center"">" & vbCrLf
            Response.Write "        <input type=""checkbox"" value=" & rs("FilterID") & " name=""FilterID"" onclick=""unselectall(this.form)"" >" & vbCrLf
            Response.Write "      </td>" & vbCrLf
            Response.Write "      <td width=""200"" align=""center"">"

            If rs("ItemID") = -1 Then
                Response.Write " ������Ŀ "
            ElseIf rs("ItemID") = 0 Then
                Response.Write " û��ָ����Ŀ "
            Else
                Response.Write rs("ItemName")
            End If

            Response.Write "      </td>" & vbCrLf
            Response.Write "     <td width=""80"" align=""center"">" & rs("FilterName") & "</td>" & vbCrLf
            Response.Write "      <td width=""80"" align=""center"">" & vbCrLf

            If rs("FilterObject") = 1 Then
                Response.Write "�������"
            ElseIf rs("FilterObject") = 2 Then
                Response.Write "���Ĺ���"
            Else
                Response.Write "��ѡ��"
            End If

            Response.Write "      </td>" & vbCrLf
            Response.Write "      <td width=""80"" align=""center"">"

            If rs("FilterType") = 1 Then
                Response.Write "���滻"
            ElseIf rs("FilterType") = 2 Then
                Response.Write "�߼�����"
            Else
                Response.Write "��ѡ��"
            End If

            Response.Write "      </td>" & vbCrLf
            Response.Write "     <td width=""40"" align=""center"">"

            If rs("Flag") = True Then
                Response.Write "<b>��</b>"
            Else
                Response.Write "<FONT color='red'><b>��</b></FONT>"
            End If

            Response.Write "</td>" & vbCrLf
            Response.Write "      <td width=""100"" align=""center"">" & vbCrLf

            If rs("Flag") = True Then
                Response.Write "      <a href=Admin_Filter.asp?Action=SetFlag&FilterFlag=0&FilterID=" & rs("FilterID") & ">����</a>&nbsp;" & vbCrLf
            Else
                Response.Write "      <a href=Admin_Filter.asp?Action=SetFlag&FilterFlag=1&FilterID=" & rs("FilterID") & ">����</a>&nbsp;" & vbCrLf
            End If

            Response.Write "      <a href=Admin_Filter.asp?Action=FilterModify&FilterID=" & rs("FilterID") & ">�޸�</a>&nbsp;" & vbCrLf
            Response.Write "      <a href=Admin_Filter.asp?Action=Del&FilterID=" & rs("FilterID") & " onclick='return confirm(""ȷ��Ҫɾ������Ŀ��"");'>ɾ��</a>" & vbCrLf
            Response.Write "      </td>" & vbCrLf
            Response.Write "    </tr> " & vbCrLf
                
            VisitorNum = VisitorNum + 1

            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>  " & vbCrLf

        Response.Write "<table border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0""><tr><td height=""30"">" & vbCrLf
        Response.Write "<input name=""Action"" type=""hidden""  value=""Del"">   " & vbCrLf
        Response.Write "<input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" >ѡ��������Ŀ" & vbCrLf
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "<input type=""submit"" value="" ����ɾ�� "" name=""Del"" onClick=""document.form1.Action.value='Del';return confirm('��ȷ��Ҫ����ɾ����Щ������Ŀ��');"" >&nbsp;&nbsp;" & vbCrLf
        Response.Write "<input type=""submit"" value=""������м�¼"" name=""DelAll"" onclick=""document.form1.Action.value='DelAll';return confirm('�����Ҫȷ��Ҫ������й�����Ŀ��');"" >&nbsp;&nbsp;" & vbCrLf
        Response.Write "</td></tr></table>  " & vbCrLf
        Response.Write "</form>" & vbCrLf

        If totalPut > 0 Then
            Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��������Ŀ��¼", True)
        End If
    End If

    rs.Close
    Set rs = Nothing
End Sub


'=================================================
'��������FilterAdd
'��  �ã��ɼ�������Ŀ���
'=================================================
Sub FilterAdd()
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function showset(thisform){" & vbCrLf
    Response.Write "    if(thisform.FilterType.selectedIndex==1){" & vbCrLf
    Response.Write "        FilterType1.style.display = ""none"";" & vbCrLf
    Response.Write "        FilterType2.style.display = """";" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        FilterType1.style.display = """";" & vbCrLf
    Response.Write "        FilterType2.style.display = ""none"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form method=""post"" action=""Admin_Filter.asp"" name=""form1"">" & vbCrLf
    Response.Write " <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
    Response.Write "   <tr> " & vbCrLf
    Response.Write "    <td> " & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >" & vbCrLf
    Response.Write "      <tr> " & vbCrLf
    Response.Write "        <td height=""22"" colspan=""2"" class=""title"" align=""center""><strong>�� �� �� �� ��</strong></td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class=""tdbg"">" & vbCrLf
    Response.Write "        <td width=""120"" class='tdbg5' align='right'>�������ƣ�</td>" & vbCrLf
    Response.Write "        <td class=""tdbg""><input name=""FilterName"" type=""text"" id=""FilterName"" size=""25"" maxlength=""30"">" & vbCrLf
    Response.Write "        &nbsp;</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class=""tdbg"">" & vbCrLf
    Response.Write "        <td width=""120"" class='tdbg5' align='right'> �����ɼ���Ŀ��</td>" & vbCrLf
    Response.Write "        <td class=""tdbg""> "
    Call ShowItem_Option(0)
    Response.Write "          <font color=blue>&nbsp;&nbsp;������Ŀ��ִ����������Ŀ,�������β�������</font>"
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class=""tdbg""> " & vbCrLf
    Response.Write "        <td width=""120"" class='tdbg5' align='right'>���˶���</td>" & vbCrLf
    Response.Write "        <td class=""tdbg"">" & vbCrLf
    Response.Write "         <select name=""FilterObject"" id=""FilterObject"">" & vbCrLf
    Response.Write "            <option value=""1"" selected>�������</option>" & vbCrLf
    Response.Write "            <option value=""2"">���Ĺ���</option>" & vbCrLf
    Response.Write "         </select>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class=""tdbg"">" & vbCrLf
    Response.Write "        <td width=""120"" class='tdbg5' align='right'>�������ͣ�</td>" & vbCrLf
    Response.Write "        <td class=""tdbg"">" & vbCrLf
    Response.Write "         <select name=""FilterType"" id=""FilterType"" onchange=showset(this.form)>" & vbCrLf
    Response.Write "            <option value=""1"" selected >���滻</option>" & vbCrLf
    Response.Write "            <option value=""2"">�߼�����</option>" & vbCrLf
    Response.Write "         </select>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" id=""FilterType1"" style=""display:"">" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'>�������ݣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FilterContent"" cols=""49"" rows=""5""></textarea></td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  id=""FilterType2"" style=""display:none"">" & vbCrLf
    Response.Write "       <tr class=""tdbg"">" & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'>��ʼ��ǣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FisString"" cols=""49"" rows=""5""></textarea>" & vbCrLf
    Response.Write "        &nbsp;</td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'>������ǣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FioString"" cols=""49"" rows=""5""></textarea>" & vbCrLf
    Response.Write "        &nbsp;</td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >" & vbCrLf
    Response.Write "       <tr class=""tdbg"" id=""FilterRep""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'>�滻Ϊ��</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FilterRep"" cols=""49"" rows=""5""></textarea></td> " & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table> " & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >" & vbCrLf
    Response.Write "       <tr class='tdbg'>"
    Response.Write "         <td width=""120"" align='right' class=""tdbg5"">�Ƿ����ã�</td>"
    Response.Write "         <td>&nbsp;&nbsp;<input name='Flag' type='checkbox' id='Flag' value='Yes' checked></td>"
    Response.Write "       </tr>"
    Response.Write "     </table>"
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td colspan=""2"" align=""center"" class=""tdbg"" height='50'>" & vbCrLf
    Response.Write "           <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveFileter"">" & vbCrLf
    Response.Write "           <input  type=""submit"" name=""Submit"" value="" ȷ  �� "" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "           <input name=""Cancel"" type=""button"" id=""Cancel"" value="" ȡ  �� "" onClick=""window.location.href='Admin_Filter.asp?Action=main'"" >&nbsp;&nbsp;&nbsp;&nbsp; " & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "     </td>"
    Response.Write "   </tr>"
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

'=================================================
'��������FilterModify
'��  �ã��ɼ�������Ŀ�޸�
'=================================================
Sub FilterModify()
    Dim SqlItem, rsFilters, ItemID, FilterName, FilterID, FilterObject, FilterType, FilterContent, FisString, FioString, FilterRep, Flag
    FilterID = PE_Clng(Trim(Request("FilterID")))

    Set rsFilters = Server.CreateObject("adodb.recordset")
    SqlItem = "select * from PE_Filters Where FilterID=" & FilterID
    rsFilters.Open SqlItem, Conn, 1, 1
    If Not rsFilters.EOF Then
        ItemID = rsFilters("ItemID")
        FilterName = rsFilters("FilterName")
        FilterObject = rsFilters("FilterObject")
        FilterType = rsFilters("FilterType")
        FilterContent = rsFilters("FilterContent")
        FisString = rsFilters("FisString")
        FioString = rsFilters("FioString")
        FilterRep = rsFilters("FilterRep")
        Flag = rsFilters("Flag")
    Else
        Response.Write "�Ҳ�������Ŀ"
        Exit Sub
    End If

    rsFilters.Close
    Set rsFilters = Nothing
    
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function showset(thisform)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "        if(thisform.FilterType.selectedIndex==1)" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            FilterType1.style.display = ""none"";" & vbCrLf
    Response.Write "            FilterType2.style.display = """";" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            FilterType1.style.display = """";" & vbCrLf
    Response.Write "            FilterType2.style.display = ""none"";" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<form method=""post"" action=""Admin_Filter.asp"" name=""form1"">" & vbCrLf
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td> " & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  >" & vbCrLf
    Response.Write "       <tr> " & vbCrLf
    Response.Write "         <td height=""22"" colspan=""2"" class=""title""> <div align=""center""><strong>�� �� �� ��</strong></div></td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> �������ƣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><input name=""FilterName"" type=""text"" id=""FilterName"" value=" & FilterName & "  size=""25"" maxlength=""30"">&nbsp;</td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> ������Ŀ��</td>" & vbCrLf
    Response.Write "         <td class=""tdbg"">"
    Call ShowItem_Option(ItemID)
    Response.Write "            <font color=blue>&nbsp;&nbsp;������Ŀ��ִ����������Ŀ,�������β�������</font>"
    Response.Write "         </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> ���˶���</td>" & vbCrLf
    Response.Write "         <td class=""tdbg"">" & vbCrLf
    Response.Write "           <select name=""FilterObject"" id=""FilterObject"">  " & vbCrLf
    Response.Write "             <option value=""1"" "

    If FilterObject = 1 Then Response.Write "selected"
    Response.Write "             >�������</option>  " & vbCrLf
    Response.Write "             <option value=""2"" "

    If FilterObject = 2 Then Response.Write "selected"
    Response.Write "             >���Ĺ���</option>  " & vbCrLf
    Response.Write "           </select>  " & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg"">" & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> �������ͣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg"">" & vbCrLf
    Response.Write "           <select name=""FilterType"" id=""FilterType"" onchange=showset(this.form)>  " & vbCrLf
    Response.Write "             <option value=""1"" "

    If FilterType = 1 Then Response.Write "selected"
    Response.Write "             >���滻</option>" & vbCrLf
    Response.Write "             <option value=""2"" "

    If FilterType = 2 Then Response.Write "selected"
    Response.Write "             >�߼�����</option>  " & vbCrLf
    Response.Write "           </select>  " & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  id=""FilterType1"" "

    If FilterType <> 1 Then Response.Write "style='display:none'"
    Response.Write ">" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> ���ݣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FilterContent"" cols=""49"" rows=""5"">" & Server.HTMLEncode(FilterContent & "") & "</textarea></td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table> " & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" id=""FilterType2"" "

    If FilterType <> 2 Then Response.Write "style='display:none'"
    Response.Write ">" & vbCrLf
    Response.Write "       <tr class=""tdbg"">" & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> ��ʼ��ǣ�</td>  " & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FisString"" cols=""49"" rows=""5"">" & Server.HTMLEncode(FisString & "") & "</textarea>&nbsp;</td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> ������ǣ�</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FioString"" cols=""49"" rows=""5"">" & Server.HTMLEncode(FioString & "") & "</textarea>&nbsp;</td>" & vbCrLf
    Response.Write "       </tr> " & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align='right'> �滻��</td>" & vbCrLf
    Response.Write "         <td class=""tdbg""><textarea name=""FilterRep"" cols=""49"" rows=""5"" id=""FilterRep"">" & Server.HTMLEncode(FilterRep & "") & "</textarea></td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >" & vbCrLf
    Response.Write "       <tr class='tdbg' >"
    Response.Write "         <td align='right' width=""120"" class=""tdbg5"">�Ƿ����ã�</td>"
    Response.Write "         <td>&nbsp;&nbsp;<input name='Flag' type='checkbox' id='Flag' value='Yes' " & vbCrLf

    If Flag = True Then
        Response.Write "checked"
    End If

    Response.Write "          ></td>"
    Response.Write "       </tr>"
    Response.Write "     </table>"
    Response.Write "     <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td align=""center"" class=""tdbg"" height='50'>" & vbCrLf
    Response.Write "           <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveFileter"">" & vbCrLf
    Response.Write "           <input name=""FilterID"" type=""hidden"" id=""FilterID"" value=""" & FilterID & """>" & vbCrLf
    Response.Write "           <input  type=""Submit"" name=""Submit"" value="" ȷ  �� "" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "           <input name=""Cancel"" type=""button"" id=""Cancel"" value="" ��  �� "" onClick=""window.location.href='Admin_Filter.asp?Action=main'"" >&nbsp;" & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub
'=================================================
'��������SaveFileter
'��  �ã��������
'=================================================
Sub SaveFileter()
    Dim rsFilters, SqlItem
    Dim FilterName, ItemID, FilterID, FilterObject, FilterType, FilterContent, FisString, FioString, FilterRep, Flag
    
    FilterName = Trim(Request.Form("FilterName"))
    ItemID = Trim(Request.Form("ItemID"))
    FilterID = PE_CLng(Trim(Request("FilterID")))
    FilterObject = Request.Form("FilterObject")
    FilterType = Request.Form("FilterType")
    FilterContent = Request.Form("FilterContent")
    FisString = Request.Form("FisString")
    FioString = Request.Form("FioString")
    FilterRep = Request.Form("FilterRep")
    Flag = Trim(Request.Form("Flag"))
                                  
    If FilterName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������Ʋ���Ϊ��</li>"
    End If

    If ItemID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ѡ�����������Ŀ</li>"
    Else
        ItemID = CLng(ItemID)

        If ItemID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ�����������Ŀ</li>"
        End If
    End If

    If FilterObject = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ѡ����˶���</li>"
    Else
        FilterObject = PE_CLng(FilterObject)
    End If

    If FilterType = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ѡ���������</li>"
    Else
        FilterType = PE_CLng(FilterType)

        If FilterType = 1 Then
            If FilterContent = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���˵����ݲ���Ϊ��</li>"
            End If

            If Len(FilterContent) >= 50 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������ݲ��ܳ���50�ַ�,�򵥹������������໰�ȷǷ��ʻ�,�����Ҫ����html���ø߼�����</li>"
            End If

        ElseIf FilterType = 2 Then

            If FisString = "" Or FioString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��ʼ/������ǲ���Ϊ��</li>"
            End If

        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������,�����Ч���ӽ���</li>"
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    If FoundErr <> True Then
        SqlItem = "select top 1 *  from PE_Filters"
        If FilterID <> 0 Then
            SqlItem = SqlItem & " where FilterID=" & FilterID
        End If

        Set rsFilters = Server.CreateObject("adodb.recordset")
        rsFilters.Open SqlItem, Conn, 1, 3
        If FilterID = 0 Then
            rsFilters.addnew
        End If
        rsFilters("FilterName") = FilterName
        rsFilters("ItemID") = ItemID
        rsFilters("FilterObject") = FilterObject
        rsFilters("FilterType") = FilterType

        If FilterType = 1 Then
            rsFilters("FilterContent") = FilterContent
        ElseIf FilterType = 2 Then
            rsFilters("FisString") = FisString
            rsFilters("FioString") = FioString
        End If

        rsFilters("FilterRep") = FilterRep

        If Flag = "Yes" Then
            rsFilters("Flag") = True
        Else
            rsFilters("Flag") = False
        End If

        rsFilters.Update
        rsFilters.Close
        Set rsFilters = Nothing
        If FilterID = 0 Then
            Call WriteSuccessMsg("<li>�Ѿ��ɹ�����˹�����Ŀ!", "Admin_Filter.asp?Action=main")
        Else
            Call WriteSuccessMsg("<li>�Ѿ��ɹ��޸��˹�����Ŀ!", "Admin_Filter.asp?Action=main")
        End If
    Else
        Exit Sub
    End If
End Sub
'=================================================
'��������Del
'��  �ã�ɾ��������Ŀ
'=================================================
Sub Del()
    Dim FilterID, sql
    FilterID = Trim(Request("FilterID"))
	If IsValidID(FilterID) = False Then
		FilterID = ""
	End If

    If FilterID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ID����Ϊ��</li>"
    Else

        If InStr(FilterID, ",") > 0 Then
            sql = "Delete From [PE_Filters] Where FilterID In (" & FilterID & ")"
        Else
            sql = "Delete From [PE_Filters] Where FilterID=" & FilterID
        End If

        Conn.Execute (sql)
		Call WriteSuccessMsg("<li>�Ѿ��ɹ�ɾ��������Ŀ!", "Admin_Filter.asp?Action=main")
    End If
End Sub
'=================================================
'��������DelAll
'��  �ã���չ���������Ŀ
'=================================================
Sub DelAll()
    Conn.Execute ("Delete From PE_Filters")
    Call WriteSuccessMsg("<li>�Ѿ��ɹ���չ���������Ŀ!", "Admin_Filter.asp?Action=main")
End Sub
'=================================================
'��������SetFlag
'��  �ã��Ƿ�����
'=================================================
Sub SetFlag()
    Dim FilterID
    FilterID = Trim(Request("FilterID"))

    If FilterID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ID����Ϊ��</li>"
        Exit Sub
    Else
        If PE_CLng(Trim(Request("FilterFlag"))) = 1 Then
            sql = "update PE_Filters set Flag=" & PE_True & " where FilterID=" & PE_CLng(FilterID)
        Else
            sql = "update PE_Filters set Flag=" & PE_False & " where FilterID=" & PE_CLng(FilterID)
        End If
        Conn.Execute (sql)
    End If

    Response.Redirect "Admin_Filter.asp?Action=main"
End Sub
'*************************  ��ģ�����������  *******************************
'==================================================
'��������ShowItem_Option
'��  �ã���ʾ��Ŀ����
'��  ����ItemID ------��ĿID
'==================================================
Sub ShowItem_Option(ItemID)
    Dim SqlI, RsI
    SqlI = "select ItemID,ItemName from PE_Item order by ItemID desc"
    Set RsI = Server.CreateObject("adodb.recordset")
    RsI.Open SqlI, Conn, 1, 1
    Response.Write "<select Name=""ItemID"" ID=""ItemID"" >"

    If RsI.EOF And RsI.BOF Then
        Response.Write "<option value=""0"">�������Ŀ</option>"
    Else
        Response.Write "<option value=""0"" "

        If ItemID = -1 Then
            Response.Write " Selected"
        End If

        Response.Write ">��ѡ����Ŀ</option>"
        
        Do While Not RsI.EOF
            Response.Write "<option value=" & """" & RsI("ItemID") & """" & ""

            If ItemID = RsI("ItemID") Then
                Response.Write " Selected"
            End If

            Response.Write ">" & RsI("ItemName")
            Response.Write "</option>"
            RsI.MoveNext
        Loop

    End If

    Response.Write "<option value=""-1"" "

    If ItemID = -1 Then
        Response.Write " Selected"
    End If

    Response.Write ">������Ŀ</option>"
    Response.Write "</select>"
    RsI.Close
    Set RsI = Nothing
End Sub

%>