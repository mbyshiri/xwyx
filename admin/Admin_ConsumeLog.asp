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
Const PurviewLevel_Others = "ConsumeLog"   '����Ȩ��
Dim BeginDate,EndDate
BeginDate = Trim(Request("BeginDate"))
EndDate = Trim(Request("EndDate"))

strFileName = "Admin_ConsumeLog.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword &"&BeginDate="&BeginDate&"&EndDate="&EndDate

Response.Write "<html><head><title>" & PointName & "��ϸ��ѯ</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0'  marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
Call ShowPageTitle(PointName & "��ϸ��ѯ", 10044)
Response.Write "    <tr class='tdbg' height='30'> "
Response.Write "  <form name='form1' action='Admin_ConsumeLog.asp' method='get'>"
Response.Write "      <td width=400>���ٲ��ң�"
Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
Response.Write "          <option value='0'"
If SearchType = 0 Then Response.Write " selected"
Response.Write ">����" & PointName & "��ϸ��¼</option>"
Response.Write "          <option value='1'"
If SearchType = 1 Then Response.Write " selected"
Response.Write ">���10���ڵ���" & PointName & "��ϸ��¼</option>"
Response.Write "          <option value='2'"
If SearchType = 2 Then Response.Write " selected"
Response.Write ">���һ���ڵ���" & PointName & "��ϸ��¼</option>"
Response.Write "          <option value='3'"
If SearchType = 3 Then Response.Write " selected"
Response.Write ">���������¼</option>"
Response.Write "          <option value='4'"
If SearchType = 4 Then Response.Write " selected"
Response.Write ">����֧����¼</option>"
Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_ConsumeLog.asp'>" & PointName & "��ϸ��ҳ</a></td>"& vbCrLf
Response.Write "  </form>"& vbCrLf
Response.Write " <script language='javascript'>"& vbCrLf
Response.Write " function ChangeSearch(type)"& vbCrLf
Response.Write " {"& vbCrLf
Response.Write "  if (type=='LogTime')"& vbCrLf
Response.Write "     {"& vbCrLf
Response.Write "      document.getElementById(""UserNameTable"").style.display=""none""; "& vbCrLf
Response.Write "      document.getElementById(""LogTimeTable"").style.display=""""; "& vbCrLf
Response.Write "      }"& vbCrLf
Response.Write "   else"& vbCrLf
Response.Write "     {"& vbCrLf
Response.Write "     document.getElementById(""UserNameTable"").style.display=""""; "& vbCrLf
Response.Write "     document.getElementById(""LogTimeTable"").style.display=""none""; "& vbCrLf
Response.Write "     }"& vbCrLf
Response.Write " }"& vbCrLf
Response.Write " </script>"& vbCrLf
Response.Write "<form name='form2' method='post' action='Admin_ConsumeLog.asp'>"
Response.Write "    <td>"
Response.Write "<table><tr><td>�߼���ѯ��"
Response.Write "      <select name='Field' onchange=ChangeSearch(this.options[this.selectedIndex].value) id='Field'>"
Response.Write "      <option value='UserName'>�û���</option>"
Response.Write "      <option value='LogTime'>����ʱ��</option>"
Response.Write "      </select>"
Response.Write "      <td><Table id=UserNameTable style=""DISPLAY""><tr><td><input name='Keyword' style='display' type='text' id='Keyword' size='20' maxlength='30'></td></tr></Table></td>"
Response.Write "      <td><table id=LogTimeTable style=""DISPLAY: none""><tr><td>��ʼ����<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.form2.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;��������<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.form2.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr></Table></td>"
Response.Write "      <td>"	
Response.Write "      <input type='submit' name='Submit2' value=' �� ѯ '>"
Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'></td></tr></table>"
Response.Write " </td>"
Response.Write "</form>"
Response.Write "</table>"
Response.Write "<br>"

If Action = "Del" Then
    Call Del
Else
    Call main
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rsConsumeLog, sqlConsumeLog
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    sqlConsumeLog = "select * from PE_ConsumeLog "
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>�����ڵ�λ�ã�<a href='Admin_ConsumeLog.asp'>" & PointName & "��ϸ��¼����</a>&nbsp;&gt;&gt;&nbsp;"
    Select Case SearchType
        Case 0
            sqlConsumeLog = sqlConsumeLog & " order by LogID desc"
            Response.Write "����" & PointName & "��ϸ��¼"
        Case 1
            sqlConsumeLog = sqlConsumeLog & " where datediff(" & PE_DatePart_D & ",LogTime," & PE_Now & ")<10 order by LogID desc"
            Response.Write "���10���ڵ���" & PointName & "��ϸ��¼"
        Case 2
            sqlConsumeLog = sqlConsumeLog & " where datediff(" & PE_DatePart_M & ",LogTime," & PE_Now & ")<1 order by LogID desc"
            Response.Write "���һ���ڵ���" & PointName & "��ϸ��¼"
        Case 3
            sqlConsumeLog = sqlConsumeLog & " where Income_Payout=1 order by LogID desc"
            Response.Write "���������¼"
        Case 4
            sqlConsumeLog = sqlConsumeLog & " where Income_Payout=2 order by LogID desc"
            Response.Write "����֧����¼"
        Case 5
            sqlConsumeLog = sqlConsumeLog & " where Income_Payout<=2 order by LogID desc"
            Response.Write "���зǿ�����¼"
        Case 10
            If Keyword = "" and BeginDate = "" and EndDate = "" Then
                sqlConsumeLog = sqlConsumeLog & " order by LogID desc"
                Response.Write "����" & PointName & "��ϸ��¼"
            Else

                Select Case strField
                Case "UserName"
                    sqlConsumeLog = sqlConsumeLog & " where UserName like '%" & Keyword & "%' order by LogID desc"
                    Response.Write "�û����к��С� <font color=red>" & Keyword & "</font> ����" & PointName & "��ϸ��¼"
                Case "LogTime"
                    sqlConsumeLog = sqlConsumeLog & " where 1=1"
                    If (IsDate(BeginDate) and EndDate="") Or (IsDate(BeginDate) and IsDate(EndDate)) Or (IsDate(EndDate) and BeginDate="") Then
                        If SystemDatabaseType = "SQL"  Then 
                            If BeginDate<>"" Then
                                 sqlConsumeLog = sqlConsumeLog & " and LogTime>='" & BeginDate &"'"
                            End If
                            If EndDate<>"" Then
                                sqlConsumeLog = sqlConsumeLog & " and LogTime<='" & EndDate &"'"
                            End If                      
                        Else
                            If BeginDate<>"" Then
                                sqlConsumeLog = sqlConsumeLog & " and LogTime>=#" & BeginDate &"#"
                            End If
								
                            If EndDate<>"" Then
                                sqlConsumeLog = sqlConsumeLog & " and LogTime<=#" & EndDate &"#"
                            End If 
                        End If
                            sqlConsumeLog = sqlConsumeLog &"  order by LogID desc"
                            If (IsDate(BeginDate) and EndDate="") Then Response.Write "����ʱ��Ϊ <font color=red>" & BeginDate & "֮��</font> ��" & PointName & "��ϸ��¼"
                            If (IsDate(BeginDate) and IsDate(EndDate)) Then Response.Write "����ʱ��Ϊ <font color=red>" & BeginDate & "</font> �� <font color=red>"& EndDate &"</font> ֮���" & PointName & "��ϸ��¼"
                            If (IsDate(EndDate) and BeginDate="") Then Response.Write "����ʱ��Ϊ <font color=red>" & EndDate & "֮ǰ</font> ��" & PointName & "��ϸ��¼"
                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>��ѯ������ʱ�����Ϊ���ڸ�ʽ��</li>"
                    End If
                End Select
            End If
        Case Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ĳ�����</li>"
    End Select
    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    Call PopCalendarInit
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>����ʱ��</td>"
    Response.Write "    <td width='80'>�û���</td>"
    Response.Write "    <td width='100'>IP��ַ</td>"
    Response.Write "    <td width='50'>����" & PointName & "��</td>"
    Response.Write "    <td width='50'>֧��" & PointName & "��</td>"
    Response.Write "    <td width='40'>ժҪ</td>"
    Response.Write "    <td width='60'>�ظ�����</td>"
    Response.Write "    <td width='60'>����Ա</td>"
    Response.Write "    <td>��ע/˵��</td>"
    Response.Write "  </tr>"
    
    Set rsConsumeLog = Server.CreateObject("Adodb.RecordSet")
    rsConsumeLog.Open sqlConsumeLog, Conn, 1, 1
    If rsConsumeLog.BOF And rsConsumeLog.EOF Then
        TotalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κη������������Ѽ�¼��</td></tr>"
    Else
        TotalPut = rsConsumeLog.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rsConsumeLog.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim i
        i = 0
        Do While Not rsConsumeLog.EOF
            If rsConsumeLog("Income_Payout") = 1 Or rsConsumeLog("Income_Payout") = 3 Then
                TotalIncome = TotalIncome + rsConsumeLog("Point")
            Else
                TotalPayout = TotalPayout + rsConsumeLog("Point")
            End If
    
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='120' align='center'>" & rsConsumeLog("LogTime") & "</td>"
            Response.Write "    <td width='80' align='center'><a href='Admin_User.asp?Action=Show&InfoType=2&UserName=" & rsConsumeLog("UserName") & "'>" & rsConsumeLog("UserName") & "</a></td>"
            Response.Write "    <td width='100' align='center'>" & rsConsumeLog("IP") & "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsConsumeLog("Income_Payout") = 1 Then Response.Write rsConsumeLog("Point")
            Response.Write "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsConsumeLog("Income_Payout") = 2 Then Response.Write rsConsumeLog("Point")
            Response.Write "</td>"
            Response.Write "    <td width='40' align='center'>"
            Select Case rsConsumeLog("Income_Payout")
            Case 1
                Response.Write "<font color='blue'>����</font>"
            Case 2
                Response.Write "<font color='green'>֧��</font>"
            Case Else
                Response.Write "����"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>" & rsConsumeLog("Times") & "</td>"
            Response.Write "    <td width='60' align='center'>" & rsConsumeLog("Inputer") & "</td>"
            Response.Write "    <td align='left'>" & rsConsumeLog("Remark") & "</td>"
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsConsumeLog.MoveNext
        Loop
    End If
    rsConsumeLog.Close
    Set rsConsumeLog = Nothing

    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='3' align='right'>��ҳ�ϼƣ�</td>"
    Response.Write "    <td align='right'>" & TotalIncome & "</td>"
    Response.Write "    <td align='right'>" & TotalPayout & "</td>"
    Response.Write "    <td colspan='4'>&nbsp;</td>"
    Response.Write "  </tr>"

    Dim trs, TotalIncomeAll, TotalPayoutAll
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where Income_Payout=1")
    If IsNull(trs(0)) Then
        TotalIncomeAll = 0
    Else
        TotalIncomeAll = trs(0)
    End If
    Set trs = Nothing
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where Income_Payout=2")
    If IsNull(trs(0)) Then
        TotalPayoutAll = 0
    Else
        TotalPayoutAll = trs(0)
    End If
    Set trs = Nothing
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='3' align='right'>�ܼƽ�</td>"
    Response.Write "    <td align='right'>" & TotalIncomeAll & "</td>"
    Response.Write "    <td align='right'>" & TotalPayoutAll & "</td>"
    Response.Write "    <td colspan='4' align='center'>" & PointName & "����" & TotalIncomeAll - TotalPayoutAll & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "��" & PointName & "��ϸ��¼", True)
    Response.Write "<form name='myform' method='post' action='Admin_ConsumeLog.asp' onsubmit=""return confirm('ȷʵҪɾ���йؼ�¼��һ��ɾ����Щ��¼������ֻ�Ա�鿴ԭ���Ѿ������ѵ��շ���Ϣʱ�ظ��շѵ����⡣�����أ�')"">"
    Response.Write "���" & PointName & "��ϸ��¼̫�࣬Ӱ����ϵͳ���ܣ�����ɾ��һ��ʱ���ǰ�ļ�¼�Լӿ��ٶȡ������ܻ������Ա�ڲ鿴��ǰ�չ��ѵ���Ϣʱ�ظ��շѣ������������ڶ����Ѿ������⣩���޷�ͨ��" & PointName & "��ϸ��¼����ʵ������Ա������ϰ�ߵ����⡣<br>"
    Response.Write "ʱ�䷶Χ��<input type='radio' name='DatepartType' value='0'>10��ǰ&nbsp;&nbsp;<input type='radio' name='DatepartType' value='1'>1����ǰ&nbsp;&nbsp;<input type='radio' name='DatepartType' value='2'>2����ǰ&nbsp;&nbsp;<input type='radio' name='DatepartType' value='3'>3����ǰ&nbsp;&nbsp;<input type='radio' name='DatepartType' value='4'>6����ǰ&nbsp;&nbsp;<input type='radio' name='DatepartType' value='5' checked>1��ǰ&nbsp;&nbsp;<input type='submit' name='submit1' value='ɾ����¼'>"
    Response.Write "<input type='hidden' name='Action' value='Del'></form>"
End Sub

Sub Del()
    Dim DatepartType, TempDate, strDatepart
    DatepartType = PE_CLng(Trim(Request("DatepartType")))
    Select Case DatepartType
    Case 0
        TempDate = DateAdd("D", -10, Date)
        strDatepart = "10��ǰ"
    Case 1
        TempDate = DateAdd("M", -1, Date)
        strDatepart = "1����ǰ"
    Case 2
        TempDate = DateAdd("M", -2, Date)
        strDatepart = "2����ǰ"
    Case 3
        TempDate = DateAdd("M", -3, Date)
        strDatepart = "3����ǰ"
    Case 4
        TempDate = DateAdd("M", -6, Date)
        strDatepart = "6����ǰ"
    Case 5
        TempDate = DateAdd("yyyy", -1, Date)
        strDatepart = "1��ǰ"
    End Select
    If SystemDatabaseType = "SQL" Then
        Conn.Execute ("delete from PE_ConsumeLog where LogTime<'" & TempDate & "'")
    Else
        Conn.Execute ("delete from PE_ConsumeLog where LogTime<#" & TempDate & "#")
    End If
    Call WriteSuccessMsg("�ɹ�ɾ���� " & strDatepart & " �ļ�¼��", "Admin_ConsumeLog.asp")
End Sub
%>
