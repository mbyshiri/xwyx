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
Const PurviewLevel_Others = "RechargeLog"   '����Ȩ��

strFileName = "Admin_RechargeLog.asp?SearchType=" & SearchType & "&Field=" & strField & "&Keyword=" & Keyword

Response.Write "<html><head><title>��Ч����ϸ��ѯ</title>"
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link rel='stylesheet' href='Admin_Style.css' type='text/css'>"
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
Call ShowPageTitle("�� Ч �� �� ϸ �� ѯ", 10045)
Response.Write "    <tr class='tdbg' height='30'> "
Response.Write "  <form name='form1' action='Admin_RechargeLog.asp' method='get'>"
Response.Write "      <td>���ٲ��ң�"
Response.Write "      <select size=1 name='SearchType' onChange='javascript:submit()'>"
Response.Write "          <option value='0'"
If SearchType = 0 Then Response.Write " selected"
Response.Write ">������Ч����ϸ��¼</option>"
Response.Write "          <option value='1'"
If SearchType = 1 Then Response.Write " selected"
Response.Write ">���10���ڵ�����Ч����ϸ��¼</option>"
Response.Write "          <option value='2'"
If SearchType = 2 Then Response.Write " selected"
Response.Write ">���һ���ڵ�����Ч����ϸ��¼</option>"
Response.Write "          <option value='3'"
If SearchType = 3 Then Response.Write " selected"
Response.Write ">���������¼</option>"
Response.Write "          <option value='4'"
If SearchType = 4 Then Response.Write " selected"
Response.Write ">����֧����¼</option>"
Response.Write "          <option value='5'"
If SearchType = 5 Then Response.Write " selected"
Response.Write ">���зǿ�����¼</option>"
Response.Write "        </select>&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_RechargeLog.asp'>��Ч����ϸ��ҳ</a></td>"
Response.Write "  </form>"
Response.Write "<form name='form2' method='post' action='Admin_RechargeLog.asp'>"
Response.Write "    <td>�߼���ѯ��"
Response.Write "      <select name='Field' id='Field'>"
Response.Write "      <option value='UserName'>�û���</option>"
Response.Write "      <option value='LogTime'>ʱ��</option>"
Response.Write "      </select>"
Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='30'>"
Response.Write "      <input type='submit' name='Submit2' value=' �� ѯ '>"
Response.Write "      <input name='SearchType' type='hidden' id='SearchType' value='10'>"
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
    Dim rsRechargeLog, sqlRechargeLog
    Dim TotalIncome, TotalPayout
    TotalIncome = 0
    TotalPayout = 0
    
    sqlRechargeLog = "select * from PE_RechargeLog "
    Response.Write "<table width='100%'><tr><td align='left'><img src='images/img_u.gif' align='absmiddle'>�����ڵ�λ�ã�<a href='Admin_Bankroll.asp'>��Ч����ϸ��¼����</a>&nbsp;&gt;&gt;&nbsp;"
    Select Case SearchType
        Case 0
            sqlRechargeLog = sqlRechargeLog & " order by LogID desc"
            Response.Write "������Ч����ϸ��¼"
        Case 1
            sqlRechargeLog = sqlRechargeLog & " where datediff(" & PE_DatePart_D & ",LogTime," & PE_Now & ")<10 order by LogID desc"
            Response.Write "���10���ڵ�����Ч����ϸ��¼"
        Case 2
            sqlRechargeLog = sqlRechargeLog & " where datediff(" & PE_DatePart_M & ",LogTime," & PE_Now & ")<1 order by LogID desc"
            Response.Write "���һ���ڵ�����Ч����ϸ��¼"
        Case 3
            sqlRechargeLog = sqlRechargeLog & " where Income_Payout=1 order by LogID desc"
            Response.Write "���������¼"
        Case 4
            sqlRechargeLog = sqlRechargeLog & " where Income_Payout=2 order by LogID desc"
            Response.Write "����֧����¼"
        Case 10
            If Keyword = "" Then
                sqlRechargeLog = sqlRechargeLog & " order by LogID desc"
                Response.Write "������Ч����ϸ��¼"
            Else
                Select Case strField
                Case "UserName"
                    sqlRechargeLog = sqlRechargeLog & " where UserName like '%" & Keyword & "%' order by LogID desc"
                    Response.Write "�û����к��С� <font color=red>" & Keyword & "</font> ������Ч����ϸ��¼"
                Case "LogTime"
                    If IsDate(Keyword) Then
                        If SystemDatabaseType = "SQL" Then
                            sqlRechargeLog = sqlRechargeLog & " where LogTime='" & Keyword & "'  order by LogID desc"
                        Else
                            sqlRechargeLog = sqlRechargeLog & " where LogTime=#" & Keyword & "#  order by LogID desc"
                        End If
                        Response.Write "����ʱ��Ϊ <font color=red>" & Keyword & "</font> ����Ч����ϸ��¼"
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
    
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>ʱ��</td>"
    Response.Write "    <td width='80'>�û���</td>"
    Response.Write "    <td width='120'>IP��ַ</td>"
    Response.Write "    <td width='50'>������Ч��</td>"
    Response.Write "    <td width='50'>������Ч��</td>"
    Response.Write "    <td width='40'>ժҪ</td>"
    Response.Write "    <td width='60'>����Ա</td>"
    Response.Write "    <td>��ע/˵��</td>"
    Response.Write "  </tr>"
    
    Set rsRechargeLog = Server.CreateObject("Adodb.RecordSet")
    rsRechargeLog.Open sqlRechargeLog, Conn, 1, 1
    If rsRechargeLog.BOF And rsRechargeLog.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κη������������Ѽ�¼��</td></tr>"
    Else
        totalPut = rsRechargeLog.RecordCount
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
                rsRechargeLog.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim i
        i = 0
        Do While Not rsRechargeLog.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='120' align='center'>" & rsRechargeLog("LogTime") & "</td>"
            Response.Write "    <td width='80' align='center'><a href='Admin_User.asp?Action=Show&InfoType=3&UserName=" & rsRechargeLog("UserName") & "'>" & rsRechargeLog("UserName") & "</a></td>"
            Response.Write "    <td width='120' align='center'>" & rsRechargeLog("IP") & "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsRechargeLog("Income_Payout") = 1 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='50' align='right'>"
            If rsRechargeLog("Income_Payout") = 2 Then
                If rsRechargeLog("ValidNum") > 0 Then
                    Response.Write rsRechargeLog("ValidNum") & " " & arrCardUnit(rsRechargeLog("ValidUnit"))
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td width='40' align='center'>"
            Select Case rsRechargeLog("Income_Payout")
            Case 1
                Response.Write "<font color='blue'>����</font>"
            Case 2
                Response.Write "<font color='green'>����</font>"
            Case Else
                Response.Write "����"
            End Select
            Response.Write "</td>"
            Response.Write "    <td width='60' align='center'>" & rsRechargeLog("Inputer") & "</td>"
            Response.Write "    <td align='left'>" & rsRechargeLog("Remark") & "</td>"
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsRechargeLog.MoveNext
        Loop
    End If
    rsRechargeLog.Close
    Set rsRechargeLog = Nothing

    Response.Write "</table>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����Ч����ϸ��¼", True)

    Response.Write "<form name='myform' method='post' action='Admin_RechargeLog.asp' onsubmit=""return confirm('ȷʵҪɾ���йؼ�¼��һ��ɾ����Щ��¼������ֻ�Ա�鿴ԭ���Ѿ������ѵ��շ���Ϣʱ�ظ��շѵ����⡣�����أ�')"">"
    Response.Write "�����ȯ��ϸ��¼̫�࣬Ӱ����ϵͳ���ܣ�����ɾ��һ��ʱ���ǰ�ļ�¼�Լӿ��ٶȡ������ܻ������Ա�ڲ鿴��ǰ�չ��ѵ���Ϣʱ�ظ��շѣ������������ڶ����Ѿ������⣩���޷�ͨ����ȯ��ϸ��¼����ʵ������Ա������ϰ�ߵ����⡣<br>"
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
        Conn.Execute ("delete from PE_RechargeLog where LogTime<'" & TempDate & "'")
    Else
        Conn.Execute ("delete from PE_RechargeLog where LogTime<#" & TempDate & "#")
    End If
    Call WriteSuccessMsg("�ɹ�ɾ���� " & strDatepart & " �ļ�¼��", "Admin_RechargeLog.asp")
End Sub
%>
