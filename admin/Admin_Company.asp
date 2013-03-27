<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_CRM.asp"-->
<!--#include file="../Include/PowerEasy.Bankroll.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = False   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Company"   '����Ȩ��

Dim arrStatusInField, arrCompanySize, arrManagementForms
arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")

Response.Write "<html><head><title>��ҵ����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <form name='searchmyform' action='Admin_Company.asp' method='get'>"
Call ShowPageTitle("�� ҵ �� ��", 10221)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td height='30'>"
Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_Company.asp'>��ҵ������ҳ</a>&nbsp;|&nbsp;<a href='Admin_Company.asp?Action=AddCompany&ClientType=0'>�����ҵ</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  </form>" & vbCrLf
Response.Write "</table>" & vbCrLf
Select Case Action
    Case "AddCompany"
        Call AddCompany
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SaveCompany
    Case "Show"
        Call Show
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim GroupID, i
    Dim sql, Querysql, rsCompanyList

    GroupID = PE_CLng(Trim(Request("GroupID")))
    strFileName = "Admin_Company.asp?SearchType=" & SearchType & "&Field=" & strField & "&keyword=" & Keyword & "&GroupID=" & GroupID
    
    Call ShowJS_Main("��ҵ")
        
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Company.asp'>��ҵ����</a>&nbsp;&gt;&gt;&nbsp;"
    sql = "select top " & MaxPerPage & " * from PE_Company"
    Querysql = Querysql & " where 1=1 "
    Select Case SearchType
    Case 0
        Response.Write "������ҵ"
    Case 1
        Querysql = Querysql & " and ClientType=0"
        Response.Write "��ҵ��ҵ"
    Case 2
        Querysql = Querysql & " and ClientType=1"
        Response.Write "������ҵ"
    Case 10
        If Keyword = "" Then
            Response.Write "������ҵ"
        Else
            Select Case strField
            Case "CompanyID"
                If IsNumeric(Keyword) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ҵID����������</li>"
                Else
                    Querysql = Querysql & " and CompanyID=" & PE_CLng(Keyword)
                    Response.Write "��ҵID����<font color=red> " & PE_CLng(Keyword) & " </font>����ҵ"
                End If
            Case "CompanyName"
                Querysql = Querysql & " and CompanyName like '%" & Keyword & "%'"
                Response.Write "�����к��С� <font color=red>" & Keyword & "</font> ������ҵ"
            Case Else
                Querysql = Querysql & " and CompanyName like '%" & Keyword & "%'"
                Response.Write "�����к��С� <font color=red>" & Keyword & "</font> ������ҵ"
            End Select
        End If
    Case 11
        Response.Write GetArrItem(arrGroupID, GroupID)
        Querysql = Querysql & " and GroupID=" & GroupID
    End Select
    
    TotalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Company " & Querysql)(0))
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
        Querysql = Querysql & " and CompanyID < (select min(CompanyID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " CompanyID from PE_Company " & Querysql & " order by CompanyID desc) as QueryClient) "
    End If
    sql = sql & Querysql & " order by CompanyID desc"


    Response.Write "</td></tr></table>"
    If FoundErr = True Then Exit Sub
    
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Company.asp'>"
    Response.Write "      <td>"
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "        <tr class='title' height='22' align='center'>"
    Response.Write "          <td width='30'>ѡ��</td>"
    Response.Write "          <td>��ҵ����</td>"
    Response.Write "          <td>��ϵ��ַ</td>"
    Response.Write "          <td width='90'>����</td>"
    Response.Write "        </tr>"
    
    Set rsCompanyList = Server.CreateObject("Adodb.RecordSet")
    rsCompanyList.Open sql, Conn, 1, 1
    If rsCompanyList.BOF And rsCompanyList.EOF Then
        Response.Write "<tr><td colspan='20' height='50' align='center'>���ҵ� <font color=red>0</font> ����ҵ</td></tr>"
    Else
        Dim ClientNum
        ClientNum = 0
        Do While Not rsCompanyList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'><input name='CompanyID' type='checkbox' onclick=""unselectall()"" id='CompanyID' value='" & CStr(rsCompanyList("CompanyID")) & "'></td>"
            Response.Write "        <td width='100'><a href='Admin_Company.asp?Action=Show&CompanyID=" & rsCompanyList("CompanyID") & "'>" & rsCompanyList("ShortedForm") & "</a></td>"
            Response.Write "        <td><a href='Admin_Company.asp?Action=Show&CompanyID=" & rsCompanyList("CompanyID") & "'>" & rsCompanyList("CompanyName") & "</a></td>"
            Response.Write "        <td>" & rsCompanyList("Address") & "</td>"
            Response.Write "        <td width='90' align='center'>"
            Response.Write "<a href='Admin_Company.asp?Action=Show&CompanyID=" & rsCompanyList("CompanyID") & "'>�鿴</a>&nbsp;"
            Response.Write "<a href='Admin_Company.asp?Action=Modify&CompanyID=" & rsCompanyList("CompanyID") & "'>�޸�</a>&nbsp;"
            Response.Write "<a href='Admin_Company.asp?Action=DelClient&CompanyID=" & rsCompanyList("CompanyID") & "' onClick='return confirm(""ȷ��Ҫɾ������ҵ��"");'>ɾ��</a> "
            Response.Write "        </td>"
            Response.Write "      </tr>"

            ClientNum = ClientNum + 1
            If ClientNum >= MaxPerPage Then Exit Do
            rsCompanyList.MoveNext
        Loop
    End If
    rsCompanyList.Close
    Set rsCompanyList = Nothing
    Response.Write "      </table>"
    If TotalPut > 0 Then
        Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "        <tr height='30'>"
        Response.Write "          <td width='200'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form);' value='checkbox'>"
        Response.Write "          ѡ�б�ҳ��ʾ��������ҵ</td>"
        Response.Write "          <td><input type='hidden' name='Action' value=''>"
        Response.Write "          <input name='Del' type='submit' value=' ����ɾ�� ' onClick=""document.myform.Action.value='DelClient';return confirm('ȷ��Ҫɾ��ѡ������ҵ��');"">&nbsp;&nbsp;&nbsp;&nbsp;"
        'Response.Write "          <input name='BatchMove' type='submit' value=' �����ƶ� ' onClick=""document.myform.Action.value='BatchMove'"">"
        Response.Write "        </tr>"
        Response.Write "      </table>"
    End If
    Response.Write "      </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    If TotalPut > 0 Then
        Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "����ҵ", True)
    End If

    Response.Write "<br>"
    Response.Write "<form name='SearchForm' action='Admin_Company.asp' method='post'>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='1' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='80'><b>��ҵ��ѯ��</b></td>" & vbCrLf
    Response.Write "    <td>" & vbCrLf
    Response.Write "      <select name='Field' size='1'>" & vbCrLf
    Response.Write "        <option value='CompanyID'>��ҵID</option>" & vbCrLf
    Response.Write "        <option value='CompanyName' selected>��ҵ����</option>" & vbCrLf
    Response.Write "      </select>" & vbCrLf
    Response.Write "      <input name='Keyword' type='text' id='Keyword' size='20' maxlength='40'>" & vbCrLf
    Response.Write "      <input type='submit' name='Submit' value='�� ��' id='Submit'>" & vbCrLf
    Response.Write "      <input type='hidden' name='SearchType' value='10'>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub ShowJS_Check()
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "    if(document.myform.CompanyName.value==''){" & vbCrLf
    Response.Write "        alert('��ҵ���Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "        document.myform.CompanyName.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(document.myform.ShortedForm.value==''){" & vbCrLf
    Response.Write "        alert('�������Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "        document.myform.ShortedForm.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.myform.Country.value=frm1.document.regionform.Country.value;" & vbCrLf
    Response.Write "    document.myform.Province.value=frm1.document.regionform.Province.value;" & vbCrLf
    Response.Write "    document.myform.City.value=frm1.document.regionform.City.value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ChangeType(Type){" & vbCrLf
    Response.Write "  if(Type==0){" & vbCrLf
    Response.Write "    TabTitle[2].style.display='';" & vbCrLf
    Response.Write "    infoE.style.display='';" & vbCrLf
    Response.Write "    TabTitle[3].style.display='none';" & vbCrLf
    Response.Write "    TabTitle[4].style.display='none';" & vbCrLf
    Response.Write "    infoP.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    TabTitle[2].style.display='none';" & vbCrLf
    Response.Write "    infoE.style.display='none';" & vbCrLf
    Response.Write "    TabTitle[3].style.display='';" & vbCrLf
    Response.Write "    TabTitle[4].style.display='';" & vbCrLf
    Response.Write "    infoP.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function SelectParentClient(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=ClientList','','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        var ss=arr.split('$$$');" & vbCrLf
    Response.Write "        document.myform.Parenter.value=ss[0];" & vbCrLf
    Response.Write "        document.myform.ParentID.value=ss[1];" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub AddCompany()
    Call PopCalendarInit
    Call ShowJS_Check
    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Company.asp'>��ҵ����</a>&nbsp;&gt;&gt;&nbsp;�����ҵ</td></tr></table>"
    Response.Write "<form name='myform' id='myform' action='Admin_Company.asp' method='post' onSubmit='return CheckSubmit();'>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "        <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td height='100' valign='top'>"
    Response.Write "                <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF' id='Tabs' style='display:'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ҵ���ƣ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='CompanyName' type='text' id='CompanyName' size='35' maxlength='200'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ҵ��ţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='ClientNum' type='text' id='ClientNum' size='35' maxlength='30' value='" & GetNumString() & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>ͨѶ��ַ��</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <iframe name='frm1' id='frm1' src='../Region.asp' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "                            <input name='Country' id='Country' type='hidden'> <input name='Province' id='Province' type='hidden'> <input name='City' id='City' type='hidden'>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >��ϵ��ַ��</td>" & vbCrLf
    Response.Write "                                    <td><input name='Address' type='text' size='60' maxlength='255'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td align='right' class='tdbg5' align='right' >�������룺</td>" & vbCrLf
    Response.Write "                                    <td><input name='ZipCode' type='text' size='35' maxlength='10'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                            </table>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ϵ�绰��</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Phone' type='text' size='35' maxlength='30'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>������룺</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Fax1' type='text' size='35' maxlength='30'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�������У�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='BankOfDeposit' type='text' size='35' maxlength='255'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�����ʺţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='BankAccount' type='text' size='35' maxlength='255'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >˰�ţ�</td>" & vbCrLf
    Response.Write "                        <td><input name='TaxNum' type='text' size='35' maxlength='20'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ַ��</td>" & vbCrLf
    Response.Write "                        <td><input name='Homepage1' type='text' size='35' maxlength='100'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ҵ��λ��</td>" & vbCrLf
    Response.Write "                        <td><select name='StatusInField'>" & Array2Option(arrStatusInField, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��˾��ģ��</td>" & vbCrLf
    Response.Write "                        <td><select name='CompanySize'>" & Array2Option(arrCompanySize, -1) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ҵ��Χ��</td>" & vbCrLf
    Response.Write "                        <td><input name='BusinessScope' type='text' size='35' maxlength='255'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�����۶</td>" & vbCrLf
    Response.Write "                        <td><input name='AnnualSales' type='text' size='15' maxlength='20'> ��Ԫ</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��Ӫ״̬��</td>" & vbCrLf
    Response.Write "                        <td><select name='ManagementForms'>" & Array2Option(arrManagementForms, -1) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ע���ʱ���</td>" & vbCrLf
    Response.Write "                        <td><input name='RegisteredCapital' type='text' size='15' maxlength='20'> ��Ԫ</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ע��</td>" & vbCrLf
    Response.Write "                        <td colspan='3'><input name='Remark' type='text' size='35'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                </table>" & vbCrLf
    
    Response.Write "            </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' height='40'><tr align='center'><td>"
    Response.Write "    <input type='hidden' name='action' value='SaveAdd'>" & vbCrLf
    Response.Write "    <input type='submit' name='Submit' value='������ҵ��Ϣ'>" & vbCrLf
    Response.Write "</td></tr></table>"
    Response.Write "</form>" & vbCrLf
End Sub


Sub Modify()
    Dim CompanyID, Remark
    CompanyID = Trim(Request("CompanyID"))
    If CompanyID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    Else
        CompanyID = CLng(CompanyID)
    End If
    Dim rsCompany, rsContacter
    Dim Country, Province, City, Address, ZipCode
    Dim Phone, Fax
    Set rsCompany = Conn.Execute("select * from PE_Company where CompanyID=" & CompanyID & "")
    If rsCompany.BOF And rsCompany.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ�����Ӧ����ҵ��Ϣ��</li>"
    Else
        Country = rsCompany("Country")
        Province = rsCompany("Province")
        City = rsCompany("City")
        Address = rsCompany("Address")
        ZipCode = rsCompany("ZipCode")
        Phone = rsCompany("Phone")
        Fax = rsCompany("Fax")
    End If
    If FoundErr = True Then
        Exit Sub
    End If


    Call PopCalendarInit
    Call ShowJS_Check


    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Company.asp'>��ҵ����</a>&nbsp;&gt;&gt;&nbsp;�޸���ҵ��Ϣ</td></tr></table>"
    Response.Write "<form name='myform' id='myform' action='Admin_Company.asp' method='post' onSubmit='return CheckSubmit();'>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "        <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td height='100' valign='top'>"
    Response.Write "                <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF' id='Tabs' style='display:'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ҵ���ƣ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='CompanyName' type='text' id='CompanyName' size='35' maxlength='200' value='" & rsCompany("CompanyName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ҵ��ţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='ClientNum' type='text' id='ClientNum' size='35' maxlength='30' value='" & rsCompany("ClientNum") & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf

    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>ͨѶ��ַ��</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <iframe name='frm1' id='frm1' src='../Region.asp?Action=Modify&Country=" & Country & "&Province=" & Province & "&City=" & City & "' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "                            <input name='Country' id='Country' type='hidden'> <input name='Province' id='Province' type='hidden'> <input name='City' id='City' type='hidden'>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >��ϵ��ַ��</td>" & vbCrLf
    Response.Write "                                    <td><input name='Address' type='text' size='60' maxlength='255' value='" & Address & "'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td align='right' class='tdbg5' align='right' >�������룺</td>" & vbCrLf
    Response.Write "                                    <td><input name='ZipCode' type='text' size='35' maxlength='10' value='" & ZipCode & "'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                            </table>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ϵ�绰��</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Phone' type='text' size='35' maxlength='30' value='" & Phone & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>������룺</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Fax1' type='text' size='35' maxlength='30' value='" & Fax & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf

    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�������У�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='BankOfDeposit' type='text' size='35' maxlength='255' value='" & rsCompany("BankOfDeposit") & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�����ʺţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='BankAccount' type='text' size='35' maxlength='255' value='" & rsCompany("BankAccount") & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >˰�ţ�</td>" & vbCrLf
    Response.Write "                        <td><input name='TaxNum' type='text' size='35' maxlength='20' value='" & rsCompany("TaxNum") & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ַ��</td>" & vbCrLf
    Response.Write "                        <td><input name='Homepage1' type='text' size='35' maxlength='100' value='" & rsCompany("Homepage") & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ҵ��λ��</td>" & vbCrLf
    Response.Write "                        <td><select name='StatusInField'>" & Array2Option(arrStatusInField, rsCompany("StatusInField")) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��˾��ģ��</td>" & vbCrLf
    Response.Write "                        <td><select name='CompanySize'>" & Array2Option(arrCompanySize, rsCompany("CompanySize")) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ҵ��Χ��</td>" & vbCrLf
    Response.Write "                        <td><input name='BusinessScope' type='text' size='35' maxlength='255' value='" & rsCompany("BusinessScope") & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�����۶</td>" & vbCrLf
    Response.Write "                        <td><input name='AnnualSales' type='text' size='15' maxlength='20' value='" & rsCompany("AnnualSales") & "'> ��Ԫ</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��Ӫ״̬��</td>" & vbCrLf
    Response.Write "                        <td><select name='ManagementForms'>" & Array2Option(arrManagementForms, rsCompany("ManagementForms")) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ע���ʱ���</td>" & vbCrLf
    Response.Write "                        <td><input name='RegisteredCapital' type='text' size='15' maxlength='20' value='" & rsCompany("RegisteredCapital") & "'> ��Ԫ</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf

    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ע��</td>" & vbCrLf
    Response.Write "                        <td colspan='3'><input name='Remark' value='" & Remark & "' type='text' size='35'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                </table>" & vbCrLf

    Response.Write "            </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' height='40'><tr align='center'><td>"
    Response.Write "    <input type='hidden' name='action' value='SaveModify'><input type='hidden' name='CompanyID' value='" & CompanyID & "'>" & vbCrLf
    Response.Write "    <input type='submit' name='Submit' value='�����޸Ľ��'>" & vbCrLf
    Response.Write "</td></tr></table>"
    Response.Write "</form>" & vbCrLf
    Set rsCompany = Nothing
End Sub


Sub Show()
    Dim CompanyID
    Dim rsCompany, sqlClient
    Dim Country, Province, City, Address, ZipCode
    CompanyID = PE_CLng(Trim(Request("CompanyID")))
    If CompanyID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    End If
    Set rsCompany = Conn.Execute("select * from PE_Company where CompanyID=" & CompanyID & "")
    If rsCompany.BOF And rsCompany.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ�����Ӧ����ҵ��Ϣ��</li>"
    Else
        Country = rsCompany("Country")
        Province = rsCompany("Province")
        City = rsCompany("City")
        Address = rsCompany("Address")
        ZipCode = rsCompany("ZipCode")
    End If
    If FoundErr = True Then
        Exit Sub
    End If


    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Company.asp'>��ҵ����</a>&nbsp;&gt;&gt;&nbsp;�鿴��ҵ��Ϣ��" & rsCompany("CompanyName") & "</td></tr></table>"

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "        <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td height='100' valign='top'>"
    Response.Write "                <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "                  <tbody id='Tabs'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ҵ���ƣ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & rsCompany("CompanyName") & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��ҵ��ţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & rsCompany("ClientNum") & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>����/������</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Country & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>ʡ/�У�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Province & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��/��/����</td>" & vbCrLf
    Response.Write "                        <td>" & City & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�������룺</td>" & vbCrLf
    Response.Write "                        <td>" & ZipCode & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ϵ��ַ��</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & Address & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    

    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ϵ�绰��</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("Phone") & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >������룺</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("Fax") & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    
    
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�������У�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & rsCompany("BankOfDeposit") & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�����ʺţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & rsCompany("BankAccount") & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >˰�ţ�</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("TaxNum") & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ַ��</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("Homepage") & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ҵ��λ��</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(arrStatusInField, rsCompany("StatusInField")) & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��˾��ģ��</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(arrCompanySize, rsCompany("CompanySize")) & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ҵ��Χ��</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("BusinessScope") & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�����۶</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("AnnualSales") & " ��Ԫ</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��Ӫ״̬��</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(arrManagementForms, rsCompany("ManagementForms")) & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ע���ʱ���</td>" & vbCrLf
    Response.Write "                        <td>" & rsCompany("RegisteredCapital") & " ��Ԫ</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf

    rsCompany.Close
    Set rsCompany = Nothing


    Response.Write "            </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' height='60'><tr align='center'><td>"
    If AdminPurview = 1 Or arrPurview(3) = True Or (arrPurview(2) = True And rsCompany("Owner") = AdminName) Then
        Response.Write "<input type='button' name='modify' value='�޸���ҵ��Ϣ' onclick=""window.location.href='Admin_Company.asp?Action=Modify&ClientID=" & ClientID & "';"">"
    End If
    If ClientType = 0 And (AdminPurview = 1 Or arrPurview(1) = True) Then
        Response.Write "&nbsp;&nbsp;<input type='button' name='add' value='�����ϵ��' onclick=""window.location.href='Admin_Contacter.asp?Action=AddContacter&ClientID=" & ClientID & "';"">"
    End If
    If AdminPurview = 1 Or arrPurview(5) = True Then
        Response.Write "&nbsp;&nbsp;<input type='button' name='add' value='��ӷ����¼' onclick=""window.location.href='Admin_Service.asp?Action=Add&ClientID=" & ClientID & "';"">"
    End If
    If AdminPurview = 1 Or arrPurview(6) = True Then
        Response.Write "&nbsp;&nbsp;<input type='button' name='add' value='���Ͷ�߼�¼' onclick=""window.location.href='Admin_Complain.asp?Action=Add&ClientID=" & ClientID & "';"">"
    End If
    If AdminPurview = 1 Then
        If PE_CLng(Conn.Execute("select count(0) from PE_User where ClientID=" & ClientID & "")(0)) = 0 Then
            Response.Write "&nbsp;&nbsp;<input type='button' name='Submit' value='������л��' onClick=""window.location.href='Admin_Company.asp?Action=AddRemit&ClientID=" & ClientID & "'"">"
            Response.Write "&nbsp;&nbsp;<input type='button' name='Submit' value='�����������' onClick=""window.location.href='Admin_Company.asp?Action=AddIncome&ClientID=" & ClientID & "'"">"
            Response.Write "&nbsp;&nbsp;<input type='button' name='Submit' value='���֧����Ϣ' onClick=""window.location.href='Admin_Company.asp?Action=AddPayment&ClientID=" & ClientID & "'"">"
        End If
    End If
    If AdminPurview = 1 Or arrPurview(4) = True Then
        Response.Write "&nbsp;&nbsp;<input type='button' name='modify' value='ɾ������ҵ' onclick=""window.location.href='Admin_Company.asp?Action=DelClient&ClientID=" & ClientID & "';"">"
    End If
    Response.Write "</td></tr></table>"

End Sub



Sub SaveCompany()
    Dim ClientID, CompanyName, ClientNum, ClientType, ShortedForm

    ClientID = PE_CLng(Trim(Request.Form("ClientID")))
    ClientType = PE_CLng(Trim(Request.Form("ClientType")))
    ClientNum = Trim(Request.Form("ClientNum"))
    CompanyName = Trim(Request.Form("CompanyName"))
    ShortedForm = Trim(Request.Form("ShortedForm"))

    If CompanyName = "" Then
        FoundErr = True
        ErrMsg = "��ҵ���Ʋ���Ϊ�գ�"
    End If
    If ShortedForm = "" Then
        FoundErr = True
        ErrMsg = "��ҵ��ƣ������룩����Ϊ�գ�"
    End If

    If FoundErr Then
        Exit Sub
    End If

    Dim sqlCompany, rsCompany, CompanyID
    Set rsCompany = Server.CreateObject("adodb.recordset")

    If Action = "SaveAdd" Then
        sqlCompany = "select top 1 * From PE_Company"
        rsCompany.Open sqlCompany, Conn, 1, 3
        rsCompany.addnew
        rsCompany("CompanyID") = GetNewID("PE_Company", "CompanyID")
    Else
        sqlCompany = "select * From PE_Company Where ClientID=" & ClientID & ""
        rsCompany.Open sqlCompany, Conn, 1, 3
        If rsCompany.BOF And rsCompany.EOF Then
            rsCompany.addnew
            rsCompany("CompanyID") = GetNewID("PE_Company", "CompanyID")
        End If
    End If
    If FoundErr Then
        rsCompany.Close
        Set rsCompany = Nothing
        Exit Sub
    End If

    rsCompany("ClientID") = ClientID
    rsCompany("Country") = Trim(Request.Form("Country"))
    rsCompany("Province") = Trim(Request.Form("Province"))
    rsCompany("City") = Trim(Request.Form("City"))
    rsCompany("Address") = Trim(Request.Form("Address"))
    rsCompany("ZipCode") = Trim(Request.Form("ZipCode"))
    rsCompany("Phone") = Trim(Request.Form("Phone"))
    rsCompany("Fax") = Trim(Request.Form("Fax1"))
    rsCompany("HomePage") = Trim(Request.Form("Homepage1"))
    rsCompany("BankOfDeposit") = Trim(Request.Form("BankOfDeposit"))
    rsCompany("BankAccount") = Trim(Request.Form("BankAccount"))
    rsCompany("TaxNum") = Trim(Request.Form("TaxNum"))
    rsCompany("StatusInField") = PE_CLng(Trim(Request.Form("StatusInField")))
    rsCompany("CompanySize") = PE_CLng(Trim(Request.Form("CompanySize")))
    rsCompany("BusinessScope") = Trim(Request.Form("BusinessScope"))
    rsCompany("AnnualSales") = Trim(Request.Form("AnnualSales"))
    rsCompany("ManagementForms") = PE_CLng(Trim(Request.Form("ManagementForms")))
    rsCompany("RegisteredCapital") = Trim(Request.Form("RegisteredCapital"))
    rsCompany.Update
    rsCompany.Close
    Set rsCompany = Nothing

    Call WriteSuccessMsg("������ҵ���ϳɹ�", "Admin_Company.asp?Action=Show&CompanyID=" & CompanyID)
    
End Sub

%>
