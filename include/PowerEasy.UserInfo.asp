<%
Sub ShowInfo(iUserID, ShowCompanyInfo)
    Dim rsUser
    Dim arrUserType
    Set rsUser = Conn.Execute("select * from PE_User where UserID=" & iUserID & "")
    If rsUser.bof And rsUser.EOF Then
        Response.Write "<li>�Ҳ���ָ���Ļ�Ա��</li>"
        rsUser.Close
        Set rsUser = Nothing
        Exit Sub
    End If

    arrUserType = Array("���˻�Ա", "��ҵ��Ա�������ߣ�", "��ҵ��Ա������Ա��", "��ҵ��Ա����ͨ��Ա��", "��ҵ��Ա������˳�Ա��")
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<br><table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��Ա��Ϣ</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��ϵ��Ϣ</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>������Ϣ</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>ҵ����Ϣ</td>" & vbCrLf
    If ShowCompanyInfo = True And (rsUser("UserType") <= 2) Then
        Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>��λ��Ϣ</td>" & vbCrLf
        Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>��λ��Ա</td>" & vbCrLf
    End If
    Response.Write "            <td>&nbsp;</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "  <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>�� �� ����</td><td>" & rsUser("UserName") & "</td>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>�����ַ��</td><td width='36%'>" & rsUser("Email") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>��Ա���</td><td width='36%'>" & GroupName & "</td>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>��Ա���</td><td width='36%'>"
    If PE_CLng(rsUser("UserType")) > 4 Then
        Response.Write arrUserType(0)
    Else
        Response.Write arrUserType(PE_CLng(rsUser("UserType")))
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>�ʽ���</td><td width='36%'>" & rsUser("Balance") & "Ԫ</td>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>����" & PointName & "��</td><td width='36%'>" & rsUser("UserPoint") & "" & PointUnit & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>���û��֣�</td><td width='36%'>" & rsUser("UserExp") & "��</td>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>ʣ��������</td><td width='36%'>" & ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime")) & "��</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>��ԱȨ�ޣ�</td><td width='36%'>"
    If rsUser("SpecialPermission") = True Then
        Response.Write "�Զ���"
    Else
        Response.Write "��Ա��Ĭ��"
    End If
    Response.Write "</td>"
    If rsUser("UserType") <= 2 Then
        Response.Write "    <td width='14%' align='right' class='tdbg5'>�����Ա��</td><td width='36%'>"
        Dim trs
        Set trs = Conn.Execute("select count(0) from PE_User where UserType=4 and CompanyID=" & PE_CLng(rsUser("CompanyID")) & "")
        If trs(0) > 0 Then
            Response.Write " <b><font color=red>" & trs(0) & "</font></b> ��"
        Else
            Response.Write " <b><font color=gray>0</font></b> ��"
        End If
        Response.Write "</td>"
    Else
        Response.Write "    <td width='14%' align='right' class='tdbg5'></td><td width='36%'></td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td width='14%' align='right' class='tdbg5'>ע�����ڣ�</td><td width='36%'>" & rsUser("RegTime") & "</td>" & vbCrLf
    Response.Write "    <td width='14%' align='right' class='tdbg5'>�������ڣ�</td><td width='36%'>" & rsUser("JoinTime") & "</td>" & vbCrLf
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>����¼ʱ�䣺</td><td width='36%'>" & rsUser("LastLoginTime") & "</td>" & vbCrLf
    Response.Write "    <td width='14%' align='right' class='tdbg5'>����¼IP��</td><td width='36%'>" & rsUser("LastLoginIP") & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='14%' align='right' class='tdbg5'>���Ķ��ţ�</td><td width='36%'>"

    If rsUser("UnreadMsg") <> "" And PE_CLng(rsUser("UnreadMsg")) > 0 Then
        Response.Write " <b><font color=red>" & rsUser("UnreadMsg") & "</font></b> ��"
    Else
        Response.Write " <b><font color=gray>0</font></b> ��"
    End If
    Response.Write "</td>"
	Response.Write "    <td width='14%' align='right' class='tdbg5'>��ǩ���£�</td><td width='36%'>"
	
	If rsUser("UnsignedItems") <> "" Then
		Dim UnsignedItemNum, arrUser
		arrUser = Split(rsUser("UnsignedItems"), ",")
		UnsignedItemNum = UBound(arrUser) + 1
		Response.Write " <b><font color=red>" & UnsignedItemNum & "</font></b> ƪ"
	Else
		Response.Write " <b><font color=gray>0</font></b> ƪ"
	End If
	Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  </tbody>" & vbCrLf



    Dim rsContacter, sqlContacter
    Dim TrueName, Title, Company, Department, Position, Operation, CompanyAddress
    Dim Country, Province, City, Address, ZipCode
    Dim Mobile, OfficePhone, Homephone, Fax, PHS
    Dim HomePage, Email, QQ, ICQ, MSN, Yahoo, UC, Aim
    Dim IDCard, Birthday, NativePlace, Nation, Sex, Marriage, Income, Education, GraduateFrom, Family
    Dim InterestsOfLife, InterestsOfCulture, InterestsOfAmusement, InterestsOfSport, InterestsOfOther
    Dim arrEducation, arrIncome
    arrEducation = GetArrFromDictionary("PE_Contacter", "Education")
    arrIncome = GetArrFromDictionary("PE_Contacter", "Income")


    sqlContacter = "select * from PE_Contacter where ContacterID=" & rsUser("ContacterID") & ""
    Set rsContacter = Conn.Execute(sqlContacter)
    If rsContacter.bof And rsContacter.EOF Then
        Sex = -1
        Marriage = 0
        Income = -1
    Else
        TrueName = rsContacter("TrueName")
        Title = rsContacter("Title")
        Country = rsContacter("Country")
        Province = rsContacter("Province")
        City = rsContacter("City")
        ZipCode = rsContacter("ZipCode")
        Address = rsContacter("Address")
        OfficePhone = rsContacter("OfficePhone")
        Homephone = rsContacter("HomePhone")
        Mobile = rsContacter("Mobile")
        Fax = rsContacter("Fax")
        PHS = rsContacter("PHS")
        HomePage = rsContacter("HomePage")
        Email = rsContacter("Email")
        QQ = rsContacter("QQ")
        ICQ = rsContacter("ICQ")
        MSN = rsContacter("MSN")
        Yahoo = rsContacter("Yahoo")
        UC = rsContacter("UC")
        Aim = rsContacter("Aim")
        IDCard = rsContacter("IDCard")
        Birthday = rsContacter("Birthday")
        NativePlace = rsContacter("NativePlace")
        Nation = rsContacter("Nation")
        Sex = rsContacter("Sex")
        Marriage = rsContacter("Marriage")
        Income = rsContacter("Income")
        Education = rsContacter("Education")
        GraduateFrom = rsContacter("GraduateFrom")
        InterestsOfLife = rsContacter("InterestsOfLife")
        InterestsOfCulture = rsContacter("InterestsOfCulture")
        InterestsOfAmusement = rsContacter("InterestsOfAmusement")
        InterestsOfSport = rsContacter("InterestsOfSport")
        InterestsOfOther = rsContacter("InterestsOfOther")
        
        Company = rsContacter("Company")
        Department = rsContacter("Department")
        Position = rsContacter("Position")
        Operation = rsContacter("Operation")
        CompanyAddress = rsContacter("CompanyAddress")
    End If
    rsContacter.Close
    Set rsContacter = Nothing
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ʵ������</td>" & vbCrLf
    Response.Write "                        <td>" & TrueName & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ν��</td>" & vbCrLf
    Response.Write "                        <td>" & Title & "</td>" & vbCrLf
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
    Response.Write "                        <td class='tdbg5' align='right' >�칫�绰��</td>" & vbCrLf
    Response.Write "                        <td>" & OfficePhone & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >סլ�绰��</td>" & vbCrLf
    Response.Write "                        <td>" & Homephone & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�ƶ��绰��</td>" & vbCrLf
    Response.Write "                        <td>" & Mobile & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >������룺</td>" & vbCrLf
    Response.Write "                        <td>" & Fax & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >С��ͨ��</td>" & vbCrLf
    Response.Write "                        <td>" & PHS & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' ></td>" & vbCrLf
    Response.Write "                        <td></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >������ҳ��</td>" & vbCrLf
    Response.Write "                        <td>" & HomePage & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Email��ַ��</td>" & vbCrLf
    Response.Write "                        <td>" & Email & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >QQ���룺</td>" & vbCrLf
    Response.Write "                        <td>" & QQ & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >MSN�ʺţ�</td>" & vbCrLf
    Response.Write "                        <td>" & MSN & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ICQ���룺</td>" & vbCrLf
    Response.Write "                        <td>" & ICQ & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�Ż�ͨ�ʺţ�</td>" & vbCrLf
    Response.Write "                        <td>" & Yahoo & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >UC�ʺţ�</td>" & vbCrLf
    Response.Write "                        <td>" & UC & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Aim�ʺţ�</td>" & vbCrLf
    Response.Write "                        <td>" & Aim & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf

    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�������ڣ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Birthday & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>֤�����룺</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & IDCard & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >���᣺</td>" & vbCrLf
    Response.Write "                        <td>" & NativePlace & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >���壺</td>" & vbCrLf
    Response.Write "                        <td>" & Nation & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�Ա�</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(Array("����", "��", "Ů"), Sex) & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >����״����</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(Array("����", "δ��", "�ѻ�", "����"), Marriage) & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ѧ����</td>" & vbCrLf
    Response.Write "                        <td>" & GetArrItem(arrEducation, Education) & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��ҵѧУ��</td>" & vbCrLf
    Response.Write "                        <td>" & GraduateFrom & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >����ã�</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfLife & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�Ļ����ã�</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfCulture & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�������а��ã�</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfAmusement & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�������ã�</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfSport & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�������ã�</td>" & vbCrLf
    Response.Write "                        <td>" & InterestsOfOther & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >�� �� �룺</td>" & vbCrLf
    Response.Write "                        <td>"
    If Income > 6 Then
        Response.Write Income
    Else
        Response.Write GetArrItem(arrIncome, Income)
    End If
    Response.Write "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf
    
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>��λ���ƣ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Company & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�������ţ�</td>" & vbCrLf
    Response.Write "                        <td width='38%'>" & Department & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ְλ��</td>" & vbCrLf
    Response.Write "                        <td>" & Position & "</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >����ҵ��</td>" & vbCrLf
    Response.Write "                        <td>" & Operation & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >��λ��ַ��</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & CompanyAddress & "</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf


    If ShowCompanyInfo = True And (rsUser("UserType") <= 2) Then
        Dim CompanyName, Phone, Fax2, Homepage2
        Dim BankOfDeposit, BankAccount, TaxNum, StatusInField, CompanySize, BusinessScope, AnnualSales, ManagementForms, RegisteredCapital, CompanyIntro
        Dim arrStatusInField, arrCompanySize, arrManagementForms
        arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
        arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
        arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")
        Dim rsCompany
        Set rsCompany = Conn.Execute("select * from PE_Company where CompanyID=" & rsUser("CompanyID") & "")
        If rsCompany.bof And rsCompany.EOF Then
            Country = ""
            Province = ""
            City = ""
            ZipCode = ""
            Address = ""
            StatusInField = -1
            CompanySize = -1
            ManagementForms = -1
        Else
            CompanyName = rsCompany("CompanyName")
            Address = rsCompany("Address")
            Country = rsCompany("Country")
            Province = rsCompany("Province")
            City = rsCompany("City")
            ZipCode = rsCompany("ZipCode")
            Phone = rsCompany("Phone")
            Fax2 = rsCompany("Fax")
            BankOfDeposit = rsCompany("BankOfDeposit")
            BankAccount = rsCompany("BankAccount")
            TaxNum = rsCompany("TaxNum")
            StatusInField = rsCompany("StatusInField")
            CompanySize = rsCompany("CompanySize")
            BusinessScope = rsCompany("BusinessScope")
            AnnualSales = rsCompany("AnnualSales")
            ManagementForms = rsCompany("ManagementForms")
            RegisteredCapital = rsCompany("RegisteredCapital")
            Homepage2 = rsCompany("Homepage")
            CompanyIntro = rsCompany("CompanyIntro")
        End If
        rsCompany.Close
        Set rsCompany = Nothing
        Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>��λ���ƣ�</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & CompanyName & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>��ϵ��ַ��</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & Address & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>����/������</td>" & vbCrLf
        Response.Write "                        <td>" & Country & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>ʡ/�У�</td>" & vbCrLf
        Response.Write "                        <td>" & Province & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>��/��/����</td>" & vbCrLf
        Response.Write "                        <td>" & City & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>�������룺</td>" & vbCrLf
        Response.Write "                        <td>" & ZipCode & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>��ϵ�绰��</td>" & vbCrLf
        Response.Write "                        <td>" & Phone & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'>������룺</td>" & vbCrLf
        Response.Write "                        <td>" & Fax2 & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�������У�</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & BankOfDeposit & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>�����ʺţ�</td>" & vbCrLf
        Response.Write "                        <td width='38%'>" & BankAccount & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >˰�ţ�</td>" & vbCrLf
        Response.Write "                        <td>" & TaxNum & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >��ַ��</td>" & vbCrLf
        Response.Write "                        <td>" & Homepage2 & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >��ҵ��λ��</td>" & vbCrLf
        Response.Write "                        <td>" & GetArrItem(arrStatusInField, StatusInField) & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >��˾��ģ��</td>" & vbCrLf
        Response.Write "                        <td>" & GetArrItem(arrCompanySize, CompanySize) & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >ҵ��Χ��</td>" & vbCrLf
        Response.Write "                        <td>" & BusinessScope & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >�����۶</td>" & vbCrLf
        Response.Write "                        <td>" & AnnualSales & " ��Ԫ</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >��Ӫ״̬��</td>" & vbCrLf
        Response.Write "                        <td>" & GetArrItem(arrManagementForms, ManagementForms) & "</td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >ע���ʱ���</td>" & vbCrLf
        Response.Write "                        <td>" & RegisteredCapital & " ��Ԫ</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >��˾��飺</td>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & CompanyIntro & "</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                </tbody>" & vbCrLf
    End If
    Response.Write "  </table>" & vbCrLf
    
    If ShowCompanyInfo = True And (rsUser("UserType") <= 3) Then
        arrUserType = Array("���˻�Ա", "������", "����Ա", "��ͨ��Ա", "����˳�Ա")
        
        Response.Write "<table id='Tabs' style='display:none' width='100%'><tr class='title' align='center'><td>��Ա��</td><td>��ʵ����</td><td>��������</td><td>��ϵ��ַ</td><td>״̬</td><td>����</td></tr>"
        Dim rsMember
        If rsUser("UserType") <= 2 Then
            Set rsMember = Conn.Execute("select U.UserID,U.UserName,U.UserType,C.TrueName,C.ZipCode,C.Address from PE_User U left join PE_Contacter C on U.ContacterID=C.ContacterID where U.CompanyID=" & rsUser("CompanyID") & " order by U.UserType,U.UserID")
        Else
            Set rsMember = Conn.Execute("select U.UserID,U.UserName,U.UserType,C.TrueName,C.ZipCode,C.Address from PE_User U left join PE_Contacter C on U.ContacterID=C.ContacterID where U.CompanyID=" & rsUser("CompanyID") & " and U.UserType<4 order by U.UserType,U.UserID")
        End If
        Do While Not rsMember.EOF
            Response.Write "<tr><td align='center'><a href='User_Info.asp?Action=ShowMemberInfo&MemberID=" & rsMember("UserID") & "' target='MemberInfo'>" & rsMember("UserName") & "</a></td>"
            Response.Write "<td align='center'><a href='User_Info.asp?Action=ShowMemberInfo&MemberID=" & rsMember("UserID") & "' target='MemberInfo'>" & rsMember("TrueName") & "</a></td>"
            Response.Write "<td align='center'>" & rsMember("ZipCode") & "</td>"
            Response.Write "<td>" & rsMember("Address") & "</td>"
            Response.Write "<td align='center'>"
            If PE_CLng(rsMember("UserType")) > 4 Then
                Response.Write arrUserType(0)
            Else
                Response.Write arrUserType(PE_CLng(rsMember("UserType")))
            End If
            Response.Write "</td>"
            Response.Write "<td align='center' width='150'>"
            Select Case rsMember("UserType")
            Case 4  'δ��˳�Ա
                If rsUser("UserType") <= 2 Then
                    Response.Write "<a href='User_Info.asp?Action=Agree&MemberID=" & rsMember("UserID") & "'>��׼����</a> "
                    Response.Write "<a href='User_Info.asp?Action=Reject&MemberID=" & rsMember("UserID") & "'>�ܾ�����</a>"
                End If
            Case 3  '��ͨ��Ա
                If rsUser("UserType") <= 2 Then
                    Response.Write "<a href='User_Info.asp?Action=Remove&MemberID=" & rsMember("UserID") & "'>����ҵ��ɾ��</a> "
                    If rsUser("UserType") = 1 Then
                        Response.Write "<a href='User_Info.asp?Action=AddToAdmin&MemberID=" & rsMember("UserID") & "'>����Ϊ����Ա</a>"
                    End If
                End If
            Case 2  '����Ա
                If rsUser("UserType") = 1 Then
                    Response.Write "<a href='User_Info.asp?Action=RemoveFromAdmin&MemberID=" & rsMember("UserID") & "'>��Ϊ��ͨ��Ա</a>"
                End If
            End Select

            Response.Write "</td></tr>"
            rsMember.movenext
        Loop
        rsMember.Close
        Set rsMember = Nothing
        Response.Write "</table>"
    End If
    Response.Write "</td></tr></table>" & vbCrLf
    rsUser.Close
    Set rsUser = Nothing

End Sub
%>