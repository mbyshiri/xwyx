<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../Include/PowerEasy.UserInfo.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub Execute()
    Select Case Action
    Case "RegCompany"
        Call RegCompany
    Case "RegCompany2"
        Call RegCompany2
    Case "SaveReg"
        Call SaveReg
    Case "Modify"
        Call ModifyInfo
    Case "SaveInfo"
        Call SaveInfo
    Case "ModifyPwd"
        Call ModifyPwd
    Case "SavePwd"
        Call SavePwd
    Case "ShowMemberInfo"
        Call ShowMemberInfo
    Case "Join"
        Call JoinCompany
    Case "Exit"
        Conn.Execute ("update PE_User set UserType=0,CompanyID=0,ClientID=0 where UserID=" & UserID & " and UserType>1")
        Response.Redirect "Index.asp"
    Case "DelCompany"
        If UserType > 1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>你不是企业创建者，不能注销企业"
        End If
        If ClientID > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>你已经是客户，不能注销企业"
        End If
        If FoundErr = False Then
            Conn.Execute ("update PE_User set UserType=0,CompanyID=0 where CompanyID=" & CompanyID & "")
            Conn.Execute ("delete from PE_Company where CompanyID=" & CompanyID & "")
            Response.Redirect "Index.asp"
        End If
    Case "Agree", "Reject", "Remove", "AddToAdmin", "RemoveFromAdmin"
        Call MemberManage
    Case Else
        Response.Redirect "Index.asp"
    End Select

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub ModifyInfo()
    Dim rsUser
    Set rsUser = Conn.Execute("select * from PE_User where UserID=" & UserID & "")
    Call PopCalendarInit
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "    if(document.myform.Question.value==''){" & vbCrLf
    Response.Write "        alert('密码提示问题不能为空！');" & vbCrLf
    Response.Write "        document.myform.Question.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(document.myform.Email.value==''){" & vbCrLf
    Response.Write "        alert('电子邮件不能为空！');" & vbCrLf
    Response.Write "        document.myform.Email.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.myform.Country1.value=frm1.document.regionform.Country.value;" & vbCrLf
    Response.Write "    document.myform.Province1.value=frm1.document.regionform.Province.value;" & vbCrLf
    Response.Write "    document.myform.City1.value=frm1.document.regionform.City.value;" & vbCrLf
    If FoundInArr(RegFields_MustFill, "TrueName", ",") = True Then
        Response.Write "    if(document.myform.TrueName.value==''){" & vbCrLf
        Response.Write "        alert('用户名不能为空！');" & vbCrLf
        Response.Write "        document.myform.TrueName.focus();" & vbCrLf
        Response.Write "        return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
    End If
    If rsUser("UserType") = 1 Then
        Response.Write "    document.myform.Country2.value=frm2.document.regionform.Country.value;" & vbCrLf
        Response.Write "    document.myform.Province2.value=frm2.document.regionform.Province.value;" & vbCrLf
        Response.Write "    document.myform.City2.value=frm2.document.regionform.City.value;" & vbCrLf
    End If
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

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

    Response.Write "<br>"
    Response.Write "<form name='myform' id='myform' action='User_Info.asp' method='post' onSubmit='javascript:return CheckSubmit();'>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "        <tr align='center'>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>会员信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>联系信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>个人信息</td>" & vbCrLf
    Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>业务信息</td>" & vbCrLf
    If rsUser("UserType") = 1 Then
        Response.Write "            <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>单位信息</td>" & vbCrLf
    End If
    Response.Write "            <td>&nbsp;</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1' class='border'><tr class='tdbg'><td height='100' valign='top'>" & vbCrLf
    Response.Write "  <table width='95%' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>会 员 组：</td>" & vbCrLf
    Response.Write "        <td width='38%'>" & GroupName & "</td>" & vbCrLf
    Response.Write "        <td width='12%' align='right' class='tdbg5'>会员类别：</td>" & vbCrLf
    Response.Write "        <td width='38%'>"
    If PE_CLng(rsUser("UserType")) > 4 Then
        Response.Write arrUserType(0)
    Else
        Response.Write arrUserType(PE_CLng(rsUser("UserType")))
    End If
    Response.Write "      </td>"
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>用 户 名：</td>" & vbCrLf
    Response.Write "        <td width='38%'>" & UserName & "</td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>用户密码：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='UserPassword' type='password' id='UserPassword' size='20' maxlength='20'> <font color='#FF0000'>不修改请留空</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>提示问题：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='Question' type='text' id='Question' value='" & rsUser("Question") & "'  size='32'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>提示答案：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='Answer' type='text' id='Answer' size='20'> <font color='#FF0000'>不修改请留空</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>电子邮件：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='Email' type='text' id='Email' value='" & rsUser("Email") & "'  size='32' maxlength='255'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>隐私设定：</td>" & vbCrLf
    Response.Write "        <td width='38%'><Input type=radio value=0 name=Privacy"
    If rsUser("Privacy") = 0 Then Response.Write " checked"
    Response.Write ">全部公开 <Input type=radio value=1 name=Privacy"
    If rsUser("Privacy") = 1 Then Response.Write " checked"
    Response.Write ">部分公开 <Input type=radio value=2 name=Privacy"
    If rsUser("Privacy") = 2 Then Response.Write " checked"
    Response.Write ">完全保密</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>头像地址：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='UserFace' type='text' value='" & rsUser("UserFace") & "' size='32' maxlength='255'></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>头像宽度：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='FaceWidth' type='text' value='" & rsUser("FaceWidth") & "' size='6' maxlength='3'> 象素</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' valign='top'>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>签名信息：</td>" & vbCrLf
    Response.Write "        <td width='38%'><textarea name='Sign' cols='35' rows='5' id='Sign'>" & PE_ConvertBR(rsUser("Sign")) & "</textarea></td>" & vbCrLf
    Response.Write "        <td width='12%' class='tdbg5' align='right'>头像高度：</td>" & vbCrLf
    Response.Write "        <td width='38%'><input name='FaceHeight' type='text' id='FaceHeight' value='" & rsUser("FaceHeight") & "' size='6' maxlength='3'> 象素</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Dim arrEducation, arrIncome
    arrEducation = GetArrFromDictionary("PE_Contacter", "Education")
    arrIncome = GetArrFromDictionary("PE_Contacter", "Income")

    Dim rsContacter, sqlContacter
    Dim TrueName, Title, Company, Department, Position, Operation, CompanyAddress
    Dim Country, Province, City, Address, ZipCode
    Dim Mobile, OfficePhone, Homephone, Fax1, PHS
    Dim HomePage, Email1, QQ, ICQ, MSN, Yahoo, UC, Aim
    Dim IDCard, Birthday, NativePlace, Nation, Sex, Marriage, Income, Education, GraduateFrom, Family
    Dim InterestsOfLife, InterestsOfCulture, InterestsOfAmusement, InterestsOfSport, InterestsOfOther

    sqlContacter = "select * from PE_Contacter where ContacterID=" & PE_CLng(rsUser("ContacterID")) & ""
    Set rsContacter = Conn.Execute(sqlContacter)
    If rsContacter.BOF And rsContacter.EOF Then
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
        Fax1 = rsContacter("Fax")
        PHS = rsContacter("PHS")
        HomePage = rsContacter("HomePage")
        Email1 = rsContacter("Email")
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
    Response.Write "                        <td class='tdbg5' align='right' >真实姓名：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='TrueName' type='text' size='35' maxlength='20' value='" & TrueName & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >称谓：</td>" & vbCrLf
    Response.Write "                        <td><input name='Title' type='text' size='35' maxlength='20' value='" & Title & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>通讯地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <iframe name='frm1' id='frm1' src='../Region.asp?Action=Modify&Country=" & Country & "&Province=" & Province & "&City=" & City & "' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "                            <input name='Country1' type='hidden'> <input name='Province1' type='hidden'> <input name='City1' type='hidden'>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td colspan='3'>" & vbCrLf
    Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
    Response.Write "                                    <td><input name='Address1' type='text' size='60' maxlength='255' value='" & Address & "'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                                <tr class='tdbg'>" & vbCrLf
    Response.Write "                                    <td align='right' class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
    Response.Write "                                    <td><input name='ZipCode1' type='text' size='35' maxlength='10' value='" & ZipCode & "'></td>" & vbCrLf
    Response.Write "                                </tr>" & vbCrLf
    Response.Write "                            </table>" & vbCrLf
    Response.Write "                        </td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>办公电话：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='OfficePhone' type='text' size='35' maxlength='30' value='" & OfficePhone & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>住宅电话：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='HomePhone' type='text' size='35' maxlength='30' value='" & Homephone & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >移动电话：</td>" & vbCrLf
    Response.Write "                        <td><input name='Mobile' type='text' size='35' maxlength='30' value='" & Mobile & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >传真号码：</td>" & vbCrLf
    Response.Write "                        <td><input name='Fax1' type='text' size='35' maxlength='30' value='" & Fax1 & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >小灵通：</td>" & vbCrLf
    Response.Write "                        <td><input name='PHS' type='text' size='35' maxlength='30' value='" & PHS & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' ></td>" & vbCrLf
    Response.Write "                        <td></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >个人主页：</td>" & vbCrLf
    Response.Write "                        <td><input name='Homepage1' type='text' size='35' maxlength='255' value='" & HomePage & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Email地址：</td>" & vbCrLf
    Response.Write "                        <td><input name='Email1' type='text' size='35' maxlength='90' value='" & Email1 & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >QQ号码：</td>" & vbCrLf
    Response.Write "                        <td><input name='QQ' type='text' size='35' maxlength='20' value='" & QQ & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >MSN帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='MSN' type='text' size='35' maxlength='90' value='" & MSN & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >ICQ号码：</td>" & vbCrLf
    Response.Write "                        <td><input name='ICQ' type='text' size='35' maxlength='25' value='" & ICQ & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >雅虎通帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='Yahoo' type='text' size='35' maxlength='90' value='" & Yahoo & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >UC帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='UC' type='text' size='35' maxlength='90' value='" & UC & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >Aim帐号：</td>" & vbCrLf
    Response.Write "                        <td><input name='Aim' type='text' size='35' maxlength='90' value='" & Aim & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf
    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>出生日期：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Birthday' type='text' size='35' maxlength='10' value='" & Birthday & "' onFocus=""PopCalendar.show(document.myform.Birthday, 'yyyy-mm-dd', null, null, null, '11');""></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>证件号码：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='IDCard' type='text' size='35' maxlength='20' value='" & IDCard & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >籍贯：</td>" & vbCrLf
    Response.Write "                        <td><input name='NativePlace' type='text' size='35' maxlength='50' value='" & NativePlace & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >民族：</td>" & vbCrLf
    Response.Write "                        <td><input name='Nation' type='text' size='35' maxlength='50' value='" & Nation & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >性别：</td>" & vbCrLf
    Response.Write "                        <td><input name='Sex' type='radio' value='0' "
    If Sex <= 0 Or Sex > 2 Then Response.Write " checked"
    Response.Write ">保密 <input name='Sex' type='radio' value='1'"
    If Sex = 1 Then Response.Write " checked"
    Response.Write ">男 <input name='Sex' type='radio' value='2'"
    If Sex = 2 Then Response.Write " checked"
    Response.Write ">女</td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >婚姻状况：</td>" & vbCrLf
    Response.Write "                        <td><input name='Marriage' type='radio' value='0'"
    If Marriage = 0 Then Response.Write " checked"
    Response.Write ">保密 <input name='Marriage' type='radio' value='1'"
    If Marriage = 1 Then Response.Write " checked"
    Response.Write ">未婚 <input name='Marriage' type='radio' value='2'"
    If Marriage = 2 Then Response.Write " checked"
    Response.Write ">已婚 <input name='Marriage' type='radio' value='3'"
    If Marriage = 3 Then Response.Write " checked"
    Response.Write ">离异</td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >学历：</td>" & vbCrLf
    Response.Write "                        <td><select name='Education'>" & Array2Option(arrEducation, Education) & "</select></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >毕业学校：</td>" & vbCrLf
    Response.Write "                        <td><input name='GraduateFrom' type='text' size='35' maxlength='255' value='" & GraduateFrom & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >生活爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfLife' type='text' size='35' maxlength='255' value='" & InterestsOfLife & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >文化爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfCulture' type='text' size='35' maxlength='255' value='" & InterestsOfCulture & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >娱乐休闲爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfAmusement' type='text' size='35' maxlength='255' value='" & InterestsOfAmusement & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >体育爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfSport' type='text' size='35' maxlength='255' value='" & InterestsOfSport & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >其他爱好：</td>" & vbCrLf
    Response.Write "                        <td><input name='InterestsOfOther' type='text' size='35' maxlength='255' value='" & InterestsOfOther & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >月 收 入：</td>" & vbCrLf
    Response.Write "                        <td><select name='Income'>" & Array2Option(arrIncome, Income) & "</select></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf

    Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>单位名称：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Company' type='text' size='35' maxlength='100' value='" & Company & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right'  width='12%'>所属部门：</td>" & vbCrLf
    Response.Write "                        <td width='38%'><input name='Department' type='text' size='35' maxlength='50' value='" & Department & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >职位：</td>" & vbCrLf
    Response.Write "                        <td><input name='Position' type='text' size='35' maxlength='50' value='" & Position & "'></td>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >负责业务：</td>" & vbCrLf
    Response.Write "                        <td><input name='Operation' type='text' size='35' maxlength='50' value='" & Operation & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                    <tr class='tdbg'>" & vbCrLf
    Response.Write "                        <td class='tdbg5' align='right' >单位地址：</td>" & vbCrLf
    Response.Write "                        <td colspan='3'><input name='CompanyAddress' type='text' size='35' maxlength='200' value='" & CompanyAddress & "'></td>" & vbCrLf
    Response.Write "                    </tr>" & vbCrLf
    Response.Write "                  </tbody>" & vbCrLf

    Dim Company2, Phone, Fax2, Country2, Province2, City2, Address2, ZipCode2, HomePage2
    Dim BankOfDeposit, BankAccount, TaxNum, StatusInField, CompanySize, BusinessScope, AnnualSales, ManagementForms, RegisteredCapital
    Dim CompanyIntro, CompamyPic
    Dim arrStatusInField, arrCompanySize, arrManagementForms
    arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
    arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
    arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")
    If rsUser("UserType") = 1 Then
        Dim rsCompany
        Set rsCompany = Conn.Execute("select * from PE_Company where CompanyID=" & PE_CLng(rsUser("CompanyID")) & "")
        If rsCompany.BOF And rsCompany.EOF Then
            StatusInField = -1
            CompanySize = -1
            ManagementForms = -1
        Else
            Company2 = rsCompany("CompanyName")
            Address2 = rsCompany("Address")
            Country2 = rsCompany("Country")
            Province2 = rsCompany("Province")
            City2 = rsCompany("City")
            ZipCode2 = rsCompany("ZipCode")
            Phone = rsCompany("Phone")
            Fax2 = rsCompany("Fax")
            HomePage2 = rsCompany("Homepage")
            BankOfDeposit = rsCompany("BankOfDeposit")
            BankAccount = rsCompany("BankAccount")
            TaxNum = rsCompany("TaxNum")
            StatusInField = rsCompany("StatusInField")
            CompanySize = rsCompany("CompanySize")
            BusinessScope = rsCompany("BusinessScope")
            AnnualSales = rsCompany("AnnualSales")
            ManagementForms = rsCompany("ManagementForms")
            RegisteredCapital = rsCompany("RegisteredCapital")
            CompanyIntro = rsCompany("CompanyIntro")
            CompamyPic = rsCompany("CompamyPic")
        End If
        rsCompany.Close
        Set rsCompany = Nothing
        Response.Write "                  <tbody id='Tabs' style='display:none'>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>单位名称：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Company2' type='text' size='35' maxlength='100' value='" & Company2 & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'></td>" & vbCrLf
        Response.Write "                        <td width='38%'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>通讯地址：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & vbCrLf
        Response.Write "                            <iframe name='frm2' id='frm2' src='../Region.asp?Action=Modify&Country=" & Country2 & "&Province=" & Province2 & "&City=" & City2 & "' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
        Response.Write "                            <input name='Country2' type='hidden'> <input name='Province2' type='hidden'> <input name='City2' type='hidden'>" & vbCrLf
        Response.Write "                        </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & vbCrLf
        Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
        Response.Write "                                <tr class='tdbg'>" & vbCrLf
        Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
        Response.Write "                                    <td><input name='Address2' type='text' size='60' maxlength='255' value='" & Address2 & "'></td>" & vbCrLf
        Response.Write "                                </tr>" & vbCrLf
        Response.Write "                                <tr class='tdbg'>" & vbCrLf
        Response.Write "                                    <td align='right' class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
        Response.Write "                                    <td><input name='ZipCode2' type='text' size='35' maxlength='10' value='" & ZipCode2 & "'></td>" & vbCrLf
        Response.Write "                                </tr>" & vbCrLf
        Response.Write "                            </table>" & vbCrLf
        Response.Write "                        </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>联系电话：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Phone' type='text' size='35' maxlength='30' value='" & Phone & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>传真号码：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Fax2' type='text' size='35' maxlength='30' value='" & Fax2 & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>开户银行：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='BankOfDeposit' type='text' size='35' maxlength='255' value='" & BankOfDeposit & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>银行帐号：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='BankAccount' type='text' size='35' maxlength='255' value='" & BankAccount & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >税号：</td>" & vbCrLf
        Response.Write "                        <td><input name='TaxNum' type='text' size='35' maxlength='50' value='" & TaxNum & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >网址：</td>" & vbCrLf
        Response.Write "                        <td><input name='Homepage2' type='text' size='35' maxlength='100' value='" & HomePage2 & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >行业地位：</td>" & vbCrLf
        Response.Write "                        <td><select name='StatusInField'>" & Array2Option(arrStatusInField, StatusInField) & "</select></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司规模：</td>" & vbCrLf
        Response.Write "                        <td><select name='CompanySize'>" & Array2Option(arrCompanySize, CompanySize) & "</select></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >业务范围：</td>" & vbCrLf
        Response.Write "                        <td><input name='BusinessScope' type='text' size='35' maxlength='255' value='" & BusinessScope & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >年销售额：</td>" & vbCrLf
        Response.Write "                        <td><input name='AnnualSales' type='text' size='15' maxlength='20' value='" & AnnualSales & "'> 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >经营状态：</td>" & vbCrLf
        Response.Write "                        <td><select name='ManagementForms'>" & Array2Option(arrManagementForms, ManagementForms) & "</select></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >注册资本：</td>" & vbCrLf
        Response.Write "                        <td><input name='RegisteredCapital' type='text' size='15' maxlength='20' value='" & RegisteredCapital & "'> 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司照片：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'><input name='CompamyPic' type='text' size='35' maxlength='255' value='" & CompamyPic & "'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司简介：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'><textarea name='CompanyIntro' cols='75' rows='5' id='CompanyIntro'>" & PE_ConvertBR(CompanyIntro) & "</textarea></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                </tbody>" & vbCrLf
    End If


    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='5' cellspacing='1'><tr align='center'><td height='40'>" & vbCrLf
    Response.Write "    <input type='hidden' name='action' value='SaveInfo'>" & vbCrLf
    Response.Write "    <input type='hidden' name='UserName' value='" & UserName & "'>" & vbNewLine
    Response.Write "    <input type='submit' name='Submit' value='保存修改结果'>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "    <input type='button' name='Cancel' value=' 取 消 ' onclick=""window.location.href='Index.asp'"">" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Set rsUser = Nothing
End Sub

Sub SaveInfo()
    Dim UserPassword, LastPassword, Question, Answer, Email
    Dim sqlUser, rsUser
    Dim UserType, ClientID, ContacterID, CompanyID
    Dim TrueName

    UserPassword = ReplaceBadChar(Trim(Request.Form("UserPassword")))
    Question = ReplaceBadChar(Trim(Request.Form("Question")))
    Answer = ReplaceBadChar(Trim(Request.Form("Answer")))
    Email = ReplaceBadChar(Trim(Request.Form("Email")))
    GroupID = PE_CLng(Trim(Request.Form("GroupID")))
    TrueName = nohtml(Trim(Request.Form("TrueName")))

    If Len(TrueName) > 20 Then
        FoundErr = True
        ErrMsg = "<li>真实姓名,请添入20个以下的字符！</li>"
    End If
    If Question = "" Then
        FoundErr = True
        ErrMsg = "提示问题不能为空！"
    End If
    If Email = "" Then
        FoundErr = True
        ErrMsg = "Email不能为空！"
    End If
    If FoundErr Then
        Exit Sub
    End If
    '加入对整合接口的支持
    If API_Enable Then
        If createXmlDom Then
            sPE_Items(conAction, 1) = "update"
            sPE_Items(conUsername, 1) = UserName
            sPE_Items(conPassword, 1) = UserPassword
            sPE_Items(conEmail, 1) = Email
            sPE_Items(conQuestion, 1) = Question
            sPE_Items(conAnswer, 1) = Answer
            sPE_Items(conUserstatus, 1) = 0
            sPE_Items(conJointime, 1) = Now()
            sPE_Items(conUserip, 1) = UserTrueIP
            sPE_Items(conTruename, 1) = TrueName
            sPE_Items(conGender, 1) = exchangeGender(Trim(Request.Form("Sex")))
            sPE_Items(conBirthday, 1) = FormatDateTime(PE_CDate(Trim(Request.Form("Birthday"))), vbShortDate)
            sPE_Items(conQQ, 1) = nohtml(Trim(Request.Form("QQ")))
            sPE_Items(conMsn, 1) = nohtml(Trim(Request.Form("MSN")))
            sPE_Items(conMobile, 1) = nohtml(Trim(Request.Form("Mobile")))
            sPE_Items(conTelephone, 1) = nohtml(Trim(Request.Form("OfficePhone")))
            sPE_Items(conProvince, 1) = nohtml(Trim(Request.Form("Province1")))
            sPE_Items(conCity, 1) = nohtml(Trim(Request.Form("City1")))
            sPE_Items(conAddress, 1) = nohtml(Trim(Request.Form("Address1")))
            sPE_Items(conZipcode, 1) = nohtml(Trim(Request.Form("ZipCode1")))
            sPE_Items(conHomepage, 1) = nohtml(Trim(Request.Form("HomePage1")))
            prepareXml True
            SendPost
            If FoundErr Then
                ErrMsg = "<li>" & ErrMsg & "</li>"
            End If
        Else
            FoundErr = True
            ErrMsg = "<li>用户服务暂时不可用! [APIError-XmlDom-Runtime]</li>"
        End If
    End If
    If FoundErr Then Exit Sub
    '完毕
    sqlUser = "SELECT * FROM PE_User Where UserID=" & UserID & ""
    Set rsUser = Server.CreateObject("adodb.recordset")
    rsUser.open sqlUser, Conn, 1, 3
    If UserPassword <> "" Then
        rsUser("UserPassword") = MD5(UserPassword, 16)
    End If
    rsUser("Question") = Question
    If Answer <> "" Then
        rsUser("Answer") = MD5(Answer, 16)
    End If
    rsUser("Email") = Email
    rsUser("UserFace") = nohtml(Trim(Request.Form("UserFace")))
    rsUser("FaceWidth") = PE_CLng(Trim(Request.Form("FaceWidth")))
    rsUser("FaceHeight") = PE_CLng(Trim(Request.Form("FaceHeight")))
    rsUser("Sign") = PE_HTMLEncode(Trim(Request.Form("Sign")))
    rsUser("Privacy") = PE_CLng(Trim(Request.Form("Privacy")))
    UserType = PE_CLng(rsUser("UserType"))
    ClientID = PE_CLng(rsUser("ClientID"))
    ContacterID = PE_CLng(rsUser("ContacterID"))
    CompanyID = PE_CLng(rsUser("CompanyID"))
    rsUser.Update
    rsUser.Close
    Set rsUser = Nothing
    If FoundInArr(RegFields_MustFill, "TrueName", ",") = True And TrueName = "" Then
        FoundErr = True
        ErrMsg = "真实姓名不能为空！"
        Exit Sub
    End If
    If ClientID <> 0 Then
        Conn.Execute ("Update PE_Client Set ClientName='" & TrueName & "' Where ClientID=" & ClientID)
    End If

    Dim sqlContacter, rsContacter
    Set rsContacter = Server.CreateObject("adodb.recordset")
    sqlContacter = "select * From PE_Contacter where ContacterID=" & ContacterID & ""
    rsContacter.open sqlContacter, Conn, 1, 3
    If rsContacter.BOF And rsContacter.EOF Then
        ContacterID = GetNewID("PE_Contacter", "ContacterID")
        Conn.Execute ("update PE_User set ContacterID=" & ContacterID & " where UserID=" & UserID & "")
        rsContacter.addnew
        rsContacter("ContacterID") = ContacterID
        rsContacter("ClientID") = ClientID
        rsContacter("ParentID") = 0
        rsContacter("Family") = ""
        rsContacter("CreateTime") = Now()
        rsContacter("Owner") = ""
    End If
    rsContacter("UserType") = UserType
    rsContacter("TrueName") = TrueName
    rsContacter("Country") = nohtml(Trim(Request.Form("Country1")))
    rsContacter("Province") = nohtml(Trim(Request.Form("Province1")))
    rsContacter("City") = nohtml(Trim(Request.Form("City1")))
    rsContacter("ZipCode") = nohtml(Trim(Request.Form("ZipCode1")))
    rsContacter("Address") = nohtml(Trim(Request.Form("Address1")))
    rsContacter("Mobile") = nohtml(Trim(Request.Form("Mobile")))
    rsContacter("OfficePhone") = nohtml(Trim(Request.Form("OfficePhone")))
    rsContacter("HomePhone") = nohtml(Trim(Request.Form("HomePhone")))
    rsContacter("PHS") = nohtml(Trim(Request.Form("PHS")))
    rsContacter("Fax") = nohtml(Trim(Request.Form("Fax1")))
    rsContacter("Homepage") = nohtml(Trim(Request.Form("Homepage1")))
    rsContacter("Email") = nohtml(Trim(Request.Form("Email1")))
    rsContacter("QQ") = nohtml(Trim(Request.Form("QQ")))
    rsContacter("MSN") = nohtml(Trim(Request.Form("MSN")))
    rsContacter("ICQ") = nohtml(Trim(Request.Form("ICQ")))
    rsContacter("Yahoo") = nohtml(Trim(Request.Form("Yahoo")))
    rsContacter("UC") = nohtml(Trim(Request.Form("UC")))
    rsContacter("Aim") = nohtml(Trim(Request.Form("Aim")))
    rsContacter("Company") = nohtml(Trim(Request.Form("Company")))
    rsContacter("CompanyAddress") = nohtml(Trim(Request.Form("CompanyAddress")))
    rsContacter("Department") = nohtml(Trim(Request.Form("Department")))
    rsContacter("Position") = nohtml(Trim(Request.Form("Position")))
    rsContacter("Operation") = nohtml(Trim(Request.Form("Operation")))
    rsContacter("Title") = nohtml(Trim(Request.Form("Title")))
    rsContacter("BirthDay") = PE_CDate(Trim(Request.Form("Birthday")))
    rsContacter("IDCard") = nohtml(Trim(Request.Form("IDCard")))
    rsContacter("NativePlace") = nohtml(Trim(Request.Form("NativePlace")))
    rsContacter("Nation") = nohtml(Trim(Request.Form("Nation")))
    rsContacter("Sex") = PE_CLng(Trim(Request.Form("Sex")))
    rsContacter("Marriage") = PE_CLng(Trim(Request.Form("Marriage")))
    rsContacter("Education") = PE_CLng(Trim(Request.Form("Education")))
    rsContacter("GraduateFrom") = nohtml(Trim(Request.Form("GraduateFrom")))
    rsContacter("InterestsOfLife") = nohtml(Trim(Request.Form("InterestsOfLife")))
    rsContacter("InterestsOfCulture") = nohtml(Trim(Request.Form("InterestsOfCulture")))
    rsContacter("InterestsOfAmusement") = nohtml(Trim(Request.Form("InterestsOfAmusement")))
    rsContacter("InterestsOfSport") = nohtml(Trim(Request.Form("InterestsOfSport")))
    rsContacter("InterestsOfOther") = nohtml(Trim(Request.Form("InterestsOfOther")))
    rsContacter("Income") = PE_CLng(Trim(Request.Form("Income")))
    rsContacter("UpdateTime") = Now()

    rsContacter.Update
    rsContacter.Close
    Set rsContacter = Nothing

    Dim Company2, Address2, ZipCode2
    Company2 = Trim(Request.Form("Company2"))
    Address2 = Trim(Request.Form("Address2"))
    ZipCode2 = Trim(Request.Form("ZipCode2"))

    If UserType = 1 Then
        If Company2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入单位名称！</li>"
        End If
        If Address2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入单位的联系地址！</li>"
        End If
        If ZipCode2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入单位的邮政编码！</li>"
        End If
        If FoundErr = True Then
            Exit Sub
        End If

        Dim sqlCompany, rsCompany
        Set rsCompany = Server.CreateObject("adodb.recordset")
        sqlCompany = "select * From PE_Company where CompanyID=" & CompanyID & ""
        rsCompany.open sqlCompany, Conn, 1, 3
        If rsCompany.BOF And rsCompany.EOF Then
            CompanyID = GetNewID("PE_Company", "CompanyID")
            Conn.Execute ("update PE_User set CompanyID=" & CompanyID & " where UserID=" & UserID & "")
            rsCompany.addnew
            rsCompany("CompanyID") = CompanyID
            rsCompany("ClientID") = 0
        End If

        rsCompany("CompanyName") = nohtml(Trim(Request.Form("Company2")))
        rsCompany("Country") = nohtml(Trim(Request.Form("Country2")))
        rsCompany("Province") = nohtml(Trim(Request.Form("Province2")))
        rsCompany("City") = nohtml(Trim(Request.Form("City2")))
        rsCompany("Address") = nohtml(Trim(Request.Form("Address2")))
        rsCompany("ZipCode") = nohtml(Trim(Request.Form("ZipCode2")))
        rsCompany("Phone") = nohtml(Trim(Request.Form("Phone")))
        rsCompany("Fax") = nohtml(Trim(Request.Form("Fax2")))
        rsCompany("HomePage") = nohtml(Trim(Request.Form("Homepage2")))
        rsCompany("BankOfDeposit") = nohtml(Trim(Request.Form("BankOfDeposit")))
        rsCompany("BankAccount") = nohtml(Trim(Request.Form("BankAccount")))
        rsCompany("TaxNum") = nohtml(Trim(Request.Form("TaxNum")))
        rsCompany("StatusInField") = PE_CLng(Trim(Request.Form("StatusInField")))
        rsCompany("CompanySize") = PE_CLng(Trim(Request.Form("CompanySize")))
        rsCompany("BusinessScope") = nohtml(Trim(Request.Form("BusinessScope")))
        rsCompany("AnnualSales") = nohtml(Trim(Request.Form("AnnualSales")))
        rsCompany("ManagementForms") = PE_CLng(Trim(Request.Form("ManagementForms")))
        rsCompany("RegisteredCapital") = nohtml(Trim(Request.Form("RegisteredCapital")))
        rsCompany("CompanyIntro") = PE_HTMLEncode(Trim(Request.Form("CompanyIntro")))
        rsCompany("CompamyPic") = nohtml(Trim(Request.Form("CompamyPic")))
        rsCompany.Update
        rsCompany.Close
        Set rsCompany = Nothing
    End If

    Call WriteSuccessMsg("修改信息成功！", ComeUrl)
End Sub


Sub ModifyPwd()
    Response.Write "<form name='myform' action='User_Info.asp' method='post'>" & vbCrLf
    Response.Write "  <table width='400' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='title'>" & vbCrLf
    Response.Write "      <td height='22' colSpan='2'><b>修 改 密 码</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>用 户 名：</td>" & vbCrLf
    Response.Write "      <td>" & UserName & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>旧 密 码：</td>" & vbCrLf
    Response.Write "      <td><input name='OldPassword' type='password' maxLength='16' size='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>新 密 码：</td>" & vbCrLf
    Response.Write "      <td> <input name='Password' type='password' maxLength='16' size='30'> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>确认密码：</td>" & vbCrLf
    Response.Write "      <td><input name='PwdConfirm' type='password' maxLength='16' size='30'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='40' colspan='2'>" & vbCrLf
    Response.Write "        <input name='UserName' type='hidden' id='UserName' value='" & UserName & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SavePwd'>" & vbCrLf
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 保 存 '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub SavePwd()
    Dim OldPassword, Password, PwdConfirm
    Dim rsUser, sqlUser
    OldPassword = Trim(Request.Form("OldPassword"))
    Password = Trim(Request.Form("Password"))
    PwdConfirm = Trim(Request.Form("PwdConfirm"))
    If OldPassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入旧密码！</li>"
    Else
        If CheckBadChar(OldPassword) = False Then
            ErrMsg = ErrMsg + "<li>旧密码中含有非法字符</li>"
            FoundErr = True
        End If
    End If
    If Len(Password) > 12 Or Len(Password) < 6 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入新密码(不能大于12小于6)。</li>"
    Else
        If CheckBadChar(Password) = False Then
            ErrMsg = ErrMsg + "<li>新密码中含有非法字符</li>"
            FoundErr = True
        End If
    End If
    If PwdConfirm = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入确认密码！</li>"
    Else
        If PwdConfirm <> Password Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>确认密码与新密码不一致！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    If API_Enable Then
        If createXmlDom Then
            sPE_Items(conAction, 1) = "update"
            sPE_Items(conUsername, 1) = UserName
            sPE_Items(conPassword, 1) = Password
            prepareXml True
            SendPost
            If FoundErr Then
                ErrMsg = "<li>" & ErrMsg & "</li>"
            End If
        Else
            FoundErr = True
            ErrMsg = "<li>用户服务当前不可用。 [APIError-XmlDom-Runtime]</li>"
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID & ""
    rsUser.open sqlUser, Conn, 1, 3
    If MD5(OldPassword, 16) <> rsUser("UserPassword") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你输入的旧密码不对，没有权限修改！</li>"
    Else
        Password = MD5(Password, 16)
        rsUser("UserPassword") = Password
        rsUser.Update
        Response.Cookies(Site_Sn)("UserPassword") = Password
        Response.Write "<br><br>" & vbCrLf
        Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
        Response.Write "  <tr align='center' class='title'><td height='22'><strong>恭喜你！</strong></td></tr>" & vbCrLf
        Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>密码已经修改成功，请记住您的新密码：<font color=red>" & PwdConfirm & "</font></td></tr>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg'><td>"
        If ComeUrl <> "" Then
            Response.Write "<a href='" & ComeUrl & "'>&lt;&lt; 返回上一页</a>"
        Else
            Response.Write "<a href='javascript:window.close();'>【关闭】</a>"
        End If
        Response.Write "</td></tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub RegCompany()
    If CompanyID > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您已经注册了企业！</li>"
        Exit Sub
    End If
    If ContacterID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您必须先填写好详细的联系信息后才能注册企业！</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Info.asp' method='post'>" & vbCrLf
    Response.Write "  <table width='600' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='title'>" & vbCrLf
    Response.Write "      <td height='22' colSpan='2'><b>注册我的企业（第一步）</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='180' align='right' class='tdbg5'>请输入要注册的企业完整名称：</td>" & vbCrLf
    Response.Write "      <td><input name='CompanyName' type='text' maxLength='200' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='40' colspan='2'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='RegCompany2'>" & vbCrLf
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub RegCompany2()
    If CompanyID > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您已经注册了企业！</li>"
        Exit Sub
    End If
    If ContacterID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您必须先填写好详细的联系信息后才能注册企业！</li>"
        Exit Sub
    End If
    Dim CompanyName, rsCompany
    CompanyName = ReplaceBadChar(Trim(Request("CompanyName")))
    If CompanyName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入企业名称！</li>"
        Exit Sub
    Else
        If Len(CompanyName) < 6 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您输入的企业名称太短！请不要恶意注册！</li>"
        ElseIf Len(CompanyName) > 100 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您输入的企业名称太长！请不要恶意注册！</li>"
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function CheckSubmit(){" & vbCrLf
    Response.Write "    document.myform.Country.value=frm1.document.regionform.Country.value;" & vbCrLf
    Response.Write "    document.myform.Province.value=frm1.document.regionform.Province.value;" & vbCrLf
    Response.Write "    document.myform.City.value=frm1.document.regionform.City.value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='title'>" & vbCrLf
    Response.Write "      <td height='22' colSpan='10'><b>注册我的企业（第二步）</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Set rsCompany = Conn.Execute("select top 5 * from PE_Company where CompanyName like '%" & CompanyName & "%'")
    If rsCompany.BOF And rsCompany.EOF Then
        Dim arrStatusInField, arrCompanySize, arrManagementForms
        arrStatusInField = GetArrFromDictionary("PE_Company", "StatusInField")
        arrCompanySize = GetArrFromDictionary("PE_Company", "CompanySize")
        arrManagementForms = GetArrFromDictionary("PE_Company", "ManagementForms")
        Response.Write "<form name='myform' action='User_Info.asp' method='post'>" & vbCrLf
        Response.Write "<tr class='tdbg' height='50'><td colspan='10'>您输入的企业名称还没有其他人注册。赶紧填写详细信息完成注册吧！ 注册成功后，您将成为这家企业的创建人，拥有这家企业的管理权限（如审核批准其他人的申请）。并成为我们的企业会员，享受更多服务！</td></tr>"

        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'>单位名称：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='CompanyName' type='text' size='35' maxlength='200' value='" & CompanyName & "'></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' width='12%'></td>" & vbCrLf
        Response.Write "                        <td width='38%'></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td rowspan='2' class='tdbg5' align='right'  width='12%'>通讯地址：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & vbCrLf
        Response.Write "                            <iframe name='frm' id='frm1' src='../Region.asp?Action=Modify&Country=&Province=&City=' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
        Response.Write "                            <input name='Country' type='hidden'> <input name='Province' type='hidden'> <input name='City' type='hidden'>" & vbCrLf
        Response.Write "                        </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td colspan='3'>" & vbCrLf
        Response.Write "                            <table width='100%'  border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
        Response.Write "                                <tr class='tdbg'>" & vbCrLf
        Response.Write "                                    <td width='12%' align='right' class='tdbg5' align='right' >联系地址：</td>" & vbCrLf
        Response.Write "                                    <td><input name='Address' type='text' size='60' maxlength='255' value=''></td>" & vbCrLf
        Response.Write "                                </tr>" & vbCrLf
        Response.Write "                                <tr class='tdbg'>" & vbCrLf
        Response.Write "                                    <td align='right' class='tdbg5' align='right' >邮政编码：</td>" & vbCrLf
        Response.Write "                                    <td><input name='ZipCode' type='text' size='35' maxlength='10' value=''></td>" & vbCrLf
        Response.Write "                                </tr>" & vbCrLf
        Response.Write "                            </table>" & vbCrLf
        Response.Write "                        </td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>联系电话：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Phone' type='text' size='35' maxlength='30' value=''></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>传真号码：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='Fax' type='text' size='35' maxlength='30' value=''></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>开户银行：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='BankOfDeposit' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right'  width='12%'>银行帐号：</td>" & vbCrLf
        Response.Write "                        <td width='38%'><input name='BankAccount' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >税号：</td>" & vbCrLf
        Response.Write "                        <td><input name='TaxNum' type='text' size='35' maxlength='20' value=''></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >网址：</td>" & vbCrLf
        Response.Write "                        <td><input name='Homepage' type='text' size='35' maxlength='100' value=''></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >行业地位：</td>" & vbCrLf
        Response.Write "                        <td><select name='StatusInField'>" & Array2Option(arrStatusInField, -1) & "</select></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司规模：</td>" & vbCrLf
        Response.Write "                        <td><select name='CompanySize'>" & Array2Option(arrCompanySize, -1) & "</select></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >业务范围：</td>" & vbCrLf
        Response.Write "                        <td><input name='BusinessScope' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >年销售额：</td>" & vbCrLf
        Response.Write "                        <td><input name='AnnualSales' type='text' size='15' maxlength='20' value=''> 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >经营状态：</td>" & vbCrLf
        Response.Write "                        <td><select name='ManagementForms'>" & Array2Option(arrManagementForms, -1) & "</select></td>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >注册资本：</td>" & vbCrLf
        Response.Write "                        <td><input name='RegisteredCapital' type='text' size='15' maxlength='20' value=''> 万元</td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司照片：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'><input name='CompamyPic' type='text' size='35' maxlength='255' value=''></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "                    <tr class='tdbg'>" & vbCrLf
        Response.Write "                        <td class='tdbg5' align='right' >公司简介：</td>" & vbCrLf
        Response.Write "                        <td colspan='3'><textarea name='CompanyIntro' cols='75' rows='5' id='CompanyIntro'></textarea></td>" & vbCrLf
        Response.Write "                    </tr>" & vbCrLf
        Response.Write "<tr class='tdbg' height='50'><td colspan='10' align='center'><input type='submit' name='Join' value='注册此企业' onclick='CheckSubmit()'><input type='hidden' name='Action' value='SaveReg'></td></tr>"
        Response.Write "</form>"
    Else
        Response.Write "<tr class='tdbg' height='50'><td colspan='10'>已经存在与“" & CompanyName & "”相同或相近的企业（详见下表）。您要注册的企业是否就在其中？<br>如果是，请在对应企业下面点击[加入此企业]按钮向此企业创建人发送请求，并等待他的通过。企业创建人在审核您的申请时，可以查看您的有关信息。申请通过后，您将成为我们的企业会员，享受更多服务！ <br>如果不是，请返回上一步输入其他企业名称。</td></tr>"
        Do While Not rsCompany.EOF
            Response.Write "<form name='myform' action='User_Info.asp' method='post'>" & vbCrLf
            Response.Write "                    <tr class='tdbg'>" & vbCrLf
            Response.Write "                        <td class='tdbg5' align='right' width='12%'>单位名称：</td>" & vbCrLf
            Response.Write "                        <td width='38%'>" & rsCompany("CompanyName") & "</td>" & vbCrLf
            Response.Write "                        <td class='tdbg5' align='right' width='12%'>联系地址：</td>" & vbCrLf
            Response.Write "                        <td width='38%'>" & Left(rsCompany("Address"), Len(rsCompany("Address")) - 4) & "******" & "</td>" & vbCrLf
            Response.Write "                    </tr>" & vbCrLf
            Response.Write "                    <tr class='tdbg'>" & vbCrLf
            Response.Write "                        <td class='tdbg5' align='right'>国家/地区：</td>" & vbCrLf
            Response.Write "                        <td>" & rsCompany("Country") & "</td>" & vbCrLf
            Response.Write "                        <td class='tdbg5' align='right'>省/市：</td>" & vbCrLf
            Response.Write "                        <td>" & rsCompany("Province") & "</td>" & vbCrLf
            Response.Write "                    </tr>" & vbCrLf
            Response.Write "                    <tr class='tdbg'>" & vbCrLf
            Response.Write "                        <td class='tdbg5' align='right'>市/县/区：</td>" & vbCrLf
            Response.Write "                        <td>" & rsCompany("City") & "</td>" & vbCrLf
            Response.Write "                        <td class='tdbg5' align='right'>邮政编码：</td>" & vbCrLf
            Response.Write "                        <td>" & rsCompany("ZipCode") & "</td>" & vbCrLf
            Response.Write "                    </tr>" & vbCrLf
            Response.Write "<tr class='tdbg'><td colspan='10' align='center'><input type='submit' name='Join' value='加入此企业'><input type='hidden' name='Action' value='Join'><input type='hidden' name='CompanyName' value='" & rsCompany("CompanyName") & "'><br><br></td></tr>"
            Response.Write "</form>"
            rsCompany.movenext
        Loop
        Response.Write "<tr class='tdbg'><td colspan='10' align='center'><br><br><input type='button' name='Back' value='返回上一步' onclick=""window.location.href='User_Info.asp?Action=RegCompany'""></td></tr>"
    End If
    rsCompany.Close
    Set rsCompany = Nothing
    Response.Write "</table>"
End Sub

Sub JoinCompany()
    If CompanyID > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您已经注册了企业！</li>"
        Exit Sub
    End If
    If ContacterID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您必须先填写好详细的联系信息后才能注册企业！</li>"
        Exit Sub
    End If
    Dim CompanyName, rsCompany
    CompanyName = ReplaceBadChar(Trim(Request("CompanyName")))
    If CompanyName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入企业名称！</li>"
        Exit Sub
    Else
        If Len(CompanyName) < 6 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您输入的企业名称太短！请不要恶意注册！</li>"
        ElseIf Len(CompanyName) > 100 Then
            FoundErr = True
             ErrMsg = ErrMsg & "<li>您输入的企业名称太长！请不要恶意注册！</li>"
        End If
    End If
    Set rsCompany = Conn.Execute("select CompanyID from PE_Company where CompanyName='" & CompanyName & "'")
    If rsCompany.BOF And rsCompany.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您要加入的企业不存在！</li>"
    Else
        Conn.Execute ("update PE_User set UserType=4,CompanyID=" & rsCompany(0) & " where UserID=" & UserID & "")
    End If
    rsCompany.Close
    Set rsCompany = Nothing
    Call WriteSuccessMsg("已经向" & CompanyName & "的企业创建人发送了加入请求！请等待他（她）的审核批准。", ComeUrl)
End Sub

Sub SaveReg()
    If CompanyID > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您已经注册了企业！</li>"
        Exit Sub
    End If
    If ContacterID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您必须先填写好详细的联系信息后才能注册企业！</li>"
        Exit Sub
    End If
    Dim CompanyName, Address, ZipCode
    CompanyName = ReplaceBadChar(Trim(Request("CompanyName")))
    Address = nohtml(Trim(Request.Form("Address")))
    ZipCode = nohtml(Trim(Request.Form("ZipCode")))
    If CompanyName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入企业名称！</li>"
    Else
        If Len(CompanyName) < 6 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您输入的企业名称太短！请不要恶意注册！</li>"
        ElseIf Len(CompanyName) > 100 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您输入的企业名称太长！请不要恶意注册！</li>"
        End If
    End If
    If Len(Address) < 10 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入单位的详细联系地址（不能少于10个字符）！</li>"
    End If
    If ZipCode = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入单位的邮政编码！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    CompanyID = GetNewID("PE_Company", "CompanyID")
    Dim sqlCompany, rsCompany
    Set rsCompany = Server.CreateObject("adodb.recordset")
    sqlCompany = "select top 1 * From PE_Company"
    rsCompany.open sqlCompany, Conn, 1, 3
    rsCompany.addnew
    rsCompany("CompanyID") = CompanyID
    rsCompany("ClientID") = ClientID
    rsCompany("CompanyName") = CompanyName
    rsCompany("Country") = nohtml(Trim(Request.Form("Country")))
    rsCompany("Province") = nohtml(Trim(Request.Form("Province")))
    rsCompany("City") = nohtml(Trim(Request.Form("City")))
    rsCompany("Address") = Address
    rsCompany("ZipCode") = ZipCode
    rsCompany("Phone") = nohtml(Trim(Request.Form("Phone")))
    rsCompany("Fax") = nohtml(Trim(Request.Form("Fax")))
    rsCompany("HomePage") = nohtml(Trim(Request.Form("Homepage")))
    rsCompany("BankOfDeposit") = nohtml(Trim(Request.Form("BankOfDeposit")))
    rsCompany("BankAccount") = nohtml(Trim(Request.Form("BankAccount")))
    rsCompany("TaxNum") = nohtml(Trim(Request.Form("TaxNum")))
    rsCompany("StatusInField") = PE_CLng(Trim(Request.Form("StatusInField")))
    rsCompany("CompanySize") = PE_CLng(Trim(Request.Form("CompanySize")))
    rsCompany("BusinessScope") = nohtml(Trim(Request.Form("BusinessScope")))
    rsCompany("AnnualSales") = nohtml(Trim(Request.Form("AnnualSales")))
    rsCompany("ManagementForms") = PE_CLng(Trim(Request.Form("ManagementForms")))
    rsCompany("RegisteredCapital") = nohtml(Trim(Request.Form("RegisteredCapital")))
    rsCompany("CompanyIntro") = PE_HTMLEncode(Trim(Request.Form("CompanyIntro")))
    rsCompany("CompamyPic") = nohtml(Trim(Request.Form("CompamyPic")))
    rsCompany.Update
    rsCompany.Close
    Set rsCompany = Nothing
    Conn.Execute ("update PE_User set UserType=1,CompanyID=" & CompanyID & " where UserID=" & UserID & "")
    If ClientID > 0 Then
        Conn.Execute ("update PE_Client set ClientName='" & CompanyName & "',ShortedForm='" & Left(CompanyName, 6) & "',ClientType=0 where ClientID=" & ClientID & "")
    End If
    Call WriteSuccessMsg("已经成功注册企业：" & CompanyName & "<br>从现在起，您是这家企业的创建人，拥有这家企业的管理权限（如审核批准其他人的申请）。同时您成为了我们的企业会员，可以享受更多服务！", ComeUrl)
End Sub

Sub ShowMemberInfo()
    If UserType = 4 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你没有管理权限！</li>"
        Exit Sub
    End If
    Dim MemberID, rsMember
    MemberID = PE_CLng(Trim(Request("MemberID")))
    If MemberID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要查看的成员ID！</li>"
        Exit Sub
    End If
    Response.Write "<br>您现在的位置：用户管理 >> 查看企业成员信息<br>"
    Set rsMember = Conn.Execute("select CompanyID from PE_User where UserID=" & MemberID & "")
    If rsMember.BOF And rsMember.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的成员！</li>"
    Else
        If rsMember(0) <> CompanyID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您不能查看非本企业的成员信息！</li>"
        Else
            Call ShowInfo(MemberID, False)
        End If
    End If
    rsMember.Close
    Set rsMember = Nothing
End Sub

Sub MemberManage()
    If UserType > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你没有管理权限！</li>"
        Exit Sub
    End If
    If (Action = "AddToAadmin" Or Action = "RemoveFromAdmin") And UserType > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>你没有管理权限！</li>"
        Exit Sub
    End If
    Dim MemberID, rsMember
    MemberID = PE_CLng(Trim(Request("MemberID")))
    If MemberID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定成员ID！</li>"
        Exit Sub
    End If
    Set rsMember = Conn.Execute("select CompanyID from PE_User where UserID=" & MemberID & "")
    If rsMember.BOF And rsMember.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的成员！</li>"
    Else
        If rsMember(0) <> CompanyID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您不能对非本企业的成员进行操作！</li>"
        Else
            Select Case Action
            Case "Agree"
                Conn.Execute ("update PE_User set UserType=3,ClientID=" & ClientID & " where UserType=4 and UserID=" & MemberID & "")
            Case "Reject", "Remove"
                Conn.Execute ("update PE_User set UserType=0,CompanyID=0,ClientID=0 where UserID=" & MemberID & "")
            Case "AddToAdmin"
                Conn.Execute ("update PE_User set UserType=2 where UserType>2 and UserID=" & MemberID & "")
            Case "RemoveFromAdmin"
                Conn.Execute ("update PE_User set UserType=3 where UserType=2 and UserID=" & MemberID & "")
            End Select
        End If
    End If
    rsMember.Close
    Set rsMember = Nothing
    Response.Redirect "Index.asp"
End Sub

%>
