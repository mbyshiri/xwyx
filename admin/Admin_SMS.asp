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
Const PurviewLevel_Others = ""   '����Ȩ��

Dim SendTo
SendTo = Trim(Request("SendTo"))

'������Ա����Ȩ��
If AdminPurview > 1 Then
    Select Case SendTo
    Case "Member"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToMember")
    Case "Contacter"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToContacter")
    Case "Consignee"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToConsignee")
    Case "Other"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToOther")
    End Select
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
        Response.End
    End If
End If
%>
<html>
<head>
<Title>�����ֻ�����</Title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312' />
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
<script language='JavaScript' src='PopCalendar.js'></script>
<script language='JavaScript'>
    PopCalendar = getCalendarInstance()
    PopCalendar.startAt = 0 // 0 - sunday ; 1 - monday
    PopCalendar.showWeekNumber = 0 // 0 - don't show; 1 - show
    PopCalendar.showTime = 0 // 0 - don't show; 1 - show
    PopCalendar.showToday = 0 // 0 - don't show; 1 - show
    PopCalendar.showWeekend = 1 // 0 - don't show; 1 - show
    PopCalendar.showHolidays = 1 // 0 - don't show; 1 - show
    PopCalendar.showSpecialDay = 1 // 0 - don't show, 1 - show
    PopCalendar.selectWeekend = 0 // 0 - don't Select; 1 - Select
    PopCalendar.selectHoliday = 0 // 0 - don't Select; 1 - Select
    PopCalendar.addCarnival = 0 // 0 - don't Add; 1- Add to Holiday
    PopCalendar.addGoodFriday = 0 // 0 - don't Add; 1- Add to Holiday
    PopCalendar.language = 0 // 0 - Chinese; 1 - English
    PopCalendar.defaultFormat = 'yyyy-mm-dd' //Default Format dd-mm-yyyy
    PopCalendar.fixedX = -1 // x position (-1 if to appear below control)
    PopCalendar.fixedY = -1 // y position (-1 if to appear below control)
    PopCalendar.fade = .5 // 0 - don't fade; .1 to 1 - fade (Only IE)
    PopCalendar.shadow = 1 // 0  - don't shadow, 1 - shadow
    PopCalendar.move = 1 // 0  - don't move, 1 - move (Only IE)
    PopCalendar.saveMovePos = 1  // 0  - don't save, 1 - save
    PopCalendar.centuryLimit = 40 // 1940 - 2039
    PopCalendar.initCalendar()
</script>
<script language = 'JavaScript'>
function SelectUser(){
    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=UserList&DefaultValue='+document.myform.inceptUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');
    if (arr != null){
        document.myform.inceptUser.value=arr;
    }
}
function CheckForm(){
  if (document.myform.Content.value==''){
     alert('�������ݲ���Ϊ�գ�');
     return false;
  }
<%if Trim(Request("SendTo"))="Contacter" then%>
    document.myform.Country.value=frm1.document.regionform.Country.value;
    document.myform.Province.value=frm1.document.regionform.Province.value;
    document.myform.City.value=frm1.document.regionform.City.value;
<%end if%>
  return true;
}
function checkLength(){
  myform.lencount.value=myform.Content.value.length;
}
</script>
</head>
<body>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
  <tr class='topbg'>
    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>�� �� �� �� �� ��</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10047' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>
  </tr>
  <tr class='tdbg'>
    <td width='70' height='30'><strong>��������</strong></td>
    <td><a href='Admin_SMS.asp?SendTo=Member'>����Ա���Ͷ���</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Contacter'>����ϵ�˷��Ͷ���</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Consignee'>�������е��ջ��˷��Ͷ���</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Other'>�������˷��Ͷ���</a>    </td>
  </tr>
</table>

<%
Select Case SendTo
Case "Member"
    Call SendToMember
Case "Contacter"
    Call SendToContacter
Case "Consignee"
    Call SendToConsignee
Case "Order"
    Call SendToOrder
Case "Other"
    Call SendToOther
Case Else
    Call SendToMember
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
%>
</body>
</html>

<%
Sub SendToMember()
    Dim InceptType, UserName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>����Ա�����ֻ�����</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>�����ˣ�</td>"
    Response.Write "      <td><table><tr><td>��Ա�飺</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td>�û�����</td><td><input type='text' name='inceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=��<a href='#' onclick=""SelectUser();""><font color='green'>��Ա�б�</font></a>��</font>"
    Response.Write "����û���������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr>"
    Response.Write "<tr><td>ID��Χ��</td><td>��ʼID��<input type='text' name='BeginID'  size='10' value=''>&nbsp;��ֹID��<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>�������ݣ�</td>"
    Response.Write "    <td>�����ڶ���������ʹ�ü���������<br>{$1}���ֻ������С��ͨ����<br>{$2}����ʵ����<br>{$3}���û���<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}�����ã�����һ�����Զ��š�----" & SiteName & vbCrLf & "������ظ��˶��ţ�</textarea>"
    Call ShowCommonHTML("SendToMember")
End Sub

Private Sub SendToOrder()
    Dim InceptType, UserName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>�����ﳵ���ʹߵ�����</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>�����ˣ�</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0'"
    If InceptType = 0 Then Response.Write " checked"
    Response.Write "> ���л�Ա</td><td><font color='blue'>ϵͳֻ���Ѿ���д����ȷ�ֻ������С��ͨ������û����Ͷ���</font></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1'"
    If InceptType = 1 Then Response.Write " checked"
    Response.Write "> ָ����Ա��</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2'"
    If InceptType = 2 Then Response.Write " checked"
    Response.Write "> ָ���û���</td><td><input type='text' name='inceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=��<a href='#' onclick=""SelectUser();""><font color='green'>��Ա�б�</font></a>��</font>"
    Response.Write "����û���������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='3'"
    If InceptType = 3 Then Response.Write " checked"
    Response.Write "> ָ��ID��Χ</td><td>��ʼID��<input type='text' name='BeginID'  size='10' value=''>&nbsp;��ֹID��<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>�������ݣ�</td>"
    Response.Write "    <td>�����ڶ���������ʹ�ü���������<br>{$1}���ֻ������С��ͨ����<br>{$2}���û���<br>{$3}����ʵ����<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$3}���ã�����" & SiteName & vbCrlf & "���ﳵ��Ʒû���ύ�������ύ����������벦�����ǵĵ绰��</textarea>"
    Call ShowCommonHTML("SendToMember")
End Sub

Sub SendToContacter()
    Dim InceptType, TrueName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    TrueName = Trim(Request("TrueName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>����ϵ�˷����ֻ�����</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>�����ˣ�</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0'"
    If InceptType = 0 Then Response.Write " checked"
    Response.Write "> ������ϵ��</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1'"
    If InceptType = 1 Then Response.Write " checked"
    Response.Write "> ָ������</td><td>"
    Response.Write "<iframe name='frm1' id='frm1' src='../Region.asp' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "<input name='Country' id='Country' type='hidden'> <input name='Province' id='Province' type='hidden'> <input name='City' id='City' type='hidden'>" & vbCrLf
    Response.Write "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2'"
    If InceptType = 2 Then Response.Write " checked"
    Response.Write "> ָ����ϵ��</td><td><input type='text' name='TrueName' size='40' value='" & TrueName & "'> �����ϵ������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>�������ݣ�</td>"
    Response.Write "    <td>�����ڶ���������ʹ�ü���������<br>{$1}���ֻ������С��ͨ����<br>{$2}����ʵ����<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}�����ã�����һ�����Զ��š�----" & SiteName & vbCrLf & "������ظ��˶��ţ�</textarea>"
    Call ShowCommonHTML("SendToContacter")
End Sub

Sub SendToConsignee()
    Dim InceptType, OrderFormID
    InceptType = PE_CLng(Trim(Request("InceptType")))
    OrderFormID = Trim(Request("OrderFormID"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>�������е��ջ��˷����ֻ�����</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>�����ˣ�</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0' onclick=""searchform.style.display='none'"""
    If InceptType = 0 Then Response.Write " checked"
    Response.Write "> ���ж����е��ջ���</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1' onclick=""searchform.style.display=''"""
    If InceptType = 1 Then Response.Write " checked"
    Response.Write "> ��ѯ����</td><td>"
    
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='searchform' style='display:none'>"
    Response.Write "<tr class='title' align='center'><td colspan='6'>�� �� �� ѯ �� ��</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>�ɣķ�Χ��</td><td>��ʼ�ɣ�<input type='text' name='BeginID'  size='10' value=''>&nbsp;��ֹ�ɣ�<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>���ڷ�Χ��</td><td>��ʼ����<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.formSearch.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;��������<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.formSearch.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>��Χ��</td><td><input type='text' name='MinMoney'  size='10' value=''> �� <input type='text' name='MaxMoney'  size='10' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>����״̬��</td><td><input type='radio' name='OrderStatus' value='1'>�ȴ�ȷ��&nbsp;&nbsp;<input type='radio' name='OrderStatus' value='2'>�Ѿ�ȷ��&nbsp;&nbsp;<input type='radio' name='OrderStatus' value='4'>�Ѿ�����&nbsp;&nbsp;<input type='radio' name='OrderStatus' value='-1' checked>����״̬</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>����״̬��</td><td><input type='radio' name='PayStatus' value='0'>�ȴ�����&nbsp;&nbsp;<input type='radio' name='PayStatus' value='1'>�Ѹ�����&nbsp;&nbsp;<input type='radio' name='PayStatus' value='2'>�Ѿ�����&nbsp;&nbsp;<input type='radio' name='PayStatus' value='-1' checked>����״̬</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>����״̬��</td><td><input type='radio' name='DeliverStatus' value='1'>������&nbsp;&nbsp;&nbsp;&nbsp;<input type='radio' name='DeliverStatus' value='2'>�ѷ���&nbsp;&nbsp;&nbsp;&nbsp;<input type='radio' name='DeliverStatus' value='3'>��ǩ��&nbsp;&nbsp;&nbsp;&nbsp;<input type='radio' name='DeliverStatus' value='-1' checked>����״̬</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>������ţ�</td><td><input type='text' name='OrderFormNum'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>�ͻ����ƣ�</td><td><input type='text' name='ClientName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>�û�����</td><td><input type='text' name='UserName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>�����̣�</td><td><input type='text' name='AgentName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>�ջ���������</td><td><input type='text' name='ContacterName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>��ϵ��ַ��</td><td><input type='text' name='Address'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>��ϵ�绰��</td><td><input type='text' name='Phone'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>�ֻ��ţ�</td><td><input type='text' name='Mobile'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>��ע���ԣ�</td><td><input type='text' name='Remark'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>��Ʒ���ƣ�</td><td><input type='text' name='ProductName'  size='30' value=''></td></tr>"
    Response.Write "</table>"
    
    
    Response.Write "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2' onclick=""searchform.style.display='none'"""
    If InceptType = 2 Then Response.Write " checked"
    Response.Write "> ָ������ID</td><td><input type='text' name='OrderFormID' size='40' value='" & OrderFormID & "'> �������ID������<font color='#0000FF'>Ӣ�ĵĶ���</font>�ָ�</td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>�������ݣ�</td>"
    Response.Write "    <td>�����ڶ���������ʹ�ü���������<br>{$1}���ֻ������С��ͨ����<br>{$2}���ջ�������<br>{$3}���������<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}�����ã����Ķ����Ѿ�������----" & SiteName & vbCrLf & "������ظ��˶��ţ�</textarea>"
    Call ShowCommonHTML("SendToConsignee")
End Sub

Sub SendToOther()
    Dim InceptType, UserName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>�������˷����ֻ�����</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>�����ˣ�</td>"
    Response.Write "      <td>����ͬʱ����˷��Ͷ��š�ÿһ��Ϊһ���ֻ�����<br>һ���п���ʹ�ö��Ż�ո�ָ������Ϣ���ֱ��Ӧ�����е�{$1} {$2} {$3} ����<br>"
    Response.Write "      <textarea name='Receiver' id='SendNum' cols='50' rows='6'>13800000000,����,2380" & vbCrLf & "13900000000 ���� 3278</textarea></td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>�������ݣ�</td>"
    Response.Write "    <td>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}���㱾�µĹ���Ϊ{$3}�������Ѿ�������������ʻ�����ע����գ�----����" & vbCrLf & "������ظ��˶��ţ�</textarea>"
    Call ShowCommonHTML("SendToOther")
End Sub

Sub ShowCommonHTML(SendTo)
    Response.Write "<br>ÿ70���ּ���Ϊһ�����ŷ���&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ���д��������<INPUT type='text' name='lencount' size='3' value='0' readOnly>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>����ʱ�䣺</td>"
    Response.Write "    <td><input type='radio' name='SendTiming' value='0' checked>��������&nbsp;&nbsp;<input type='radio' name='SendTiming' value='1'>��ʱ"
    Response.Write "<input name='SendDate' type='text' size='10' maxlength='10' value='" & FormatDateTime(Now(), 2) & "' onFocus=""PopCalendar.show(document.myform.SendDate, 'yyyy-mm-dd', null, null, null, '11');""> "
    Response.Write "<input name='SendTime_Hour' type='text' size='2' maxlength='2' value='" & Hour(Now()) & "'>ʱ<input name='SendTime_Minute' type='text' size='2' maxlength='2' value='" & Minute(Now()) & "'>��"
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan=2 align=center>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='Send'><input name='SendTo' type='hidden' id='SendTo' value='" & SendTo & "'>"
    Response.Write "      <input name='Submit' type='submit' id='Submit' value=' �� �� '>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<SCRIPT LANGUAGE='JavaScript'>checkLength();</SCRIPT>"

End Sub

%>
