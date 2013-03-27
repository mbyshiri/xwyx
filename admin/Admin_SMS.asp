<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim SendTo
SendTo = Trim(Request("SendTo"))

'检查管理员操作权限
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
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>对不起，你没有此项操作的权限。</font></p>"
        Response.End
    End If
End If
%>
<html>
<head>
<Title>发送手机短信</Title>
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
     alert('短信内容不能为空！');
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
    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>发 送 手 机 短 信</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10047' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>
  </tr>
  <tr class='tdbg'>
    <td width='70' height='30'><strong>管理导航：</strong></td>
    <td><a href='Admin_SMS.asp?SendTo=Member'>给会员发送短信</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Contacter'>给联系人发送短信</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Consignee'>给订单中的收货人发送短信</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Other'>给其他人发送短信</a>    </td>
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
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>给会员发送手机短信</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>接收人：</td>"
    Response.Write "      <td><table><tr><td>会员组：</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td>用户名：</td><td><input type='text' name='inceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=【<a href='#' onclick=""SelectUser();""><font color='green'>会员列表</font></a>】</font>"
    Response.Write "多个用户名间请用<font color='#0000FF'>英文的逗号</font>分隔</td></tr>"
    Response.Write "<tr><td>ID范围：</td><td>起始ID：<input type='text' name='BeginID'  size='10' value=''>&nbsp;终止ID：<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>短信内容：</td>"
    Response.Write "    <td>可以在短信内容中使用几个变量：<br>{$1}：手机号码或小灵通号码<br>{$2}：真实姓名<br>{$3}：用户名<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}，您好！这是一条测试短信。----" & SiteName & vbCrLf & "（请勿回复此短信）</textarea>"
    Call ShowCommonHTML("SendToMember")
End Sub

Private Sub SendToOrder()
    Dim InceptType, UserName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>给购物车发送催单短信</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>接收人：</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0'"
    If InceptType = 0 Then Response.Write " checked"
    Response.Write "> 所有会员</td><td><font color='blue'>系统只向已经填写了正确手机号码或小灵通号码的用户发送短信</font></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1'"
    If InceptType = 1 Then Response.Write " checked"
    Response.Write "> 指定会员组</td><td>" & GetUserGroup("", "") & "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2'"
    If InceptType = 2 Then Response.Write " checked"
    Response.Write "> 指定用户名</td><td><input type='text' name='inceptUser' size='40' value='" & UserName & "'>"
    Response.Write "<font color='blue'><=【<a href='#' onclick=""SelectUser();""><font color='green'>会员列表</font></a>】</font>"
    Response.Write "多个用户名间请用<font color='#0000FF'>英文的逗号</font>分隔</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='3'"
    If InceptType = 3 Then Response.Write " checked"
    Response.Write "> 指定ID范围</td><td>起始ID：<input type='text' name='BeginID'  size='10' value=''>&nbsp;终止ID：<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>短信内容：</td>"
    Response.Write "    <td>可以在短信内容中使用几个变量：<br>{$1}：手机号码或小灵通号码<br>{$2}：用户名<br>{$3}：真实姓名<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$3}您好！您在" & SiteName & vbCrlf & "购物车商品没有提交，烦请提交，如需帮助请拨打我们的电话。</textarea>"
    Call ShowCommonHTML("SendToMember")
End Sub

Sub SendToContacter()
    Dim InceptType, TrueName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    TrueName = Trim(Request("TrueName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>给联系人发送手机短信</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>接收人：</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0'"
    If InceptType = 0 Then Response.Write " checked"
    Response.Write "> 所有联系人</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1'"
    If InceptType = 1 Then Response.Write " checked"
    Response.Write "> 指定地区</td><td>"
    Response.Write "<iframe name='frm1' id='frm1' src='../Region.asp' width='100%' height='75' frameborder='0' scrolling='no'></iframe>" & vbCrLf
    Response.Write "<input name='Country' id='Country' type='hidden'> <input name='Province' id='Province' type='hidden'> <input name='City' id='City' type='hidden'>" & vbCrLf
    Response.Write "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2'"
    If InceptType = 2 Then Response.Write " checked"
    Response.Write "> 指定联系人</td><td><input type='text' name='TrueName' size='40' value='" & TrueName & "'> 多个联系人请用<font color='#0000FF'>英文的逗号</font>分隔</td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>短信内容：</td>"
    Response.Write "    <td>可以在短信内容中使用几个变量：<br>{$1}：手机号码或小灵通号码<br>{$2}：真实姓名<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}，您好！这是一条测试短信。----" & SiteName & vbCrLf & "（请勿回复此短信）</textarea>"
    Call ShowCommonHTML("SendToContacter")
End Sub

Sub SendToConsignee()
    Dim InceptType, OrderFormID
    InceptType = PE_CLng(Trim(Request("InceptType")))
    OrderFormID = Trim(Request("OrderFormID"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>给订单中的收货人发送手机短信</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>接收人：</td>"
    Response.Write "      <td><table><tr><td><input type='radio' name='InceptType' value='0' onclick=""searchform.style.display='none'"""
    If InceptType = 0 Then Response.Write " checked"
    Response.Write "> 所有订单中的收货人</td><td></td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='1' onclick=""searchform.style.display=''"""
    If InceptType = 1 Then Response.Write " checked"
    Response.Write "> 查询订单</td><td>"
    
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border' id='searchform' style='display:none'>"
    Response.Write "<tr class='title' align='center'><td colspan='6'>订 单 查 询 条 件</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>ＩＤ范围：</td><td>起始ＩＤ<input type='text' name='BeginID'  size='10' value=''>&nbsp;终止ＩＤ<input type='text' name='EndID'  size='10' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>日期范围：</td><td>起始日期<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.formSearch.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;结束日期<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.formSearch.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>金额范围：</td><td><input type='text' name='MinMoney'  size='10' value=''> 至 <input type='text' name='MaxMoney'  size='10' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>订单状态：</td><td><input type='radio' name='OrderStatus' value='1'>等待确认&nbsp;&nbsp;<input type='radio' name='OrderStatus' value='2'>已经确认&nbsp;&nbsp;<input type='radio' name='OrderStatus' value='4'>已经结清&nbsp;&nbsp;<input type='radio' name='OrderStatus' value='-1' checked>所有状态</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>付款状态：</td><td><input type='radio' name='PayStatus' value='0'>等待付款&nbsp;&nbsp;<input type='radio' name='PayStatus' value='1'>已付定金&nbsp;&nbsp;<input type='radio' name='PayStatus' value='2'>已经付清&nbsp;&nbsp;<input type='radio' name='PayStatus' value='-1' checked>所有状态</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>物流状态：</td><td><input type='radio' name='DeliverStatus' value='1'>配送中&nbsp;&nbsp;&nbsp;&nbsp;<input type='radio' name='DeliverStatus' value='2'>已发货&nbsp;&nbsp;&nbsp;&nbsp;<input type='radio' name='DeliverStatus' value='3'>已签收&nbsp;&nbsp;&nbsp;&nbsp;<input type='radio' name='DeliverStatus' value='-1' checked>所有状态</td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>订单编号：</td><td><input type='text' name='OrderFormNum'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>客户名称：</td><td><input type='text' name='ClientName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>用户名：</td><td><input type='text' name='UserName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>代理商：</td><td><input type='text' name='AgentName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>收货人姓名：</td><td><input type='text' name='ContacterName'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>联系地址：</td><td><input type='text' name='Address'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>联系电话：</td><td><input type='text' name='Phone'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>手机号：</td><td><input type='text' name='Mobile'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>备注留言：</td><td><input type='text' name='Remark'  size='30' value=''></td></tr>"
    Response.Write "<tr class='tdbg'><td width='120' class='tdbg5' align='right'>商品名称：</td><td><input type='text' name='ProductName'  size='30' value=''></td></tr>"
    Response.Write "</table>"
    
    
    Response.Write "</td></tr>"
    Response.Write "<tr><td valign='top'><input type='radio' name='InceptType' value='2' onclick=""searchform.style.display='none'"""
    If InceptType = 2 Then Response.Write " checked"
    Response.Write "> 指定订单ID</td><td><input type='text' name='OrderFormID' size='40' value='" & OrderFormID & "'> 多个订单ID间请用<font color='#0000FF'>英文的逗号</font>分隔</td></tr>"
    Response.Write "</table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>短信内容：</td>"
    Response.Write "    <td>可以在短信内容中使用几个变量：<br>{$1}：手机号码或小灵通号码<br>{$2}：收货人姓名<br>{$3}：订单编号<br>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}，您好！您的订单已经××。----" & SiteName & vbCrLf & "（请勿回复此短信）</textarea>"
    Call ShowCommonHTML("SendToConsignee")
End Sub

Sub SendToOther()
    Dim InceptType, UserName
    InceptType = PE_CLng(Trim(Request("InceptType")))
    UserName = Trim(Request("UserName"))
    Response.Write "<form name='myform' method='post' onSubmit='return CheckForm();' action='Admin_SMSPost.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' class='title' colspan=2 align=center><b>给其他人发送手机短信</b></td>"
    Response.Write "  </tr>"
    
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right' class='tdbg5'>接收人：</td>"
    Response.Write "      <td>可以同时向多人发送短信。每一行为一个手机号码<br>一行中可以使用逗号或空格分隔多个信息，分别对应内容中的{$1} {$2} {$3} ……<br>"
    Response.Write "      <textarea name='Receiver' id='SendNum' cols='50' rows='6'>13800000000,张三,2380" & vbCrLf & "13900000000 李四 3278</textarea></td>"
    Response.Write "    </tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>短信内容：</td>"
    Response.Write "    <td>"
    Response.Write "      <textarea name='Content' id='Content' cols='50' rows='8' onpropertychange='checkLength();'>{$2}，你本月的工资为{$3}。工资已经存入你的银行帐户，请注意查收！----财务部" & vbCrLf & "（请勿回复此短信）</textarea>"
    Call ShowCommonHTML("SendToOther")
End Sub

Sub ShowCommonHTML(SendTo)
    Response.Write "<br>每70个字计算为一条短信发送&nbsp;&nbsp;&nbsp;&nbsp;已经填写的字数：<INPUT type='text' name='lencount' size='3' value='0' readOnly>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='right' class='tdbg5'>发送时间：</td>"
    Response.Write "    <td><input type='radio' name='SendTiming' value='0' checked>立即发送&nbsp;&nbsp;<input type='radio' name='SendTiming' value='1'>定时"
    Response.Write "<input name='SendDate' type='text' size='10' maxlength='10' value='" & FormatDateTime(Now(), 2) & "' onFocus=""PopCalendar.show(document.myform.SendDate, 'yyyy-mm-dd', null, null, null, '11');""> "
    Response.Write "<input name='SendTime_Hour' type='text' size='2' maxlength='2' value='" & Hour(Now()) & "'>时<input name='SendTime_Minute' type='text' size='2' maxlength='2' value='" & Minute(Now()) & "'>分"
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan=2 align=center>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='Send'><input name='SendTo' type='hidden' id='SendTo' value='" & SendTo & "'>"
    Response.Write "      <input name='Submit' type='submit' id='Submit' value=' 发 送 '>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<SCRIPT LANGUAGE='JavaScript'>checkLength();</SCRIPT>"

End Sub

%>
