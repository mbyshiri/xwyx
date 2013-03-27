<!--#include file="Start.asp"-->
<html>
<head>
<title>会员登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
Input
{
BACKGROUND-COLOR: #ffffff;
BORDER-BOTTOM: #666666 1px solid;
BORDER-LEFT: #666666 1px solid;
BORDER-RIGHT: #666666 1px solid;
BORDER-TOP: #666666 1px solid;
COLOR: #666666;
HEIGHT: 18px;
border-color: #666666 #666666 #666666 #666666; font-size: 9pt
}
TD
{
FONT-FAMILY:宋体;FONT-SIZE: 9pt;line-height: 130%;
}
a{text-decoration: none;} /* 链接无下划线,有为underline */
a:link {color: #000000;} /* 未访问的链接 */
a:visited {color: #333333;} /* 已访问的链接 */
a:hover{COLOR: #AE0927;} /* 鼠标在链接上 */
a:active {color: #0000ff;} /* 点击激活链接 */

-->
</style>
<!--增加使userlogin.asp不能脱离框架结构的判断 -->
<script language="javascript">
//if(self==top){self.location.href="index.asp";}
</script>
<!--完毕-->
</head>
<body leftmargin=0 topmargin=0>
<%
'ComeUrl=strInstallDir & "UserLogin.asp"
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


Dim ShowType, LoginNum
ShowType = Trim(Request.QueryString("ShowType"))
If ShowType = "" Then
    ShowType = 1
Else
    ShowType = CLng(ShowType)
End If

If CheckUserLogined() = False Then
    Response.Write "<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
    Response.Write "    <form action='" & strInstallDir & "User/User_ChkLogin.asp' method='post' name='UserLogin' onSubmit='return CheckLoginForm();' target='_top'>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "        <td height='25' align='right'>用户名：</td><td height='25'><input name='UserName' type='text' id='UserName' size='16' maxlength='20' style='width:110px;'></td>" & vbCrLf
    If ShowType = 1 Then
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr>" & vbCrLf
    End If
    Response.Write "        <td height='25' align='right'>密&nbsp;&nbsp;码：</td><td height='25'><input name='UserPassword' type='password' id='Password' size='16' maxlength='20' style='width:110px;'></td>" & vbCrLf
    If ShowType = 1 Then
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr>" & vbCrLf
    End If
    If EnableCheckCodeOfLogin = True Then
        Response.Write "        <td height='25' align='right'>验证码：</td><td height='25'><input name='CheckCode' type='text' id='CheckCode' size='6' maxlength='6'><a href='javascript:refreshimg()' title='看不清楚，换个图片'><img id='checkcode' src='inc/checkcode.asp' style='border: 1px solid #ffffff'></a></td>" & vbCrLf
        If ShowType = 1 Then
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr>" & vbCrLf
        End If
    End If
    Response.Write "        <td height='25' colspan='2' align='center'>" & vbCrLf
    Response.Write "            <input type='checkbox' name='CookieDate' value='3'>永久登录 <input type='hidden' name='ComeUrl' value='" & ComeUrl & "'>" & vbCrLf
    Response.Write "            <input name='Login' type='submit' id='Login' value=' 登 录 '>" & vbCrLf
    If ShowType = 1 Then
        Response.Write "        <br><br>" & vbCrLf
    Else
        Response.Write "        </td>" & vbCrLf
        Response.Write "<td height='25'>" & vbCrLf
    End If
    Response.Write "<a href='" & strInstallDir & "Reg/User_Reg.asp' target='_blank'>新用户注册</a>&nbsp;&nbsp;<a href='" & strInstallDir & "User/User_GetPassword.asp' target='_blank'>忘记密码？</a></td>" & vbCrLf
    Response.Write "</tr></form></table>" & vbCrLf
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function refreshimg(){document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();}" & vbCrLf
    Response.Write "   function CheckLoginForm(){" & vbCrLf
    Response.Write "       if(document.UserLogin.UserName.value==''){" & vbCrLf
    Response.Write "           alert('请输入用户名！');" & vbCrLf
    Response.Write "           document.UserLogin.UserName.focus();" & vbCrLf
    Response.Write "           return false;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       if(document.UserLogin.Password.value == ''){" & vbCrLf
    Response.Write "           alert('请输入密码！');" & vbCrLf
    Response.Write "           document.UserLogin.Password.focus();" & vbCrLf
    Response.Write "           return false;" & vbCrLf
    Response.Write "       }" & vbCrLf
    If EnableCheckCodeOfLogin = True Then
        Response.Write "       if(document.UserLogin.CheckCode.value == ''){" & vbCrLf
        Response.Write "           alert('请输入验证码！');" & vbCrLf
        Response.Write "           document.UserLogin.CheckCode.focus();" & vbCrLf
        Response.Write "           return false;" & vbCrLf
        Response.Write "       }" & vbCrLf
    End If
    Response.Write "   }" & vbCrLf
    Response.Write "</script>" & vbCrLf
Else
    Call GetUser(UserName)
    Response.Write "<table  align='center' width='100%' border='0' cellspacing='0' cellpadding='2' ><tr><td>&nbsp;&nbsp;<font color=green><b>" & UserName & "</b></font>，"
    If (Hour(Now) < 6) Then
        Response.Write "<font color=##0066FF>凌晨好!</font>"
    ElseIf (Hour(Now) < 9) Then
        Response.Write "<font color=##000099>早上好!</font>"
    ElseIf (Hour(Now) < 12) Then
        Response.Write "<font color=##FF6699>上午好!</font>"
    ElseIf (Hour(Now) < 14) Then
        Response.Write "<font color=##FF6600>中午好!</font>"
    ElseIf (Hour(Now) < 17) Then
        Response.Write "<font color=##FF00FF>下午好!</font>"
    ElseIf (Hour(Now) < 18) Then
        Response.Write "<font color=##0033FF>傍晚好!</font>"
    Else
        Response.Write "<font color=##ff0000>晚上好!</font>"
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    If ShowType = 1 Then
        Response.Write "资金余额： <b><font color=blue>" & Balance & "</font></b> 元" & vbCrLf
    Else
        '注这里是登录横向显示,官方的横向数据在760宽度正好,如果您想调用其它的信息,请在这里自定义。
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    If ShowType = 1 Then
        Response.Write "经验积分： <b><font color=blue>" & UserExp & "</font></b> 分" & vbCrLf
    Else

    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    If ShowType = 1 Then
        Response.Write "可用" & PointName & "： <b><font color=blue>" & UserPoint & "</font></b> " & PointUnit & ""
    Else

    End If
    If UserChargeType > 0 Then
        If ShowType = 1 Then
            Response.Write "<br>&nbsp;&nbsp;"
        Else
            Response.Write "</td><td>"
        End If
        If ValidNum = -1 Then
            Response.Write "剩余天数： <b><font color=blue>无限期</font></b>"
        Else
            Response.Write "剩余天数： <b><font color=blue>" & ValidDays & "</font></b> 天"
        End If
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "待签文章：" & vbCrLf
    If Trim(UnsignedItems & "") = "" Then
        Response.Write " <b><font color=gray>0</font></b> 篇"
    Else
        Dim UnsignedItemNum, arrUser
        arrUser = Split(UnsignedItems, ",")
        UnsignedItemNum = UBound(arrUser) + 1
        Response.Write " <b><font color=red>" & UnsignedItemNum & "</font></b> 篇"
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "待阅短信：" & vbCrLf
    If UnreadMsg > 0 Then
        Response.Write " <b><font color=red>" & UnreadMsg & "</font></b> 条"
    Else
        Response.Write " <b><font color=gray>0</font></b> 条"
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "登录次数： <b><font color=blue>" & LoginTimes & "</font></b> 次" & vbCrLf
    If ShowType = 1 Then
        Response.Write "<br>"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "<a href='User/Index.asp' target='ControlPad'>【会员中心】</a> <a href='" & strInstallDir & "User/User_Logout.asp' target='_top'>【注销登录】</a>"
    Response.Write "</td></tr></table>" & vbCrLf
End If

If UnreadMsg <> "" And CLng(UnreadMsg) > CLng(0) Then
    Dim MessageID, rsMessage
    Set rsMessage = Conn.Execute("select Min(ID) from PE_Message where incept='" & UserName & "'and delR=0 and flag=0 and IsSend=1")
    If IsNull(rsMessage(0)) Then
        MessageID = 0
    Else
        MessageID = rsMessage(0)
    End If
    If MessageID > 0 Then
        Response.Write "<script LANGUAGE='JavaScript'>" & vbCrLf
        Response.Write "var url = 'User/User_Message.asp?Action=ReadMsg&MessageID=" & MessageID & "';" & vbCrLf
        Response.Write "window.open (url, 'newmessage', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')" & vbCrLf
        Response.Write "</script>" & vbCrLf
    End If
    rsMessage.Close
    Set rsMessage = Nothing
End If
Call CloseConn

%>
</body>
</html>