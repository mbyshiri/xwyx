<!--#include file="Start.asp"-->
<html>
<head>
<title>��Ա��¼</title>
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
FONT-FAMILY:����;FONT-SIZE: 9pt;line-height: 130%;
}
a{text-decoration: none;} /* �������»���,��Ϊunderline */
a:link {color: #000000;} /* δ���ʵ����� */
a:visited {color: #333333;} /* �ѷ��ʵ����� */
a:hover{COLOR: #AE0927;} /* ����������� */
a:active {color: #0000ff;} /* ����������� */

-->
</style>
<!--����ʹuserlogin.asp���������ܽṹ���ж� -->
<script language="javascript">
//if(self==top){self.location.href="index.asp";}
</script>
<!--���-->
</head>
<body leftmargin=0 topmargin=0>
<%
'ComeUrl=strInstallDir & "UserLogin.asp"
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
    Response.Write "        <td height='25' align='right'>�û�����</td><td height='25'><input name='UserName' type='text' id='UserName' size='16' maxlength='20' style='width:110px;'></td>" & vbCrLf
    If ShowType = 1 Then
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr>" & vbCrLf
    End If
    Response.Write "        <td height='25' align='right'>��&nbsp;&nbsp;�룺</td><td height='25'><input name='UserPassword' type='password' id='Password' size='16' maxlength='20' style='width:110px;'></td>" & vbCrLf
    If ShowType = 1 Then
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr>" & vbCrLf
    End If
    If EnableCheckCodeOfLogin = True Then
        Response.Write "        <td height='25' align='right'>��֤�룺</td><td height='25'><input name='CheckCode' type='text' id='CheckCode' size='6' maxlength='6'><a href='javascript:refreshimg()' title='�������������ͼƬ'><img id='checkcode' src='inc/checkcode.asp' style='border: 1px solid #ffffff'></a></td>" & vbCrLf
        If ShowType = 1 Then
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr>" & vbCrLf
        End If
    End If
    Response.Write "        <td height='25' colspan='2' align='center'>" & vbCrLf
    Response.Write "            <input type='checkbox' name='CookieDate' value='3'>���õ�¼ <input type='hidden' name='ComeUrl' value='" & ComeUrl & "'>" & vbCrLf
    Response.Write "            <input name='Login' type='submit' id='Login' value=' �� ¼ '>" & vbCrLf
    If ShowType = 1 Then
        Response.Write "        <br><br>" & vbCrLf
    Else
        Response.Write "        </td>" & vbCrLf
        Response.Write "<td height='25'>" & vbCrLf
    End If
    Response.Write "<a href='" & strInstallDir & "Reg/User_Reg.asp' target='_blank'>���û�ע��</a>&nbsp;&nbsp;<a href='" & strInstallDir & "User/User_GetPassword.asp' target='_blank'>�������룿</a></td>" & vbCrLf
    Response.Write "</tr></form></table>" & vbCrLf
    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "function refreshimg(){document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();}" & vbCrLf
    Response.Write "   function CheckLoginForm(){" & vbCrLf
    Response.Write "       if(document.UserLogin.UserName.value==''){" & vbCrLf
    Response.Write "           alert('�������û�����');" & vbCrLf
    Response.Write "           document.UserLogin.UserName.focus();" & vbCrLf
    Response.Write "           return false;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       if(document.UserLogin.Password.value == ''){" & vbCrLf
    Response.Write "           alert('���������룡');" & vbCrLf
    Response.Write "           document.UserLogin.Password.focus();" & vbCrLf
    Response.Write "           return false;" & vbCrLf
    Response.Write "       }" & vbCrLf
    If EnableCheckCodeOfLogin = True Then
        Response.Write "       if(document.UserLogin.CheckCode.value == ''){" & vbCrLf
        Response.Write "           alert('��������֤�룡');" & vbCrLf
        Response.Write "           document.UserLogin.CheckCode.focus();" & vbCrLf
        Response.Write "           return false;" & vbCrLf
        Response.Write "       }" & vbCrLf
    End If
    Response.Write "   }" & vbCrLf
    Response.Write "</script>" & vbCrLf
Else
    Call GetUser(UserName)
    Response.Write "<table  align='center' width='100%' border='0' cellspacing='0' cellpadding='2' ><tr><td>&nbsp;&nbsp;<font color=green><b>" & UserName & "</b></font>��"
    If (Hour(Now) < 6) Then
        Response.Write "<font color=##0066FF>�賿��!</font>"
    ElseIf (Hour(Now) < 9) Then
        Response.Write "<font color=##000099>���Ϻ�!</font>"
    ElseIf (Hour(Now) < 12) Then
        Response.Write "<font color=##FF6699>�����!</font>"
    ElseIf (Hour(Now) < 14) Then
        Response.Write "<font color=##FF6600>�����!</font>"
    ElseIf (Hour(Now) < 17) Then
        Response.Write "<font color=##FF00FF>�����!</font>"
    ElseIf (Hour(Now) < 18) Then
        Response.Write "<font color=##0033FF>�����!</font>"
    Else
        Response.Write "<font color=##ff0000>���Ϻ�!</font>"
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    If ShowType = 1 Then
        Response.Write "�ʽ��� <b><font color=blue>" & Balance & "</font></b> Ԫ" & vbCrLf
    Else
        'ע�����ǵ�¼������ʾ,�ٷ��ĺ���������760�������,������������������Ϣ,���������Զ��塣
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    If ShowType = 1 Then
        Response.Write "������֣� <b><font color=blue>" & UserExp & "</font></b> ��" & vbCrLf
    Else

    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    If ShowType = 1 Then
        Response.Write "����" & PointName & "�� <b><font color=blue>" & UserPoint & "</font></b> " & PointUnit & ""
    Else

    End If
    If UserChargeType > 0 Then
        If ShowType = 1 Then
            Response.Write "<br>&nbsp;&nbsp;"
        Else
            Response.Write "</td><td>"
        End If
        If ValidNum = -1 Then
            Response.Write "ʣ�������� <b><font color=blue>������</font></b>"
        Else
            Response.Write "ʣ�������� <b><font color=blue>" & ValidDays & "</font></b> ��"
        End If
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "��ǩ���£�" & vbCrLf
    If Trim(UnsignedItems & "") = "" Then
        Response.Write " <b><font color=gray>0</font></b> ƪ"
    Else
        Dim UnsignedItemNum, arrUser
        arrUser = Split(UnsignedItems, ",")
        UnsignedItemNum = UBound(arrUser) + 1
        Response.Write " <b><font color=red>" & UnsignedItemNum & "</font></b> ƪ"
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "���Ķ��ţ�" & vbCrLf
    If UnreadMsg > 0 Then
        Response.Write " <b><font color=red>" & UnreadMsg & "</font></b> ��"
    Else
        Response.Write " <b><font color=gray>0</font></b> ��"
    End If
    If ShowType = 1 Then
        Response.Write "<br>&nbsp;&nbsp;"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "��¼������ <b><font color=blue>" & LoginTimes & "</font></b> ��" & vbCrLf
    If ShowType = 1 Then
        Response.Write "<br>"
    Else
        Response.Write "</td><td>"
    End If
    Response.Write "<a href='User/Index.asp' target='ControlPad'>����Ա���ġ�</a> <a href='" & strInstallDir & "User/User_Logout.asp' target='_top'>��ע����¼��</a>"
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